#!/usr/bin/env node

import assert from "node:assert/strict";
import { EventEmitter } from "node:events";
import {
  createResourceReadAdmissionController,
  holdResourceReadUntilResponseDeadline,
  resourceReadAdmissionDefaults
} from "../mcp/resource-read-admission.mjs";

let checks = 0;
function check(value, message) {
  assert.ok(value, message);
  checks += 1;
}

class ResponseProbe extends EventEmitter {
  writableFinished = false;
  destroyed = false;

  destroy() {
    if (this.destroyed) return;
    this.destroyed = true;
    this.emit("close");
  }
}

const admission = createResourceReadAdmissionController({
  maxGlobal: 2,
  maxPerSubject: 1,
  retryAfterSeconds: 2
});

const first = admission.acquire("subject-a");
check(first.admitted, "the first subject should be admitted");
const firstResponse = new ResponseProbe();
first.holdUntilResponseComplete(firstResponse);

const sameSubjectBusy = admission.acquire("subject-a");
check(
  !sameSubjectBusy.admitted
    && sameSubjectBusy.retryable
    && sameSubjectBusy.retryAfterSeconds === 2
    && sameSubjectBusy.code === "RESOURCE_READ_BUSY",
  "the per-subject limit should return a bounded retryable rejection"
);

const second = admission.acquire("subject-b");
check(second.admitted, "a second subject should fit under the global limit");
const secondResponse = new ResponseProbe();
second.holdUntilResponseComplete(secondResponse);

const globalBusy = admission.acquire("subject-c");
check(!globalBusy.admitted && globalBusy.retryable, "the global limit should reject a third large response");
check(Buffer.byteLength(JSON.stringify(globalBusy)) < 512, "the retryable rejection should remain small and bounded");

firstResponse.emit("drain");
check(!admission.acquire("subject-a").admitted, "drain/backpressure relief must not release a response lease");
firstResponse.emit("finish");
const afterFinish = admission.acquire("subject-a");
check(afterFinish.admitted, "finish should release the response lease");
afterFinish.release();
check(!afterFinish.release(), "lease release should be idempotent");

secondResponse.emit("close");
const afterClose = admission.acquire("subject-c");
check(afterClose.admitted, "connection close should release the response lease");
afterClose.release();

const errored = admission.acquire("subject-error");
const errorResponse = new ResponseProbe();
errored.holdUntilResponseComplete(errorResponse);
errorResponse.emit("error", new Error("simulated response failure"));
const afterError = admission.acquire("subject-error");
check(afterError.admitted, "response errors should release the lease");
afterError.release();

assert.throws(() => admission.acquire("bad\nsubject"), /bounded authentication-subject/i);
assert.throws(() => createResourceReadAdmissionController({ maxGlobal: 2, maxPerSubject: 3 }), /maxPerSubject/i);
checks += 2;

check(
  resourceReadAdmissionDefaults.maxGlobal === 2
    && resourceReadAdmissionDefaults.maxPerSubject === 1,
  "production defaults should bound global and per-subject resource reads"
);

const deadlineAdmission = createResourceReadAdmissionController({ maxGlobal: 1, maxPerSubject: 1 });
const stalled = deadlineAdmission.acquire("deadline-subject");
const stalledResponse = new ResponseProbe();
holdResourceReadUntilResponseDeadline({ lease: stalled, response: stalledResponse, deadlineMs: 20 });
await new Promise((resolve) => setTimeout(resolve, 40));
check(stalledResponse.destroyed, "a stalled resource response should be destroyed at its absolute deadline");
const admittedAfterDeadline = deadlineAdmission.acquire("deadline-subject");
check(admittedAfterDeadline.admitted, "deadline closure should release admission for a subsequent resource response");
const successfulResponse = new ResponseProbe();
holdResourceReadUntilResponseDeadline({ lease: admittedAfterDeadline, response: successfulResponse, deadlineMs: 100 });
successfulResponse.writableFinished = true;
successfulResponse.emit("finish");
await new Promise((resolve) => setTimeout(resolve, 120));
check(!successfulResponse.destroyed, "a completed resource response should cancel its absolute deadline");
const admittedAfterSuccess = deadlineAdmission.acquire("deadline-subject");
check(admittedAfterSuccess.admitted, "a successful resource response should release admission normally");
admittedAfterSuccess.release();

console.log(JSON.stringify({ ok: true, checks }));
