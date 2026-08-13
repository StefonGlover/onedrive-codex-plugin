const defaultMaxGlobal = 1;
const defaultMaxPerSubject = 1;
const defaultRetryAfterSeconds = 1;
const allowedKinds = new Set(["buffer", "office", "renderer"]);
const genuineLeases = new WeakSet();
const leaseStates = new WeakMap();

function boundedPositiveInteger(value, label, maximum) {
  if (!Number.isSafeInteger(value) || value < 1 || value > maximum) {
    throw new Error(`${label} must be an integer from 1 through ${maximum}.`);
  }
  return value;
}

function boundedSubject(value) {
  if (
    typeof value !== "string"
    || value.length < 1
    || value.length > 256
    || value !== value.trim()
    || /[\u0000-\u001f\u007f]/u.test(value)
  ) {
    throw new Error("A bounded authentication subject is required for heavyweight-process admission.");
  }
  return value;
}

function boundedKind(value) {
  if (!allowedKinds.has(value)) throw new Error("Heavyweight-operation kind must be buffer, office, or renderer.");
  return value;
}

export function assertReusableHeavyweightSubprocessLease(candidate) {
  const state = candidate && typeof candidate === "object" ? leaseStates.get(candidate) : null;
  if (!state || !genuineLeases.has(candidate) || state.released) {
    throw new Error("A current internal heavyweight-operation lease is required.");
  }
  return candidate;
}

export function createHeavyweightSubprocessAdmissionController({
  maxGlobal = defaultMaxGlobal,
  maxPerSubject = defaultMaxPerSubject,
  retryAfterSeconds = defaultRetryAfterSeconds
} = {}) {
  boundedPositiveInteger(maxGlobal, "maxGlobal", 4);
  boundedPositiveInteger(maxPerSubject, "maxPerSubject", maxGlobal);
  boundedPositiveInteger(retryAfterSeconds, "retryAfterSeconds", 60);

  let activeGlobal = 0;
  const activeBySubject = new Map();

  function acquire({ subject: rawSubject, kind: rawKind } = {}) {
    const subject = boundedSubject(rawSubject);
    const kind = boundedKind(rawKind);
    const activeForSubject = activeBySubject.get(subject) || 0;
    if (activeGlobal >= maxGlobal || activeForSubject >= maxPerSubject) {
      return Object.freeze({
        admitted: false,
        retryable: true,
        retryAfterSeconds,
        code: "HEAVYWEIGHT_SUBPROCESS_BUSY"
      });
    }

    activeGlobal += 1;
    activeBySubject.set(subject, activeForSubject + 1);
    let completionBound = false;
    const state = { released: false };

    function release() {
      if (state.released) return false;
      state.released = true;
      activeGlobal = Math.max(0, activeGlobal - 1);
      const remaining = (activeBySubject.get(subject) || 1) - 1;
      if (remaining > 0) activeBySubject.set(subject, remaining);
      else activeBySubject.delete(subject);
      return true;
    }

    function releaseOnChildCompletion(child) {
      if (completionBound) throw new Error("A heavyweight-process lease may be bound to only one child.");
      if (!child || typeof child.once !== "function") throw new Error("A child-process event emitter is required.");
      completionBound = true;
      child.once("close", release);
      child.once("error", release);
    }

    const lease = Object.freeze({
      admitted: true,
      retryable: false,
      kind,
      release,
      releaseOnChildCompletion
    });
    genuineLeases.add(lease);
    leaseStates.set(lease, state);
    return lease;
  }

  return Object.freeze({ acquire });
}

export function heavyweightSubprocessBusyError(rejection) {
  const retryAfterSeconds = boundedPositiveInteger(
    rejection?.retryAfterSeconds || defaultRetryAfterSeconds,
    "retryAfterSeconds",
    60
  );
  const error = new Error("A document-processing operation is already in progress. Retry shortly.");
  error.code = "HEAVYWEIGHT_SUBPROCESS_BUSY";
  error.retryable = true;
  error.retryAfterSeconds = retryAfterSeconds;
  return error;
}

export const heavyweightSubprocessAdmission = createHeavyweightSubprocessAdmissionController();

export const heavyweightSubprocessAdmissionDefaults = Object.freeze({
  maxGlobal: defaultMaxGlobal,
  maxPerSubject: defaultMaxPerSubject,
  retryAfterSeconds: defaultRetryAfterSeconds
});
