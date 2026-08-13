const defaultMaxGlobal = 2;
const defaultMaxPerSubject = 1;
const defaultRetryAfterSeconds = 1;

function positiveBoundedInteger(value, label, maximum) {
  if (!Number.isSafeInteger(value) || value < 1 || value > maximum) {
    throw new Error(`${label} must be an integer from 1 through ${maximum}.`);
  }
  return value;
}

function subjectKey(value) {
  if (
    typeof value !== "string"
    || value.length < 1
    || value.length > 256
    || value !== value.trim()
    || /[\u0000-\u001f\u007f]/u.test(value)
  ) {
    throw new Error("A bounded authentication-subject key is required for resource-read admission.");
  }
  return value;
}

function retryableRejection(retryAfterSeconds) {
  return Object.freeze({
    admitted: false,
    retryable: true,
    retryAfterSeconds,
    code: "RESOURCE_READ_BUSY"
  });
}

export function holdResourceReadUntilResponseDeadline({ lease, response, deadlineMs } = {}) {
  if (!lease?.admitted || typeof lease.holdUntilResponseComplete !== "function" || typeof lease.release !== "function") {
    throw new Error("An admitted resource-read lease is required for the response deadline.");
  }
  if (!response || typeof response.once !== "function") {
    throw new Error("A response event emitter is required for the resource-read response deadline.");
  }
  if (!Number.isSafeInteger(deadlineMs) || deadlineMs < 1 || deadlineMs > 300_000) {
    throw new Error("The resource-read response deadline must be an integer from 1 through 300000 milliseconds.");
  }

  lease.holdUntilResponseComplete(response);
  let settled = false;
  const complete = () => {
    if (settled) return;
    settled = true;
    clearTimeout(timer);
    response.removeListener?.("finish", complete);
    response.removeListener?.("close", complete);
    response.removeListener?.("error", complete);
  };
  const timer = setTimeout(() => {
    if (settled) return;
    settled = true;
    lease.release();
    response.removeListener?.("finish", complete);
    response.removeListener?.("close", complete);
    response.removeListener?.("error", complete);
    // Once a large JSON response has begun, a valid replacement error cannot
    // be sent. Closing the response/socket is the only fail-closed boundary.
    if (typeof response.destroy === "function") response.destroy();
    else if (typeof response.socket?.destroy === "function") response.socket.destroy();
  }, deadlineMs);
  timer.unref?.();
  response.once("finish", complete);
  response.once("close", complete);
  response.once("error", complete);
  if (response.writableFinished === true || response.destroyed === true) complete();
  return Object.freeze({ cancel: complete });
}

/**
 * Bound concurrent large MCP resource responses globally and per authenticated
 * subject. A lease stays live until its response emits finish, close, or error;
 * response.end() and drain/backpressure alone do not release it.
 */
export function createResourceReadAdmissionController({
  maxGlobal = defaultMaxGlobal,
  maxPerSubject = defaultMaxPerSubject,
  retryAfterSeconds = defaultRetryAfterSeconds
} = {}) {
  positiveBoundedInteger(maxGlobal, "maxGlobal", 8);
  positiveBoundedInteger(maxPerSubject, "maxPerSubject", maxGlobal);
  positiveBoundedInteger(retryAfterSeconds, "retryAfterSeconds", 60);

  let activeGlobal = 0;
  const activeBySubject = new Map();

  function acquire(rawSubject) {
    const subject = subjectKey(rawSubject);
    const activeForSubject = activeBySubject.get(subject) || 0;
    if (activeGlobal >= maxGlobal || activeForSubject >= maxPerSubject) {
      return retryableRejection(retryAfterSeconds);
    }

    activeGlobal += 1;
    activeBySubject.set(subject, activeForSubject + 1);
    let released = false;
    let bound = false;

    function release() {
      if (released) return false;
      released = true;
      activeGlobal = Math.max(0, activeGlobal - 1);
      const remaining = (activeBySubject.get(subject) || 1) - 1;
      if (remaining > 0) activeBySubject.set(subject, remaining);
      else activeBySubject.delete(subject);
      return true;
    }

    function holdUntilResponseComplete(response) {
      if (bound) throw new Error("A resource-read admission lease may be bound to only one response.");
      if (!response || typeof response.once !== "function") {
        throw new Error("A response event emitter is required to hold resource-read admission.");
      }
      bound = true;
      if (response.writableFinished === true || response.destroyed === true) {
        release();
        return;
      }
      response.once("finish", release);
      response.once("close", release);
      response.once("error", release);
    }

    return Object.freeze({
      admitted: true,
      retryable: false,
      release,
      holdUntilResponseComplete
    });
  }

  return Object.freeze({ acquire });
}

export const resourceReadAdmissionDefaults = Object.freeze({
  maxGlobal: defaultMaxGlobal,
  maxPerSubject: defaultMaxPerSubject,
  retryAfterSeconds: defaultRetryAfterSeconds
});
