const REQUEST_TYPE = "DAYBOOK_COPILOT_REQUEST";
const RESPONSE_TYPE = "DAYBOOK_COPILOT_RESPONSE";
const EVENT_TYPE = "DAYBOOK_COPILOT_EVENT";
const allowedOrigin = /^(?:http:\/\/(?:localhost|127\.0\.0\.1):(?:5173|4173)|https:\/\/shadman-a\.github\.io)$/;

window.addEventListener("message", (event) => {
  if (event.source !== window || event.origin !== window.location.origin || !allowedOrigin.test(event.origin)) return;
  if (event.origin === "https://shadman-a.github.io" && !window.location.pathname.startsWith("/daybook")) return;
  const message = event.data;
  if (!message || message.source !== "daybook-web" || message.type !== REQUEST_TYPE) return;

  chrome.runtime.sendMessage({
    type: REQUEST_TYPE,
    requestId: message.requestId,
    action: message.action,
    payload: message.payload
  }).then((response) => {
    window.postMessage({
      source: "daybook-extension",
      type: RESPONSE_TYPE,
      requestId: message.requestId,
      ...response
    }, window.location.origin);
  }).catch((error) => {
    window.postMessage({
      source: "daybook-extension",
      type: RESPONSE_TYPE,
      requestId: message.requestId,
      ok: false,
      error: error instanceof Error ? error.message : "The extension bridge failed."
    }, window.location.origin);
  });
});

chrome.runtime.onMessage.addListener((message) => {
  if (!message || message.type !== EVENT_TYPE) return;
  window.postMessage({
    source: "daybook-extension",
    type: EVENT_TYPE,
    requestId: message.requestId,
    progress: message.progress
  }, window.location.origin);
});
