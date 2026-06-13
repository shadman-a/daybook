import {
  CopilotBrief,
  CopilotDay,
  CopilotExtensionStatus,
  CopilotProgress
} from "./types";

const REQUEST_TYPE = "DAYBOOK_COPILOT_REQUEST";
const RESPONSE_TYPE = "DAYBOOK_COPILOT_RESPONSE";
const EVENT_TYPE = "DAYBOOK_COPILOT_EVENT";

type ExtensionAction = "status" | "getDay" | "syncDay" | "exportDay" | "cancel" | "clear";

type BridgeResponse<T> = {
  source: "daybook-extension";
  type: typeof RESPONSE_TYPE;
  requestId: string;
  ok: boolean;
  data?: T;
  error?: string;
};

export async function getCopilotStatus(): Promise<CopilotExtensionStatus> {
  try {
    return await request<CopilotExtensionStatus>("status", undefined, 1_200);
  } catch {
    return { installed: false };
  }
}

export function getCopilotDay(date: string): Promise<CopilotDay> {
  return request("getDay", { date });
}

export function syncCopilotDay(date: string, onProgress: (progress: CopilotProgress) => void): Promise<CopilotDay> {
  return request("syncDay", { date }, 240_000, onProgress);
}

export function exportDayToCopilot(
  payload: {
    date: string;
    exportId: string;
    filename: string;
    markdown: string;
    fingerprint: string;
  },
  onProgress: (progress: CopilotProgress) => void
): Promise<CopilotBrief> {
  return request("exportDay", payload, 300_000, onProgress);
}

export function cancelCopilotOperation(): Promise<void> {
  return request("cancel");
}

export function clearCopilotHistory(): Promise<void> {
  return request("clear");
}

function request<T>(
  action: ExtensionAction,
  payload?: unknown,
  timeout = 12_000,
  onProgress?: (progress: CopilotProgress) => void
): Promise<T> {
  const requestId = crypto.randomUUID();

  return new Promise<T>((resolve, reject) => {
    const timer = window.setTimeout(() => {
      cleanup();
      reject(new Error(action === "status"
        ? "Daybook Copilot extension is not installed."
        : "The Copilot extension did not respond in time."));
    }, timeout);

    function cleanup() {
      window.clearTimeout(timer);
      window.removeEventListener("message", receive);
    }

    function receive(event: MessageEvent) {
      if (event.source !== window || !event.data || event.data.source !== "daybook-extension") return;

      if (event.data.type === EVENT_TYPE && event.data.requestId === requestId) {
        onProgress?.(event.data.progress as CopilotProgress);
        return;
      }

      const response = event.data as BridgeResponse<T>;
      if (response.type !== RESPONSE_TYPE || response.requestId !== requestId) return;

      cleanup();
      if (response.ok && response.data !== undefined) resolve(response.data);
      else if (response.ok) resolve(undefined as T);
      else reject(new Error(response.error || "The Copilot extension could not complete the request."));
    }

    window.addEventListener("message", receive);
    window.postMessage({
      source: "daybook-web",
      type: REQUEST_TYPE,
      requestId,
      action,
      payload
    }, window.location.origin);
  });
}
