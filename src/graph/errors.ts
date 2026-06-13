type GraphErrorLike = {
  code?: string;
  message?: string;
  statusCode?: number;
  status?: number;
};

export function getErrorMessage(error: unknown, fallback: string): string {
  const graphError = error as GraphErrorLike | null;
  const status = graphError?.statusCode ?? graphError?.status;

  if (status === 401) return "Your Microsoft session expired. Sign in again and retry.";
  if (status === 403) return "Microsoft Graph denied access. Check delegated permission consent and retry.";
  if (status === 429) return "Microsoft Graph is throttling requests. Wait a moment and retry.";
  if (typeof graphError?.message === "string" && graphError.message.trim()) return graphError.message;
  if (error instanceof Error && error.message.trim()) return error.message;
  return fallback;
}
