export type ManualTokenIdentity = {
  name: string;
  username: string;
  expiresAt?: Date;
};

type TokenClaims = {
  name?: string;
  preferred_username?: string;
  upn?: string;
  exp?: number;
};

export function normalizeAccessToken(value: string): string {
  return value.trim().replace(/^Bearer\s+/i, "");
}

export function getManualTokenIdentity(accessToken: string): ManualTokenIdentity {
  const claims = readTokenClaims(accessToken);
  const username = claims?.preferred_username || claims?.upn || "Graph Explorer token";

  return {
    name: claims?.name || "Graph Explorer session",
    username,
    expiresAt: claims?.exp ? new Date(claims.exp * 1000) : undefined
  };
}

export function isExpiredAccessToken(accessToken: string): boolean {
  const expiresAt = getManualTokenIdentity(accessToken).expiresAt;
  return expiresAt !== undefined && expiresAt.getTime() <= Date.now();
}

function readTokenClaims(accessToken: string): TokenClaims | undefined {
  const payload = accessToken.split(".")[1];
  if (!payload) return undefined;

  try {
    const normalized = payload.replace(/-/g, "+").replace(/_/g, "/");
    const padded = normalized.padEnd(Math.ceil(normalized.length / 4) * 4, "=");
    return JSON.parse(atob(padded)) as TokenClaims;
  } catch {
    return undefined;
  }
}
