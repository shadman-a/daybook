export type CopilotRole = "user" | "assistant";

export type CopilotRecord = {
  id: string;
  conversationId: string;
  conversationTitle: string;
  role: CopilotRole;
  text: string;
  date: string;
  timestamp?: string;
  sourceUrl?: string;
  links?: Array<{ title: string; url: string }>;
  generated?: boolean;
  exportId?: string;
};

export type CopilotBrief = {
  id: string;
  date: string;
  exportId: string;
  fingerprint: string;
  markdown: string;
  conversationId?: string;
  conversationUrl?: string;
  createdAt: string;
};

export type CopilotDay = {
  records: CopilotRecord[];
  brief?: CopilotBrief;
  lastSyncedAt?: string;
};

export type CopilotProgress = {
  stage: "opening" | "discovering" | "syncing" | "uploading" | "waiting" | "importing" | "complete";
  message: string;
  current?: number;
  total?: number;
};

export type CopilotExtensionStatus = {
  installed: boolean;
  version?: string;
};
