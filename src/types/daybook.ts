export type DaybookSource = "Teams" | "Meeting" | "Email" | "Sent Email" | "File" | "Copilot";

export type DaybookWarningSource = "Teams" | "Copilot" | "Calendar" | "Mail" | "Files";

export type DaybookWarning = {
  source: DaybookWarningSource;
  message: string;
};

export type DaybookPerson = {
  id?: string;
  displayName?: string;
  email?: string;
};

export type DaybookItem = {
  id: string;
  source: DaybookSource;
  timestamp?: string;
  date?: string;
  endTimestamp?: string;
  title: string;
  preview?: string;
  people: DaybookPerson[];
  conversationId?: string;
  conversationTitle?: string;
  sourceUrl?: string;
  importance?: string;
  generated?: boolean;
  raw: unknown;
};

export type DaybookData = {
  date: string;
  items: DaybookItem[];
  teamsCount: number;
  copilotCount: number;
  meetingCount: number;
  emailCount: number;
  fileCount: number;
  warnings: DaybookWarning[];
};
