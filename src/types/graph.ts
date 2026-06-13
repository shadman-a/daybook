export type GraphCollection<T> = {
  value?: T[];
  "@odata.nextLink"?: string;
};

export type GraphPerson = {
  id?: string;
  displayName?: string;
  userIdentityType?: string;
};

export type GraphIdentitySet = {
  user?: GraphPerson;
  application?: GraphPerson;
  device?: GraphPerson;
};

export type GraphChat = {
  id: string;
  topic?: string;
  chatType?: string;
  lastUpdatedDateTime?: string;
  webUrl?: string;
};

export type GraphChatMessage = {
  id: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  deletedDateTime?: string;
  messageType?: string;
  importance?: string;
  subject?: string;
  body?: { content?: string; contentType?: string };
  from?: GraphIdentitySet;
  webUrl?: string;
};

export type GraphEmailAddress = {
  name?: string;
  address?: string;
};

export type GraphRecipient = {
  emailAddress?: GraphEmailAddress;
};

export type GraphEvent = {
  id: string;
  subject?: string;
  start?: { dateTime?: string; timeZone?: string };
  end?: { dateTime?: string; timeZone?: string };
  organizer?: GraphRecipient;
  attendees?: Array<{ emailAddress?: GraphEmailAddress }>;
  bodyPreview?: string;
  webLink?: string;
  onlineMeetingUrl?: string;
  onlineMeeting?: { joinUrl?: string };
};

export type GraphMailMessage = {
  id: string;
  subject?: string;
  from?: GraphRecipient;
  toRecipients?: GraphRecipient[];
  receivedDateTime?: string;
  sentDateTime?: string;
  bodyPreview?: string;
  importance?: string;
  webLink?: string;
};

export type GraphDriveItem = {
  id?: string;
  name?: string;
  webUrl?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  createdBy?: GraphIdentitySet;
  lastModifiedBy?: GraphIdentitySet;
};

export type GraphUsedInsight = {
  id?: string;
  lastUsed?: { lastAccessedDateTime?: string };
  resource?: GraphDriveItem;
};

export type GraphFileActivity = GraphDriveItem | GraphUsedInsight;
