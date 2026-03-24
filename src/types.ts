export interface MailRecipient {
  emailAddress: {
    address: string;
  };
}

export interface SendMailPayload {
  message: {
    subject: string;
    body: {
      contentType: "HTML" | "Text";
      content: string;
    };
    toRecipients: MailRecipient[];
    ccRecipients?: MailRecipient[];
  };
}

export interface ReplyMailPayload {
  comment: string;
}

export interface CalendarEventPayload {
  subject: string;
  body?: {
    contentType: "HTML" | "Text";
    content: string;
  };
  start: {
    dateTime: string;
    timeZone: string;
  };
  end: {
    dateTime: string;
    timeZone: string;
  };
  location?: {
    displayName: string;
  };
  attendees?: Array<{
    emailAddress: { address: string };
    type: "required" | "optional";
  }>;
  isOnlineMeeting?: boolean;
  onlineMeetingProvider?: "teamsForBusiness";
}

export interface CalendarEventUpdatePayload {
  subject?: string;
  body?: {
    contentType: "HTML" | "Text";
    content: string;
  };
  start?: {
    dateTime: string;
    timeZone: string;
  };
  end?: {
    dateTime: string;
    timeZone: string;
  };
  location?: {
    displayName: string;
  };
  attendees?: Array<{
    emailAddress: { address: string };
    type: "required" | "optional";
  }>;
  isOnlineMeeting?: boolean;
  onlineMeetingProvider?: "teamsForBusiness";
}

export interface GraphErrorResponse {
  error?: {
    code?: string;
    message?: string;
  };
}

// ── Teams types ─────────────────────────────────────────

export interface TeamsChatMessagePayload {
  body: {
    contentType: "html" | "text";
    content: string;
  };
}

export interface TeamsChannelMessagePayload {
  body: {
    contentType: "html" | "text";
    content: string;
  };
  subject?: string;
}
