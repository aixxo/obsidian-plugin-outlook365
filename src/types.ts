export interface CalendarEvent {
	id: string;
	subject: string;
	start: string;       // ISO 8601
	end: string;         // ISO 8601
	location: string;
	bodyPreview: string;
	bodyHtml: string;
	attendees: string[]; // display names
	organizer: string;
	isOnlineMeeting: boolean;
	onlineMeetingUrl: string;
}

export interface TokenData {
	accessToken: string;
	refreshToken: string;
	expiresAt: number;   // Unix timestamp (ms)
	scope: string;
}

export interface OutlookPluginSettings {
	clientId: string;
	tenantId: string;
	templateFolder: string;
	outputFolder: string;
	tokenData: TokenData | null;
}
