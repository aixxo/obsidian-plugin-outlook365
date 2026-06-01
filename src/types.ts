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

export type DateRange = 'combined' | 'past7' | 'past30' | 'next7' | 'next30';

export interface OutlookPluginSettings {
	clientId: string;
	tenantId: string;
	templateFolder: string;
	outputFolder: string;
	defaultDateRange: DateRange;
	defaultTemplate: string;   // vault path of default template, '' = built-in
	openAfterCreate: boolean;
	onDuplicate: 'skip' | 'suffix'; // 'skip' = show notice, 'suffix' = append HHmmss
	tokenData: TokenData | null;
}
