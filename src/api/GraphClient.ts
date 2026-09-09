import {requestUrl} from 'obsidian';
import {CalendarEvent} from '../types';
import {AuthManager} from '../auth/AuthManager';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const SELECT_FIELDS = 'id,subject,start,end,location,bodyPreview,body,attendees,organizer,isOnlineMeeting,onlineMeetingUrl';

// ── Raw Graph API response types ─────────────────────────────────────────────

interface RawAttendee {
	emailAddress?: { name?: string; address?: string };
}

interface RawEvent {
	id: string;
	subject?: string;
	start?: { dateTime: string };
	end?: { dateTime: string };
	location?: { displayName?: string };
	bodyPreview?: string;
	body?: { content?: string };
	attendees?: RawAttendee[];
	organizer?: { emailAddress?: { name?: string } };
	isOnlineMeeting?: boolean;
	onlineMeetingUrl?: string;
}

interface RawCalendarViewResponse {
	value: RawEvent[];
}

// Graph's default (no `Prefer: outlook.timezone` header) `dateTime` values are naive UTC
// clock strings without a `Z`/offset suffix, which `Date` would otherwise parse as local time.
export function normalizeGraphDateTime(dateTime: string): string {
	if (!dateTime) return dateTime;
	return /[Zz]|[+-]\d{2}:?\d{2}$/.test(dateTime) ? dateTime : `${dateTime}Z`;
}

export class GraphClient {
	private auth: AuthManager;

	constructor(auth: AuthManager) {
		this.auth = auth;
	}

	/**
	 * Fetches calendar events within [now - daysBack, now + daysAhead].
	 * Uses /me/calendarView to correctly expand recurring events.
	 */
	async fetchEvents(daysBack = 0, daysAhead = 7): Promise<CalendarEvent[]> {
		const now = new Date();
		const start = new Date(now);
		start.setDate(start.getDate() - daysBack);
		start.setHours(0, 0, 0, 0);

		const end = new Date(now);
		end.setDate(end.getDate() + daysAhead);
		end.setHours(23, 59, 59, 999);

		// Build URL manually – URLSearchParams encodes '$' as '%24' which can
		// break OData query parameters on some tenant configurations.
		const url =
			`${GRAPH_BASE}/me/calendarView` +
			`?startDateTime=${encodeURIComponent(start.toISOString())}` +
			`&endDateTime=${encodeURIComponent(end.toISOString())}` +
			`&$select=${SELECT_FIELDS}` +
			`&$top=100` +
			`&$orderby=start/dateTime`;
		const raw = await this.graphGet(url) as RawCalendarViewResponse;
		return raw.value.map((e) => this.mapEvent(e));
	}

	// ── Private ───────────────────────────────────────────────────────────────

	private async graphGet(url: string): Promise<unknown> {
		const token = await this.auth.getAccessToken();
		const response = await requestUrl({
			url,
			headers: {
				Authorization: `Bearer ${token}`,
				'Content-Type': 'application/json',
			},
			throw: false,
		});

		if (response.status === 401) {
			throw new Error('Nicht autorisiert (401). Bitte neu anmelden.');
		}

		if (response.status < 200 || response.status >= 300) {
			throw new Error(`Graph API Fehler (${response.status}): ${response.text}`);
		}

		return response.json as unknown;
	}

	private mapEvent(raw: RawEvent): CalendarEvent {
		const attendees: string[] = (raw.attendees ?? [])
			.map((a) => a.emailAddress?.name ?? a.emailAddress?.address ?? '')
			.filter(Boolean);

		return {
			id: raw.id,
			subject: raw.subject ?? '(Kein Titel)',
			start: normalizeGraphDateTime(raw.start?.dateTime ?? ''),
			end: normalizeGraphDateTime(raw.end?.dateTime ?? ''),
			location: raw.location?.displayName ?? '',
			bodyPreview: raw.bodyPreview ?? '',
			bodyHtml: raw.body?.content ?? '',
			attendees,
			organizer: raw.organizer?.emailAddress?.name ?? '',
			isOnlineMeeting: raw.isOnlineMeeting ?? false,
			onlineMeetingUrl: raw.onlineMeetingUrl ?? '',
		};
	}
}
