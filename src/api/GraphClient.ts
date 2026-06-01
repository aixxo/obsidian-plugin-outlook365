import {CalendarEvent} from '../types';
import {AuthManager} from '../auth/AuthManager';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const SELECT_FIELDS = 'id,subject,start,end,location,bodyPreview,body,attendees,organizer,isOnlineMeeting,onlineMeetingUrl';

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
		const raw = await this.graphGet(url);

		// eslint-disable-next-line @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-explicit-any
		return ((raw as any).value as unknown[]).map((e) => this.mapEvent(e));
	}

	// ── Private ───────────────────────────────────────────────────────────────

	private async graphGet(url: string): Promise<unknown> {
		const token = await this.auth.getAccessToken();
		const response = await fetch(url, {
			headers: {
				Authorization: `Bearer ${token}`,
				'Content-Type': 'application/json',
			},
		});

		if (response.status === 401) {
			// Token may have just expired; one retry after refresh is handled by getAccessToken
			throw new Error('Nicht autorisiert (401). Bitte neu anmelden.');
		}

		if (!response.ok) {
			const text = await response.text();
			throw new Error(`Graph API Fehler (${response.status}): ${text}`);
		}

		return response.json();
	}

	// eslint-disable-next-line @typescript-eslint/no-explicit-any
	private mapEvent(raw: any): CalendarEvent {
		const attendees: string[] = (raw.attendees ?? [])
			// eslint-disable-next-line @typescript-eslint/no-explicit-any
			.map((a: any) => (a.emailAddress?.name as string) ?? (a.emailAddress?.address as string) ?? '')
			.filter(Boolean);

		return {
			id: raw.id as string,
			subject: (raw.subject as string) ?? '(Kein Titel)',
			start: raw.start?.dateTime as string,
			end: raw.end?.dateTime as string,
			location: (raw.location?.displayName as string) ?? '',
			bodyPreview: (raw.bodyPreview as string) ?? '',
			bodyHtml: (raw.body?.content as string) ?? '',
			attendees,
			organizer: (raw.organizer?.emailAddress?.name as string) ?? '',
			isOnlineMeeting: (raw.isOnlineMeeting as boolean) ?? false,
			onlineMeetingUrl: (raw.onlineMeetingUrl as string) ?? '',
		};
	}
}
