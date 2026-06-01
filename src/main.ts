import {Notice, Plugin, requestUrl} from 'obsidian';
import {DEFAULT_SETTINGS} from './settings';
import {OutlookPluginSettings} from './types';
import {AuthManager} from './auth/AuthManager';
import {GraphClient} from './api/GraphClient';
import {EventSelectorModal} from './ui/EventSelectorModal';
import {OutlookSettingsTab} from './ui/SettingsTab';

export default class OutlookCalendarPlugin extends Plugin {
	settings: OutlookPluginSettings;
	authManager: AuthManager;
	graphClient: GraphClient;

	async onload(): Promise<void> {
		await this.loadSettings();

		this.authManager = new AuthManager(this);
		this.graphClient = new GraphClient(this.authManager);

		// OAuth 2.0 callback – triggered when Microsoft redirects back to obsidian://
		this.registerObsidianProtocolHandler('outlook-calendar-notes', async (params) => {
			await this.authManager.handleCallback(params as Record<string, string>);
		});

		// Ribbon icon – opens the event picker (or prompts login if not authenticated)
		this.addRibbonIcon('calendar', 'Outlook-Termin in Notiz umwandeln', async () => {
			if (!this.authManager.isAuthenticated()) {
				new Notice('Outlook Calendar: Bitte zuerst in den Einstellungen anmelden.');
				return;
			}
			new EventSelectorModal(this.app, this).open();
		});

		// Command palette entry
		this.addCommand({
			id: 'outlook-to-note',
			name: 'Termin aus Outlook in Notiz umwandeln',
			callback: () => {
				if (!this.authManager.isAuthenticated()) {
					new Notice('Outlook Calendar: Bitte zuerst in den Einstellungen anmelden.');
					return;
				}
				new EventSelectorModal(this.app, this).open();
			},
		});

		// ── Debug command ─────────────────────────────────────────────────────
		// Opens Obsidian Developer Console (Ctrl+Shift+I / Cmd+Opt+I) to see results.
		this.addCommand({
			id: 'outlook-debug',
			name: 'Outlook Calendar: Debug-Info in Console ausgeben',
			callback: async () => {
				console.debug('Outlook Calendar – Debug start');
				try {
					// 1. Token
					const token = await this.authManager.getAccessToken();
					console.debug('✓ Access Token vorhanden (erste 20 Zeichen):', token.substring(0, 20) + '…');

					// 2. /me – Anmeldestatus prüfen
					const meRes = await requestUrl({
						url: 'https://graph.microsoft.com/v1.0/me',
						headers: {Authorization: `Bearer ${token}`},
						throw: false,
					});
					const me = meRes.json as Record<string, unknown>;
					console.debug('✓ /me Response:', me);
					const displayName = (me.displayName ?? me.userPrincipalName ?? '?') as string;

					// 3. calendarView ohne zusätzliche Parameter (nackte Abfrage)
					const now = new Date();
					const s = new Date(now); s.setDate(s.getDate() - 7); s.setHours(0, 0, 0, 0);
					const e = new Date(now); e.setDate(e.getDate() + 7); e.setHours(23, 59, 59, 999);
					const debugUrl =
						`https://graph.microsoft.com/v1.0/me/calendarView` +
						`?startDateTime=${encodeURIComponent(s.toISOString())}` +
						`&endDateTime=${encodeURIComponent(e.toISOString())}` +
						`&$top=5`;
					console.debug('→ calendarView URL:', debugUrl);

					const calRes = await requestUrl({
						url: debugUrl,
						headers: {Authorization: `Bearer ${token}`},
						throw: false,
					});
					const cal = calRes.json as Record<string, unknown>;
					console.debug('✓ calendarView Response:', cal);

					const count = Array.isArray(cal.value) ? cal.value.length : '(value nicht vorhanden)';
					new Notice(
						`Outlook Debug:\nAngemeldet als: ${displayName}\nTermine (±7 Tage): ${count}\nDetails in Console (Ctrl+Shift+I)`
					);
				} catch (err) {
					console.error('✗ Fehler:', err);
					new Notice(`Debug-Fehler: ${err instanceof Error ? err.message : String(err)}`);
				}
				console.debug('Outlook Calendar – Debug end');
			},
		});

		this.addSettingTab(new OutlookSettingsTab(this.app, this));
	}

	onunload(): void {
		// All registered listeners are cleaned up automatically by Obsidian.
	}

	async loadSettings(): Promise<void> {
		this.settings = Object.assign(
{},
DEFAULT_SETTINGS,
await this.loadData() as Partial<OutlookPluginSettings>
		);
	}

	async saveSettings(): Promise<void> {
		await this.saveData(this.settings);
	}
}
