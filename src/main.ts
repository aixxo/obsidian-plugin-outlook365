import {Notice, Plugin} from 'obsidian';
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
