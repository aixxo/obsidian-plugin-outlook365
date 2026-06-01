import {App, Notice, PluginSettingTab, Setting} from 'obsidian';
import type OutlookCalendarPlugin from '../main';

export class OutlookSettingsTab extends PluginSettingTab {
	plugin: OutlookCalendarPlugin;

	constructor(app: App, plugin: OutlookCalendarPlugin) {
		super(app, plugin);
		this.plugin = plugin;
	}

	display(): void {
		const {containerEl} = this;
		containerEl.empty();

		containerEl.createEl('h2', {text: 'Outlook Calendar – Einstellungen'});

		// ── Azure App Registration ─────────────────────────────────────────────

		containerEl.createEl('h3', {text: 'Azure App-Registrierung'});

		new Setting(containerEl)
			.setName('Client-ID (Application ID)')
			.setDesc(
				'Die Application (Client) ID deiner Azure AD App-Registrierung. ' +
				'Redirect-URI muss auf "obsidian://outlook-calendar-notes" gesetzt sein (Mobile & Desktop).'
			)
			.addText((text) =>
				text
					.setPlaceholder('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx')
					.setValue(this.plugin.settings.clientId)
					.onChange(async (value) => {
						this.plugin.settings.clientId = value.trim();
						await this.plugin.saveSettings();
					})
			);

		new Setting(containerEl)
			.setName('Tenant-ID')
			.setDesc(
				'Deine Azure AD Tenant-ID oder "common" für Multi-Tenant-Anmeldungen.'
			)
			.addText((text) =>
				text
					.setPlaceholder('common')
					.setValue(this.plugin.settings.tenantId)
					.onChange(async (value) => {
						this.plugin.settings.tenantId = value.trim() || 'common';
						await this.plugin.saveSettings();
					})
			);

		// ── Authentication ─────────────────────────────────────────────────────

		containerEl.createEl('h3', {text: 'Authentifizierung'});

		const authStatus = this.plugin.authManager.isAuthenticated()
			? '✓ Angemeldet'
			: '✗ Nicht angemeldet';

		new Setting(containerEl)
			.setName('Status')
			.setDesc(authStatus)
			.addButton((btn) => {
				if (this.plugin.authManager.isAuthenticated()) {
					btn.setButtonText('Abmelden').setWarning().onClick(async () => {
						await this.plugin.authManager.logout();
						this.display();
					});
				} else {
					btn.setButtonText('Jetzt anmelden').setCta().onClick(async () => {
						if (!this.plugin.settings.clientId) {
							new Notice('Bitte zuerst die Client-ID eintragen.');
							return;
						}
						await this.plugin.authManager.startLogin();
						this.display();
					});
				}
			});

		// ── Folders ────────────────────────────────────────────────────────────

		containerEl.createEl('h3', {text: 'Ordner'});

		new Setting(containerEl)
			.setName('Vorlagen-Ordner')
			.setDesc('Pfad im Vault, der Vorlagen-Dateien enthält (z. B. "Templates").')
			.addText((text) =>
				text
					.setPlaceholder('Templates')
					.setValue(this.plugin.settings.templateFolder)
					.onChange(async (value) => {
						this.plugin.settings.templateFolder = value.trim() || 'Templates';
						await this.plugin.saveSettings();
					})
			);

		new Setting(containerEl)
			.setName('Ausgabe-Ordner')
			.setDesc('Pfad im Vault, in dem erstellte Notizen gespeichert werden (z. B. "Meetings").')
			.addText((text) =>
				text
					.setPlaceholder('Meetings')
					.setValue(this.plugin.settings.outputFolder)
					.onChange(async (value) => {
						this.plugin.settings.outputFolder = value.trim() || 'Meetings';
						await this.plugin.saveSettings();
					})
			);

		// ── Help ───────────────────────────────────────────────────────────────

		containerEl.createEl('h3', {text: 'Einrichtung'});
		const help = containerEl.createEl('ol');
		[
			'Gehe zu portal.azure.com → App-Registrierungen → Neue Registrierung.',
			'Wähle als Redirect-URI: Plattform "Mobile- und Desktopanwendungen" → URI: obsidian://outlook-calendar-notes',
			'Füge die API-Berechtigung Calendars.Read (delegiert) hinzu.',
			'Kopiere die Application (Client) ID oben ein.',
			'Klicke "Jetzt anmelden".',
		].forEach((step) => help.createEl('li', {text: step}));
	}
}
