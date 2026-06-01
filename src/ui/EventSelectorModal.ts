import {App, Modal, Notice, Setting, TFile} from 'obsidian';
import type OutlookCalendarPlugin from '../main';
import {CalendarEvent, DateRange} from '../types';
import {listTemplates, fillTemplate} from '../notes/TemplateParser';
import {generateNotes} from '../notes/NoteGenerator';

const DATE_RANGE_LABELS: Record<DateRange, string> = {
	combined: 'Letzte & nächste 7 Tage',
	past7:    'Letzte 7 Tage',
	past30:   'Letzte 30 Tage',
	next7:    'Nächste 7 Tage',
	next30:   'Nächste 30 Tage',
};

const DATE_RANGE_PARAMS: Record<DateRange, {daysBack: number; daysAhead: number}> = {
	combined: {daysBack: 7, daysAhead: 7},
	past7:    {daysBack: 7, daysAhead: 0},
	past30:   {daysBack: 30, daysAhead: 0},
	next7:    {daysBack: 0, daysAhead: 7},
	next30:   {daysBack: 0, daysAhead: 30},
};

export class EventSelectorModal extends Modal {
	private plugin: OutlookCalendarPlugin;
	private events: CalendarEvent[] = [];
	private templates: TFile[] = [];
	private selected: Set<string> = new Set();
	private selectedTemplate: TFile | null = null;
	private dateRange: DateRange;
	private loading = false;
	private errorMessage: string | null = null;

	constructor(app: App, plugin: OutlookCalendarPlugin) {
		super(app);
		this.plugin = plugin;
		this.dateRange = plugin.settings.defaultDateRange;
		// selectedTemplate is resolved once templates are loaded in loadData()
	}

	async onOpen(): Promise<void> {
		this.titleEl.setText('Outlook-Termin in Notiz umwandeln');
		await this.renderControls();
		await this.loadData();
	}

	onClose(): void {
		this.contentEl.empty();
	}

	// ── Render ────────────────────────────────────────────────────────────────

	private async renderControls(): Promise<void> {
		const {contentEl} = this;
		contentEl.empty();

		// Date-range selector
		new Setting(contentEl)
			.setName('Zeitraum')
			.addDropdown((dd) => {
				for (const [key, label] of Object.entries(DATE_RANGE_LABELS)) {
					dd.addOption(key, label);
				}
				dd.setValue(this.dateRange);
				dd.onChange((val) => {
					this.dateRange = val as DateRange;
					void this.loadData();
				});
			});

		// Template selector
		new Setting(contentEl)
			.setName('Vorlage')
			.addDropdown((dd) => {
				dd.addOption('', '– Eingebauter Standard –');
				for (const t of this.templates) {
					dd.addOption(t.path, t.basename);
				}
				if (this.templates.length === 0) {
					// keep only the built-in option; no extra placeholder needed
				}
				// Set current value (may have been pre-selected from settings)
				dd.setValue(this.selectedTemplate?.path ?? '');
				dd.onChange((path) => {
					this.selectedTemplate = this.templates.find((t) => t.path === path) ?? null;
				});
			});

		// Event list placeholder
		contentEl.createEl('div', {attr: {id: 'ocn-event-list'}});

		// Action button
		const btnRow = contentEl.createEl('div', {cls: 'modal-button-container'});
		const btn = btnRow.createEl('button', {
			text: 'Notizen erstellen',
			cls: 'mod-cta',
		});
		btn.addEventListener('click', () => { void this.onCreate(); });
	}

	private renderEventList(): void {
		const container = this.contentEl.querySelector('#ocn-event-list');
		if (!container) return;
		container.empty();

		if (this.loading) {
			container.createEl('p', {text: 'Lade Termine…', cls: 'ocn-status'});
			return;
		}

		if (this.errorMessage) {
			const errEl = container.createEl('div', {cls: 'ocn-error'});
			errEl.createEl('strong', {text: 'Fehler: '});
			errEl.createEl('span', {text: this.errorMessage});
			return;
		}

		if (this.events.length === 0) {
			container.createEl('p', {
				text: 'Keine Termine im gewählten Zeitraum gefunden.',
				cls: 'ocn-status',
			});
			return;
		}

		const list = container.createEl('div', {cls: 'ocn-event-list'});

		// Gruppe nach Datum (YYYY-MM-DD)
		const groups = new Map<string, CalendarEvent[]>();
		for (const event of this.events) {
			const dayKey = event.start.substring(0, 10);
			if (!groups.has(dayKey)) groups.set(dayKey, []);
			groups.get(dayKey)!.push(event);
		}

		for (const [dayKey, dayEvents] of groups) {
			// Tag-Header
			list.createEl('div', {
				text: formatDayHeader(dayKey),
				cls: 'ocn-day-header',
			});

			for (const event of dayEvents) {
				const row = list.createEl('div', {cls: 'ocn-event-row'});
				const cb = row.createEl('input', {type: 'checkbox'});
				cb.checked = this.selected.has(event.id);
				cb.addEventListener('change', () => {
					if (cb.checked) {
						this.selected.add(event.id);
					} else {
						this.selected.delete(event.id);
					}
				});

				const label = row.createEl('label');
				label.createEl('strong', {text: event.subject});
				label.createEl('span', {
					text: `  ${formatTime(event.start)} – ${formatTime(event.end)}`,
					cls: 'ocn-event-time',
				});
				if (event.location) {
					label.createEl('span', {text: ` · ${event.location}`, cls: 'ocn-event-location'});
				}
			}
		}
	}

	// ── Data loading ──────────────────────────────────────────────────────────

	private async loadData(): Promise<void> {
		this.loading = true;
		this.errorMessage = null;
		this.renderEventList();

		const {daysBack, daysAhead} = DATE_RANGE_PARAMS[this.dateRange];

		try {
			[this.events, this.templates] = await Promise.all([
				this.plugin.graphClient.fetchEvents(daysBack, daysAhead),
				listTemplates(this.app.vault, this.plugin.settings.templateFolder),
			]);
			// Pre-select default template from settings (only on first load)
			if (this.selectedTemplate === null && this.plugin.settings.defaultTemplate) {
				this.selectedTemplate = this.templates.find(
					(t) => t.path === this.plugin.settings.defaultTemplate
				) ?? null;
			}
		} catch (err) {
			this.errorMessage = err instanceof Error ? err.message : String(err);
			this.events = [];
			console.error('Outlook Calendar – fetchEvents:', err);
		} finally {
			this.loading = false;
		}

		// Re-render the dropdown with loaded templates (re-render full controls)
		await this.renderControls();
		this.renderEventList();
	}

	// ── Actions ───────────────────────────────────────────────────────────────

	private async onCreate(): Promise<void> {
		if (this.selected.size === 0) {
			new Notice('Bitte mindestens einen Termin auswählen.');
			return;
		}

		let templateContent: string;

		if (this.selectedTemplate) {
			const singleEvent = this.events.find((e) => this.selected.has(e.id));
			if (singleEvent) {
				templateContent = await fillTemplate(this.app.vault, this.selectedTemplate, singleEvent);
			} else {
				templateContent = await this.app.vault.read(this.selectedTemplate);
			}
			// For batch use, we pass raw template and let NoteGenerator fill per-event
			templateContent = await this.app.vault.read(this.selectedTemplate);
		} else {
			templateContent = DEFAULT_TEMPLATE;
		}

		const selectedEvents = this.events.filter((e) => this.selected.has(e.id));
		const result = await generateNotes(
			this.app.vault,
			selectedEvents,
			templateContent,
			this.plugin.settings.outputFolder
		);

		const summary = [
			result.created.length > 0 ? `${result.created.length} erstellt` : null,
			result.skipped.length > 0 ? `${result.skipped.length} übersprungen` : null,
			result.errors.length > 0 ? `${result.errors.length} Fehler` : null,
		]
			.filter(Boolean)
			.join(', ');

		new Notice(`Outlook Calendar: ${summary || 'Fertig.'}`);

		// Open first created note if setting is enabled
		if (this.plugin.settings.openAfterCreate && result.created.length > 0) {
			const firstPath = result.created[0];
			if (firstPath) {
				const file = this.app.vault.getFileByPath(firstPath);
				if (file) {
					await this.app.workspace.getLeaf(false).openFile(file);
				}
			}
		}

		this.close();
	}
}

// ── Default template (used when no template file is selected) ─────────────────

const DEFAULT_TEMPLATE = `# {{Titel}}

**Datum:** {{Datum}}
**Zeit:** {{Startzeit}} – {{Endzeit}}
**Ort:** {{Ort}}
**Organizer:** {{Organizer}}
**Teilnehmer:** {{Teilnehmer}}

---

## Agenda / Notizen

{{Beschreibung}}

---

## Aktionspunkte

- [ ] 

`;

// ── Formatting helpers ────────────────────────────────────────────────────────

function formatDayHeader(dayKey: string): string {
	// dayKey is 'YYYY-MM-DD'
	const date = new Date(dayKey + 'T12:00:00'); // noon to avoid UTC offset issues
	return date.toLocaleDateString('de-DE', {
		weekday: 'long',
		day: '2-digit',
		month: 'long',
		year: 'numeric',
	});
}

function formatTime(iso: string): string {
	if (!iso) return '–';
	return new Date(iso).toLocaleTimeString('de-DE', {hour: '2-digit', minute: '2-digit'});
}
