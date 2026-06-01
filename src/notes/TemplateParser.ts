import {TFile, Vault} from 'obsidian';
import {CalendarEvent} from '../types';

const PLACEHOLDER_MAP: Record<string, (e: CalendarEvent) => string> = {
	Titel:        (e) => e.subject,
	Datum:        (e) => formatDate(e.start),
	Startzeit:    (e) => formatTime(e.start),
	Endzeit:      (e) => formatTime(e.end),
	Ort:          (e) => e.location || '–',
	Beschreibung: (e) => e.bodyPreview || '–',
	Teilnehmer:   (e) => e.attendees.join(', ') || '–',
	Organizer:    (e) => e.organizer || '–',
	OnlineMeetingUrl: (e) => e.onlineMeetingUrl || '–',
};

/** Returns a list of all .md files in the given vault folder. */
export async function listTemplates(vault: Vault, folderPath: string): Promise<TFile[]> {
	const folder = vault.getFolderByPath(folderPath);
	if (!folder) return [];
	return folder.children.filter(
		(f): f is TFile => f instanceof TFile && f.extension === 'md'
	);
}

/** Reads a template file and fills in all {{Placeholder}} values. */
export async function fillTemplate(
	vault: Vault,
	templateFile: TFile,
	event: CalendarEvent
): Promise<string> {
	const content = await vault.read(templateFile);
	return replacePlaceholders(content, event);
}

/** Fills placeholders in an arbitrary template string. */
export function replacePlaceholders(template: string, event: CalendarEvent): string {
	return template.replace(/\{\{(\w+)\}\}/g, (_match, key: string) => {
		const fn = PLACEHOLDER_MAP[key];
		return fn ? fn(event) : '';
	});
}

// ── Date helpers ──────────────────────────────────────────────────────────────

function formatDate(iso: string): string {
	if (!iso) return '–';
	return new Date(iso).toLocaleDateString('de-DE', {
		day: '2-digit',
		month: '2-digit',
		year: 'numeric',
	});
}

function formatTime(iso: string): string {
	if (!iso) return '–';
	return new Date(iso).toLocaleTimeString('de-DE', {
		hour: '2-digit',
		minute: '2-digit',
	});
}
