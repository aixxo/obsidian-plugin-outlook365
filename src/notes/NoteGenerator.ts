import {Notice, TFile, Vault} from 'obsidian';
import {CalendarEvent} from '../types';
import {replacePlaceholders} from './TemplateParser';

/** Sanitizes a string for use as a file-system-safe filename component. */
function sanitize(text: string): string {
	return text.replace(/[\\/:*?"<>|#^[\]]/g, '_').trim();
}

/** Formats an ISO date string as YYYY-MM-DD. */
function isoDate(iso: string): string {
	if (!iso) return 'unbekannt';
	return new Date(iso).toISOString().split('T')[0] ?? 'unbekannt';
}

export interface NoteGeneratorResult {
	created: string[];
	skipped: string[];
	errors: string[];
}

/**
 * Generates Obsidian notes from calendar events using the provided template content.
 *
 * @param vault         The Obsidian vault
 * @param events        Events to convert
 * @param templateContent  Raw template string with {{Placeholder}} markers
 * @param outputFolder  Target folder path (will be created if missing)
 */
export async function generateNotes(
	vault: Vault,
	events: CalendarEvent[],
	templateContent: string,
	outputFolder: string
): Promise<NoteGeneratorResult> {
	const result: NoteGeneratorResult = {created: [], skipped: [], errors: []};

	// Ensure output folder exists
	if (!vault.getFolderByPath(outputFolder)) {
		await vault.createFolder(outputFolder);
	}

	for (const event of events) {
		const fileName = buildFileName(event, outputFolder);
		try {
			const existing = vault.getFileByPath(fileName);
			if (existing instanceof TFile) {
				result.skipped.push(fileName);
				new Notice(`Übersprungen (existiert bereits): ${fileName}`);
				continue;
			}

			const content = replacePlaceholders(templateContent, event);
			await vault.create(fileName, content);
			result.created.push(fileName);
		} catch (err) {
			const msg = err instanceof Error ? err.message : String(err);
			result.errors.push(`${fileName}: ${msg}`);
			new Notice(`Fehler beim Erstellen von "${fileName}": ${msg}`);
		}
	}

	return result;
}

function buildFileName(event: CalendarEvent, folder: string): string {
	const title = sanitize(event.subject);
	const date = isoDate(event.start);
	return `${folder}/${title}_${date}.md`;
}
