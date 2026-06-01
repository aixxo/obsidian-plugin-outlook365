import {Notice} from 'obsidian';
import type OutlookCalendarPlugin from '../main';
import {generateCodeVerifier, generateCodeChallenge, generateState} from './pkce';
import {TokenData} from '../types';

const SCOPES = 'Calendars.Read offline_access';
const REDIRECT_URI = 'obsidian://outlook-calendar-notes';
const TOKEN_ENDPOINT = (tenantId: string) =>
	`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
const AUTH_ENDPOINT = (tenantId: string) =>
	`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`;

/** 10-second window before expiry triggers a refresh. */
const EXPIRY_BUFFER_MS = 10_000;

interface PendingAuth {
	codeVerifier: string;
	state: string;
	resolve: (token: TokenData) => void;
	reject: (err: Error) => void;
}

export class AuthManager {
	private plugin: OutlookCalendarPlugin;
	private pending: PendingAuth | null = null;

	constructor(plugin: OutlookCalendarPlugin) {
		this.plugin = plugin;
	}

	/** Opens the Microsoft login page in the system browser and waits for the callback. */
	async startLogin(): Promise<void> {
		const {clientId, tenantId} = this.plugin.settings;
		if (!clientId) {
			new Notice('Outlook Calendar: Bitte zuerst die Client-ID in den Einstellungen eintragen.');
			return;
		}

		const codeVerifier = generateCodeVerifier();
		const codeChallenge = await generateCodeChallenge(codeVerifier);
		const state = generateState();

		const params = new URLSearchParams({
			client_id: clientId,
			response_type: 'code',
			redirect_uri: REDIRECT_URI,
			scope: SCOPES,
			state,
			code_challenge: codeChallenge,
			code_challenge_method: 'S256',
			response_mode: 'query',
		});

		const authUrl = `${AUTH_ENDPOINT(tenantId)}?${params.toString()}`;

		const tokenData = await new Promise<TokenData>((resolve, reject) => {
			this.pending = {codeVerifier, state, resolve, reject};
			// Open in system browser – works on Windows, macOS, Linux via Electron shell
			window.open(authUrl);
			new Notice('Outlook Calendar: Browser öffnet sich zur Anmeldung…');
		});

		this.plugin.settings.tokenData = tokenData;
		await this.plugin.saveSettings();
		new Notice('Outlook Calendar: Anmeldung erfolgreich!');
	}

	/**
	 * Called by the Obsidian protocol handler when the auth redirect arrives.
	 * URL: obsidian://outlook-calendar-notes?code=...&state=...
	 */
	async handleCallback(params: Record<string, string>): Promise<void> {
		if (!this.pending) {
			return;
		}

		const {code, state, error, error_description} = params;

		if (error) {
			const err = new Error(`OAuth-Fehler: ${error} – ${error_description ?? ''}`);
			this.pending.reject(err);
			this.pending = null;
			new Notice(`Outlook Calendar: ${err.message}`);
			return;
		}

		if (state !== this.pending.state) {
			const err = new Error('OAuth state mismatch – möglicher CSRF-Angriff.');
			this.pending.reject(err);
			this.pending = null;
			new Notice('Outlook Calendar: Sicherheitsfehler bei der Anmeldung.');
			return;
		}

		if (!code) {
			const err = new Error('Kein Autorisierungscode im OAuth-Callback erhalten.');
			this.pending.reject(err);
			this.pending = null;
			new Notice('Outlook Calendar: Anmeldung fehlgeschlagen – kein Code erhalten.');
			return;
		}

		try {
			const tokenData = await this.exchangeCode(code, this.pending.codeVerifier);
			this.pending.resolve(tokenData);
		} catch (err) {
			this.pending.reject(err instanceof Error ? err : new Error(String(err)));
		} finally {
			this.pending = null;
		}
	}

	/** Returns a valid access token, refreshing silently if needed. */
	async getAccessToken(): Promise<string> {
		const {tokenData} = this.plugin.settings;
		if (!tokenData) {
			throw new Error('Nicht angemeldet. Bitte zuerst einloggen.');
		}

		if (Date.now() < tokenData.expiresAt - EXPIRY_BUFFER_MS) {
			return tokenData.accessToken;
		}

		// Silent refresh
		const refreshed = await this.refreshToken(tokenData.refreshToken);
		this.plugin.settings.tokenData = refreshed;
		await this.plugin.saveSettings();
		return refreshed.accessToken;
	}

	isAuthenticated(): boolean {
		return this.plugin.settings.tokenData !== null;
	}

	async logout(): Promise<void> {
		this.plugin.settings.tokenData = null;
		await this.plugin.saveSettings();
		new Notice('Outlook Calendar: Abgemeldet.');
	}

	// ── Private helpers ───────────────────────────────────────────────────────

	private async exchangeCode(code: string, codeVerifier: string): Promise<TokenData> {
		const {clientId, tenantId} = this.plugin.settings;
		const body = new URLSearchParams({
			client_id: clientId,
			grant_type: 'authorization_code',
			code,
			redirect_uri: REDIRECT_URI,
			code_verifier: codeVerifier,
		});

		return this.fetchToken(TOKEN_ENDPOINT(tenantId), body);
	}

	private async refreshToken(refreshToken: string): Promise<TokenData> {
		const {clientId, tenantId} = this.plugin.settings;
		const body = new URLSearchParams({
			client_id: clientId,
			grant_type: 'refresh_token',
			refresh_token: refreshToken,
			scope: SCOPES,
		});

		return this.fetchToken(TOKEN_ENDPOINT(tenantId), body);
	}

	private async fetchToken(url: string, body: URLSearchParams): Promise<TokenData> {
		const response = await fetch(url, {
			method: 'POST',
			headers: {'Content-Type': 'application/x-www-form-urlencoded'},
			body: body.toString(),
		});

		if (!response.ok) {
			const text = await response.text();
			throw new Error(`Token-Anfrage fehlgeschlagen (${response.status}): ${text}`);
		}

		const json = await response.json() as {
			access_token: string;
			refresh_token?: string;
			expires_in: number;
			scope: string;
		};

		const currentRefreshToken = body.get('refresh_token') ?? '';

		return {
			accessToken: json.access_token,
			refreshToken: json.refresh_token ?? currentRefreshToken,
			expiresAt: Date.now() + json.expires_in * 1000,
			scope: json.scope,
		};
	}
}
