/**
 * OAuth 2.0 PKCE helpers using the Web Crypto API (available in Electron/Obsidian).
 * Implements RFC 7636 – Proof Key for Code Exchange.
 */

function base64UrlEncode(buffer: ArrayBuffer): string {
	const bytes = new Uint8Array(buffer);
	let binary = '';
	for (let i = 0; i < bytes.byteLength; i++) {
		binary += String.fromCharCode(bytes[i] as number);
	}
	return btoa(binary)
		.replace(/\+/g, '-')
		.replace(/\//g, '_')
		.replace(/=/g, '');
}

/** Generates a cryptographically random code verifier (43–128 chars, RFC 7636). */
export function generateCodeVerifier(): string {
	const array = new Uint8Array(32);
	crypto.getRandomValues(array);
	return base64UrlEncode(array.buffer);
}

/** Derives the S256 code challenge from a code verifier. */
export async function generateCodeChallenge(verifier: string): Promise<string> {
	const encoder = new TextEncoder();
	const data = encoder.encode(verifier);
	const digest = await crypto.subtle.digest('SHA-256', data);
	return base64UrlEncode(digest);
}

/** Generates a cryptographically random state string for CSRF protection. */
export function generateState(): string {
	const array = new Uint8Array(16);
	crypto.getRandomValues(array);
	return base64UrlEncode(array.buffer);
}
