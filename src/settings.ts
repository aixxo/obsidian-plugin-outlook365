import {OutlookPluginSettings} from "./types";

export const DEFAULT_SETTINGS: OutlookPluginSettings = {
	clientId: '',
	tenantId: 'common',
	templateFolder: 'Templates',
	outputFolder: 'Meetings',
	defaultDateRange: 'combined',
	defaultTemplate: '',
	openAfterCreate: false,
	onDuplicate: 'skip',
	tokenData: null,
};
