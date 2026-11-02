/**
 * ITR Generator — frontend configuration
 * Paste your Google Apps Script Web App URL below (Deploy → Manage deployments → Web app URL).
 * You can also set it from the "Backend settings" link on the sign-in screen; that value is saved
 * in the browser and overrides this one.
 */
window.ITR_CONFIG = {
  API_URL: 'PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE',
  FIRST_YEAR: 2026,      // earliest taxable year offered
  YEARS_AHEAD: 4,        // list runs to at least current year + this many
  PARALLEL_UPLOADS: 3,
  IMAGE_MAX_DIMENSION: 2000,
  IMAGE_QUALITY: 0.85
};
