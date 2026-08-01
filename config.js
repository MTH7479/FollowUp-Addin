/*
 * Follow Up add-in configuration.
 * FLOW_URL is the HTTP trigger of the deployed "PP FollowUp - Create" flow.
 * SECURITY: this URL contains a shared-access signature. Treat it as a secret.
 * Anyone who can read this file can POST to the flow. Host the add-in only on
 * an internal/trusted HTTPS site, and consider adding a shared-secret check in
 * the flow (see SETUP-GUIDE.md, "Security").
 */
window.FOLLOWUP_CONFIG = {
  FLOW_URL: "https://defaultac53278e996549ee9675e63f500525.2f.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/14/workflows/2c47552ad3b04a36800e499d004bb425/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=qbWwsSrvkKz8moGE47joW_IAiPHYyE8M5qz-cdTPIUo",
  DEFAULT_REMINDER_DAYS: 3,
  DEFAULT_MAX_REMINDERS: 3,
  DEFAULT_PRIORITY: 126760001 // Normal
};
