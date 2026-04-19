// =============================================================================
// This script is maintained in a Git repository and is normally edited in an
// external editor, co-authored with Claude Code.
// Repository: https://github.com/oleksiisedun/blade-runner
//
// ⚠ WARNING: Any changes made directly in the Apps Script web editor may be
// overwritten the next time the code is pushed from the repository.
// =============================================================================

/** @type {PropertiesService.Properties} */
const scriptProperties = PropertiesService.getScriptProperties();

/** @param {String} key */
const getStore = key => JSON.parse(scriptProperties.getProperty(key) ?? "{}");

/**
 * @param {String} key
 * @param {Object} value
 */
const setStore = (key, value) => scriptProperties.setProperty(key, JSON.stringify(value));
