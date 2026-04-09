/* global Office */

// Ribbon button command handlers

Office.onReady(() => {
  // Register command handlers
});

/**
 * Shows the task pane when the ribbon button is clicked.
 */
function showTaskpane(event: Office.AddinCommands.Event) {
  Office.addin.showAsTaskpane();
  event.completed();
}

/**
 * Hides the task pane.
 */
function hideTaskpane(event: Office.AddinCommands.Event) {
  Office.addin.hide();
  event.completed();
}

// Register functions with Office
(globalThis as Record<string, unknown>).showTaskpane = showTaskpane;
(globalThis as Record<string, unknown>).hideTaskpane = hideTaskpane;
