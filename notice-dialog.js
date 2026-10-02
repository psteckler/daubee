/******/ (() => { // webpackBootstrap
/******/ 	"use strict";

// Standalone page opened via Office.context.ui.displayDialogAsync from the
// task pane to show a single-message notice with an OK button. It has no
// Excel/document access of its own; it just renders whatever message the
// task pane sends it and messages back once dismissed.
// Cleared once the message text actually arrives -- see the "ready" ping
// below.
let readyIntervalId;
Office.onReady(() => {
    const okButton = document.getElementById("notice-ok");
    okButton.onclick = () => dismiss();
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, onParentMessage);
    // Tells the task pane this dialog is ready to receive the message; it
    // can't be passed via the dialog URL since the message text has no fixed
    // size or safe encoding for a query string. The task pane only starts
    // listening for this once its displayDialogAsync callback fires, which
    // can happen after this page has already loaded and pinged once -- so
    // keep pinging until the message actually arrives, rather than assuming a
    // single ping landed.
    const sendReady = () => Office.context.ui.messageParent(JSON.stringify({ type: "ready" }));
    sendReady();
    readyIntervalId = window.setInterval(sendReady, 200);
});
function dismiss() {
    Office.context.ui.messageParent(JSON.stringify({ type: "dismissed" }));
}
function onParentMessage(arg) {
    const data = JSON.parse(arg.message);
    if (data.type === "message") {
        if (readyIntervalId !== undefined) {
            window.clearInterval(readyIntervalId);
            readyIntervalId = undefined;
        }
        document.getElementById("notice-message").textContent = data.text;
    }
}

/******/ })()
;