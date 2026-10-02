/******/ (() => { // webpackBootstrap
/******/ 	"use strict";
// Standalone page opened via Office.context.ui.displayDialogAsync from the
// task pane for a generic confirmation with a message and two to four
// buttons. It has no Excel/document access of its own; it just renders
// whatever message/labels the task pane sends it and messages back the
// user's choice.
// Cleared once the prompt text actually arrives -- see the "ready" ping
// below.
let readyIntervalId;
Office.onReady(() => {
    const cancelButton = document.getElementById("confirm-cancel");
    const confirmButton = document.getElementById("confirm-ok");
    const extraButton = document.getElementById("confirm-extra");
    const extraCancelButton = document.getElementById("confirm-extra-cancel");
    cancelButton.onclick = () => sendResolution("cancel");
    confirmButton.onclick = () => sendResolution("confirm");
    extraButton.onclick = () => sendResolution("extra");
    extraCancelButton.onclick = () => sendResolution("extraCancel");
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, onParentMessage);
    // Tells the task pane this dialog is ready to receive the prompt text; it
    // can't be passed via the dialog URL since the message has no fixed size
    // or safe encoding for a query string. The task pane only starts
    // listening for this once its displayDialogAsync callback fires, which
    // can happen after this page has already loaded and pinged once -- so
    // keep pinging until the prompt actually arrives, rather than assuming a
    // single ping landed.
    const sendReady = () => Office.context.ui.messageParent(JSON.stringify({ type: "ready" }));
    sendReady();
    readyIntervalId = window.setInterval(sendReady, 200);
});
function sendResolution(choice) {
    const checkbox = document.getElementById("confirm-checkbox");
    Office.context.ui.messageParent(JSON.stringify({ type: "resolved", choice, checked: checkbox.checked }));
}
function setExtraButton(id, label) {
    const button = document.getElementById(id);
    if (label) {
        button.textContent = label;
        button.hidden = false;
    }
    else {
        button.hidden = true;
    }
}
function onParentMessage(arg) {
    var _a, _b;
    const data = JSON.parse(arg.message);
    if (data.type === "prompt") {
        if (readyIntervalId !== undefined) {
            window.clearInterval(readyIntervalId);
            readyIntervalId = undefined;
        }
        document.getElementById("confirm-title").textContent = (_a = data.title) !== null && _a !== void 0 ? _a : "Confirm";
        document.getElementById("confirm-message").textContent = data.text;
        document.getElementById("confirm-cancel").textContent = data.cancelLabel;
        document.getElementById("confirm-ok").textContent =
            data.confirmLabel;
        setExtraButton("confirm-extra", data.extraLabel);
        setExtraButton("confirm-extra-cancel", data.extraCancelLabel);
        const checkboxRow = document.getElementById("confirm-checkbox-row");
        const checkbox = document.getElementById("confirm-checkbox");
        if (data.checkboxLabel) {
            document.getElementById("confirm-checkbox-label").textContent = data.checkboxLabel;
            checkbox.checked = (_b = data.checkboxDefault) !== null && _b !== void 0 ? _b : false;
            checkboxRow.hidden = false;
        }
        else {
            checkbox.checked = false;
            checkboxRow.hidden = true;
        }
    }
}


/******/ })()
;