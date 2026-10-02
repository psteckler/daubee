/******/ (() => { // webpackBootstrap
/******/ 	"use strict";
// Standalone page opened via Office.context.ui.displayDialogAsync from
// commands.ts's own "openAbout" ExecuteFunction handler. Its content is
// entirely static -- there's nothing to request from a parent -- but it
// still has to ask that parent to close it: Office.context.ui.closeContainer()
// is documented as closing a *task pane* (or having no effect from a
// UI-less command button), never a dialog opened via displayDialogAsync,
// so a dialog can't close itself at all -- only the parent holding the
// Dialog object from displayDialogAsync's own callback can call .close()
// on it. Same ready-ping-free, message-on-dismiss shape every other
// dialog here uses, just without ever needing to receive anything back.
Office.onReady(() => {
    document.getElementById("about-close").onclick = () => Office.context.ui.messageParent(JSON.stringify({ type: "closed" }));
});


/******/ })()
;