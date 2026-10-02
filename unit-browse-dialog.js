/******/ (() => { // webpackBootstrap
/******/ 	"use strict";
// Standalone page opened via Office.context.ui.displayDialogAsync directly
// from taskpane.ts (NOT relayed through commands.ts, unlike the Settings/
// Units dialogs) -- this one needs no Office.context.document access at
// all. It just displays whatever unit list taskpane.ts sends it and
// messages back which unit the user clicked, so taskpane.ts can insert it
// into the annotation textarea in its own DOM. taskpane.ts is the one that
// opened this dialog and stays running the whole time, so the same
// ready-ping/messageParent protocol showNotice/confirmDialog already use
// with their own parent applies here too.
// Cleared once the actual unit list arrives -- see the "ready" ping below.
let readyIntervalId;
function send(message) {
    Office.context.ui.messageParent(JSON.stringify(message));
}
function unitToken(text) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "unit-token";
    button.textContent = text;
    button.onclick = () => send({ type: "selected", text });
    return button;
}
function renderUnits(categories) {
    const container = document.getElementById("unit-list");
    container.innerHTML = "";
    if (categories.length === 0) {
        const empty = document.createElement("p");
        empty.className = "unit-list-empty";
        empty.textContent = "No units defined yet.";
        container.appendChild(empty);
        return;
    }
    for (const category of categories) {
        const categoryBlock = document.createElement("details");
        categoryBlock.className = "unit-category";
        const heading = document.createElement("summary");
        heading.textContent = category.name;
        categoryBlock.appendChild(heading);
        const list = document.createElement("ul");
        for (const unit of category.units) {
            const item = document.createElement("li");
            item.appendChild(unitToken(unit.canonicalName));
            if (unit.aliases.length > 0) {
                item.appendChild(document.createTextNode(" ("));
                unit.aliases.forEach((alias, index) => {
                    item.appendChild(unitToken(alias));
                    if (index < unit.aliases.length - 1) {
                        item.appendChild(document.createTextNode(", "));
                    }
                });
                item.appendChild(document.createTextNode(")"));
            }
            if (unit.isoCode) {
                item.appendChild(document.createTextNode(` [${unit.isoCode}]`));
            }
            if (unit.expansion) {
                item.appendChild(document.createTextNode(` = ${unit.expansion}`));
            }
            list.appendChild(item);
        }
        categoryBlock.appendChild(list);
        container.appendChild(categoryBlock);
    }
}
function onParentMessage(arg) {
    const data = JSON.parse(arg.message);
    if (data.type === "units") {
        if (readyIntervalId !== undefined) {
            window.clearInterval(readyIntervalId);
            readyIntervalId = undefined;
        }
        renderUnits(data.categories);
    }
}
Office.onReady(() => {
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, onParentMessage);
    // taskpane.ts's own handler only starts listening once
    // displayDialogAsync's callback fires, which can happen after this page
    // has already loaded and pinged once -- so keep pinging until the
    // actual unit list arrives, rather than assuming a single ping landed.
    const sendReady = () => send({ type: "ready" });
    sendReady();
    readyIntervalId = window.setInterval(sendReady, 200);
    document.getElementById("dialog-close").onclick =
        () => {
            send({ type: "closed" });
        };
    // Static (not part of the unit data taskpane.ts sends over), so wired up
    // immediately rather than waiting for the "units" message -- lets a user
    // build a compound formula (e.g. "kilogram-meter/second^2") entirely by
    // clicking, without switching back to the keyboard for the operators.
    document
        .querySelectorAll(".syntax-token")
        .forEach((button) => {
        button.onclick = () => { var _a; return send({ type: "selected", text: (_a = button.textContent) !== null && _a !== void 0 ? _a : "" }); };
    });
});


/******/ })()
;