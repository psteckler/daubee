/******/ (() => { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ 3152
(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   OI: () => (/* binding */ describeScope)
/* harmony export */ });
/* unused harmony exports EXCEL_MAX_ROW_INDEX, EXCEL_MAX_COL_INDEX, columnIndexToLabel, parseCellRef, scopeBounds, clampToSheet, boundsToAddress, scopeContainsAddress, boundsContain, scopeAddress, AnnotationStore */
// Data model and persistence for cell/range/row/column annotations.
//
// Annotation "formula" text is arbitrary for now; the structure will be
// defined later. Row and column annotations are stored once per row/column,
// not duplicated per cell.
const SETTINGS_KEY = "daubee.annotations";
function generateId() {
    return `ann_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`;
}
function scopesEqual(a, b) {
    if (a.kind !== b.kind || a.sheetId !== b.sheetId) {
        return false;
    }
    switch (a.kind) {
        case "cell":
        case "range":
            return a.address === b.address;
        case "column":
            return a.columnIndex === b.columnIndex;
        case "row":
            return a.rowIndex === b.rowIndex;
    }
}
const MIN_INDEX = 0;
const MAX_INDEX = Number.MAX_SAFE_INTEGER;
// Real Excel worksheet limits, used to turn an unbounded column/row scope's
// "infinite" edge (MAX_INDEX) into a concrete address when a piece of it
// survives a split (e.g. the part of a column below an overridden cell).
const EXCEL_MAX_ROW_INDEX = 1048575;
const EXCEL_MAX_COL_INDEX = 16383;
function columnIndexToLabel(columnIndex) {
    let index = columnIndex;
    let label = "";
    do {
        label = String.fromCharCode((index % 26) + 65) + label;
        index = Math.floor(index / 26) - 1;
    } while (index >= 0);
    return label;
}
function columnLabelToIndex(label) {
    let index = 0;
    for (const ch of label.toUpperCase()) {
        index = index * 26 + (ch.charCodeAt(0) - 64);
    }
    return index - 1;
}
function parseCellRef(ref) {
    const match = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(ref);
    if (!match) {
        throw new Error(`Unrecognized cell reference: ${ref}`);
    }
    return { row: parseInt(match[2], 10) - 1, col: columnLabelToIndex(match[1]) };
}
function addressBounds(address) {
    const [start, end] = address.split(":");
    const startRef = parseCellRef(start);
    const endRef = end ? parseCellRef(end) : startRef;
    return {
        minRow: Math.min(startRef.row, endRef.row),
        maxRow: Math.max(startRef.row, endRef.row),
        minCol: Math.min(startRef.col, endRef.col),
        maxCol: Math.max(startRef.col, endRef.col),
    };
}
function scopeBounds(scope) {
    switch (scope.kind) {
        case "cell":
        case "range":
            return addressBounds(scope.address);
        case "column":
            return {
                minRow: MIN_INDEX,
                maxRow: MAX_INDEX,
                minCol: scope.columnIndex,
                maxCol: scope.columnIndex,
            };
        case "row":
            return {
                minRow: scope.rowIndex,
                maxRow: scope.rowIndex,
                minCol: MIN_INDEX,
                maxCol: MAX_INDEX,
            };
    }
}
function boundsOverlap(a, b) {
    return (a.minRow <= b.maxRow &&
        b.minRow <= a.maxRow &&
        a.minCol <= b.maxCol &&
        b.minCol <= a.maxCol);
}
function isEmptyBounds(b) {
    return b.minRow > b.maxRow || b.minCol > b.maxCol;
}
// Exported alongside boundsContain below for taskpane.ts's bulk "delete all
// units within selection" action, which needs to compare a column/row
// annotation's unbounded (MAX_INDEX) edge against a selection's own concrete
// edge on equal footing.
function clampToSheet(b) {
    return {
        minRow: Math.max(b.minRow, MIN_INDEX),
        maxRow: Math.min(b.maxRow, EXCEL_MAX_ROW_INDEX),
        minCol: Math.max(b.minCol, MIN_INDEX),
        maxCol: Math.min(b.maxCol, EXCEL_MAX_COL_INDEX),
    };
}
// Splits `base` into the rectangles that remain once the part overlapping
// `hole` is removed: a full-width strip above the hole, a full-width strip
// below it, then left/right strips spanning just the hole's own rows. That
// tiling exactly covers base \ hole with no overlaps, in up to 4 pieces.
function subtractBounds(base, hole) {
    const clippedHole = {
        minRow: Math.max(base.minRow, hole.minRow),
        maxRow: Math.min(base.maxRow, hole.maxRow),
        minCol: Math.max(base.minCol, hole.minCol),
        maxCol: Math.min(base.maxCol, hole.maxCol),
    };
    if (isEmptyBounds(clippedHole)) {
        return [base];
    }
    const pieces = [];
    if (clippedHole.minRow > base.minRow) {
        pieces.push({
            minRow: base.minRow,
            maxRow: clippedHole.minRow - 1,
            minCol: base.minCol,
            maxCol: base.maxCol,
        });
    }
    if (clippedHole.maxRow < base.maxRow) {
        pieces.push({
            minRow: clippedHole.maxRow + 1,
            maxRow: base.maxRow,
            minCol: base.minCol,
            maxCol: base.maxCol,
        });
    }
    if (clippedHole.minCol > base.minCol) {
        pieces.push({
            minRow: clippedHole.minRow,
            maxRow: clippedHole.maxRow,
            minCol: base.minCol,
            maxCol: clippedHole.minCol - 1,
        });
    }
    if (clippedHole.maxCol < base.maxCol) {
        pieces.push({
            minRow: clippedHole.minRow,
            maxRow: clippedHole.maxRow,
            minCol: clippedHole.maxCol + 1,
            maxCol: base.maxCol,
        });
    }
    return pieces;
}
function boundsToAddress(bounds) {
    const start = `${columnIndexToLabel(bounds.minCol)}${bounds.minRow + 1}`;
    if (bounds.minRow === bounds.maxRow && bounds.minCol === bounds.maxCol) {
        return start;
    }
    const end = `${columnIndexToLabel(bounds.maxCol)}${bounds.maxRow + 1}`;
    return `${start}:${end}`;
}
// Converts a leftover piece of a split annotation back into a concrete
// cell/range scope (column/row scopes only apply to a whole column/row, so a
// partial leftover can no longer be one).
function boundsToScope(sheetId, sheetName, bounds) {
    const clamped = clampToSheet(bounds);
    const address = boundsToAddress(clamped);
    return clamped.minRow === clamped.maxRow && clamped.minCol === clamped.maxCol
        ? { kind: "cell", sheetId, sheetName, address }
        : { kind: "range", sheetId, sheetName, address };
}
// Whether a single cell address falls within a scope's bounds.
function scopeContainsAddress(scope, address) {
    const { row, col } = parseCellRef(address);
    const bounds = scopeBounds(scope);
    return (row >= bounds.minRow &&
        row <= bounds.maxRow &&
        col >= bounds.minCol &&
        col <= bounds.maxCol);
}
// Whether `inner`'s bounds fit entirely within `outer`'s -- used to find
// every annotation fully covered by a selection, e.g. for a bulk-delete
// action, as opposed to boundsOverlap's looser "the two rectangles touch at
// all" test used for conflict detection.
function boundsContain(outer, inner) {
    return (inner.minRow >= outer.minRow &&
        inner.maxRow <= outer.maxRow &&
        inner.minCol >= outer.minCol &&
        inner.maxCol <= outer.maxCol);
}
// The Excel range address a scope covers, suitable for Worksheet.getRange
// (column/row scopes use the whole-column/row "A:A" / "5:5" syntax).
function scopeAddress(scope) {
    switch (scope.kind) {
        case "cell":
        case "range":
            return scope.address;
        case "column":
            return `${scope.columnLabel}:${scope.columnLabel}`;
        case "row":
            return `${scope.rowIndex + 1}:${scope.rowIndex + 1}`;
    }
}
function describeScope(scope) {
    switch (scope.kind) {
        case "cell":
            return `${scope.sheetName}!${scope.address} (cell)`;
        case "range":
            return `${scope.sheetName}!${scope.address} (range)`;
        case "column":
            return `${scope.sheetName} column ${scope.columnLabel}`;
        case "row":
            return `${scope.sheetName} row ${scope.rowIndex + 1}`;
    }
}
// Wraps Office.context.document.settings, which persists as part of the
// workbook file, so annotations survive save/close/reopen.
class AnnotationStore {
    constructor(annotations) {
        this.annotations = annotations;
    }
    static load() {
        const raw = Office.context.document.settings.get(SETTINGS_KEY);
        const annotations = raw ? JSON.parse(raw) : [];
        return new AnnotationStore(annotations);
    }
    persist() {
        Office.context.document.settings.set(SETTINGS_KEY, JSON.stringify(this.annotations));
        return new Promise((resolve, reject) => {
            Office.context.document.settings.saveAsync((result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    reject(result.error.message);
                }
                else {
                    resolve();
                }
            });
        });
    }
    getAll() {
        return [...this.annotations];
    }
    getForSheet(sheetId) {
        return this.annotations.filter((a) => a.scope.sheetId === sheetId);
    }
    // The single annotation (if any) whose scope covers the given cell
    // address. Conflicting scopes are always split apart by upsert()/
    // override(), so at most one annotation should ever cover a given cell.
    findForAddress(sheetId, address) {
        return this.annotations.find((a) => a.scope.sheetId === sheetId && scopeContainsAddress(a.scope, address));
    }
    // Existing annotations (on the same sheet) whose scope overlaps the given
    // scope but isn't identical to it. An identical scope is treated as an
    // in-place update by upsert(), not a conflict, so it's excluded here.
    findConflicts(scope) {
        const bounds = scopeBounds(scope);
        return this.annotations.filter((a) => a.scope.sheetId === scope.sheetId &&
            !scopesEqual(a.scope, scope) &&
            boundsOverlap(scopeBounds(a.scope), bounds));
    }
    // Creates a new annotation, or overwrites the existing one for the same
    // scope (e.g. re-annotating the same column replaces its text).
    applyUpsert(scope, formula, now) {
        const existingIndex = this.annotations.findIndex((a) => scopesEqual(a.scope, scope));
        let annotation;
        if (existingIndex >= 0) {
            annotation = Object.assign(Object.assign({}, this.annotations[existingIndex]), { formula, updatedAt: now });
            this.annotations[existingIndex] = annotation;
        }
        else {
            annotation = {
                id: generateId(),
                scope,
                formula,
                createdAt: now,
                updatedAt: now,
            };
            this.annotations.push(annotation);
        }
        return annotation;
    }
    async upsert(scope, formula) {
        const annotation = this.applyUpsert(scope, formula, new Date().toISOString());
        await this.persist();
        return annotation;
    }
    async remove(id) {
        this.annotations = this.annotations.filter((a) => a.id !== id);
        await this.persist();
    }
    // Saves `formula` over `newScopes`, but instead of deleting each entry in
    // `conflicts` outright, carves the new scopes' cells out of it and keeps
    // the rest: e.g. overriding one cell inside an annotated column leaves the
    // annotation on every other cell of that column untouched.
    async override(conflicts, newScopes, formula) {
        const now = new Date().toISOString();
        const conflictIds = new Set(conflicts.map((c) => c.id));
        const holes = newScopes.map((scope) => scopeBounds(scope));
        const remainders = [];
        for (const conflict of conflicts) {
            let pieces = [scopeBounds(conflict.scope)];
            for (const hole of holes) {
                const next = [];
                for (const piece of pieces) {
                    for (const split of subtractBounds(piece, hole)) {
                        next.push(split);
                    }
                }
                pieces = next;
            }
            for (const bounds of pieces) {
                remainders.push({
                    id: generateId(),
                    scope: boundsToScope(conflict.scope.sheetId, conflict.scope.sheetName, bounds),
                    formula: conflict.formula,
                    createdAt: now,
                    updatedAt: now,
                });
            }
        }
        this.annotations = this.annotations
            .filter((a) => !conflictIds.has(a.id))
            .concat(remainders);
        const saved = newScopes.map((scope) => this.applyUpsert(scope, formula, now));
        await this.persist();
        return saved;
    }
}


/***/ }

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	const __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		const cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		const module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/define property getters */
/******/ 	(() => {
/******/ 		// define getter/value functions for harmony exports
/******/ 		__webpack_require__.d = (exports, definition) => {
/******/ 			if(Array.isArray(definition)) {
/******/ 				var i = 0;
/******/ 				while(i < definition.length) {
/******/ 					var key = definition[i++];
/******/ 					var binding = definition[i++];
/******/ 					if(!__webpack_require__.o(exports, key)) {
/******/ 						if(binding === 0) {
/******/ 							Object.defineProperty(exports, key, { enumerable: true, value: definition[i++] });
/******/ 						} else {
/******/ 							Object.defineProperty(exports, key, { enumerable: true, get: binding });
/******/ 						}
/******/ 					} else if(binding === 0) { i++; }
/******/ 				}
/******/ 			} else {
/******/ 				for(var key in definition) {
/******/ 					if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 						Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 					}
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	(() => {
/******/ 		__webpack_require__.o = (obj, prop) => (Object.prototype.hasOwnProperty.call(obj, prop))
/******/ 	})();
/******/ 	
/************************************************************************/
let __webpack_exports__ = {};
/* harmony import */ var _taskpane_annotations__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(3152);
// Standalone page opened via Office.context.ui.displayDialogAsync from the
// task pane when a new annotation overlaps an existing one. It has no
// Excel/document access of its own; it just renders whatever conflict list
// the task pane sends it and messages back the user's choice.

// Cleared once the conflict list actually arrives -- see the "ready" ping
// below.
let readyIntervalId;
Office.onReady(() => {
    const cancelButton = document.getElementById("conflict-cancel");
    const overrideButton = document.getElementById("conflict-override");
    cancelButton.onclick = () => sendResolution(false);
    overrideButton.onclick = () => sendResolution(true);
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, onParentMessage);
    // Tells the task pane this dialog is ready to receive the conflict list;
    // it can't be passed via the dialog URL since displayDialogAsync only
    // takes one, and the list has no fixed size. The task pane only starts
    // listening for this once its displayDialogAsync callback fires, which
    // can happen after this page has already loaded and pinged once -- so
    // keep pinging until the conflict list actually arrives, rather than
    // assuming a single ping landed.
    const sendReady = () => Office.context.ui.messageParent(JSON.stringify({ type: "ready" }));
    sendReady();
    readyIntervalId = window.setInterval(sendReady, 200);
});
function sendResolution(override) {
    Office.context.ui.messageParent(JSON.stringify({ type: "resolved", override }));
}
function onParentMessage(arg) {
    const data = JSON.parse(arg.message);
    if (data.type === "conflicts") {
        if (readyIntervalId !== undefined) {
            window.clearInterval(readyIntervalId);
            readyIntervalId = undefined;
        }
        renderConflicts(data.conflicts);
    }
}
function renderConflicts(conflicts) {
    const list = document.getElementById("conflict-list");
    list.innerHTML = "";
    for (const conflict of conflicts) {
        const item = document.createElement("li");
        item.textContent = `${(0,_taskpane_annotations__WEBPACK_IMPORTED_MODULE_0__/* .describeScope */ .OI)(conflict.scope)}: ${conflict.formula}`;
        list.appendChild(item);
    }
}

/******/ })()
;