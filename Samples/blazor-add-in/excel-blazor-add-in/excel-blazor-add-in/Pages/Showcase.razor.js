/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
console.log("Loading Showcase.razor.js");

const LOCK_FILL = "#D3D3D3";
const PROTECT_OPTIONS = {
    allowInsertRows: true,
    allowInsertColumns: true,
    allowDeleteRows: true,
    allowDeleteColumns: true,
    allowFormatCells: true,
    allowFormatRows: true,
    allowFormatColumns: true,
    allowSort: true,
    allowAutoFilter: true,
    allowPivotTables: true
};

let allCellsUnlocked = false;

async function applyLock(address, locked) {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.protection.unprotect();
        await context.sync();

        if (!allCellsUnlocked) {
            sheet.getRange().format.protection.locked = false;
            allCellsUnlocked = true;
            await context.sync();
        }

        const range = sheet.getRange(address);
        if (locked) {
            range.format.fill.color = LOCK_FILL;
            range.format.protection.locked = true;
        } else {
            range.format.fill.clear();
            range.format.protection.locked = false;
        }

        sheet.protection.protect(PROTECT_OPTIONS);
        await context.sync();
    });
}

export async function setRowLock(rowNumber, locked) {
    const n = parseInt(rowNumber, 10);
    if (!n || n < 1) return;
    await applyLock(`${n}:${n}`, locked);
}

export async function setColumnLock(letter, locked) {
    const col = (letter || "").trim().toUpperCase();
    if (!/^[A-Z]{1,3}$/.test(col)) return;
    await applyLock(`${col}:${col}`, locked);
}

export async function setCellLock(letter, rowNumber, locked) {
    const col = (letter || "").trim().toUpperCase();
    const n = parseInt(rowNumber, 10);
    if (!/^[A-Z]{1,3}$/.test(col) || !n || n < 1) return;
    await applyLock(`${col}${n}`, locked);
}

export async function writeFormToSheet(model) {
    await Excel.run(async (context) => {
        context.workbook.worksheets.getItemOrNullObject("Showcase").delete();
        const sheet = context.workbook.worksheets.add("Showcase");

        const rows = [
            ["Field", "Value"],
            ["Name", model.name ?? ""],
            ["Email", model.email ?? ""],
            ["Notes", model.notes ?? ""],
            ["Age", model.age ?? 0],
            ["Start date", new Date(model.startDate).toLocaleDateString()],
            ["Role", model.role ?? ""],
            ["Rating", `${model.rating ?? 0} / 10`],
            ["Favorite color", model.color ?? ""],
            ["Preferred contact", model.contact ?? ""],
            ["Subscribed", model.subscribe ? "Yes" : "No"]
        ];

        const range = sheet.getRangeByIndexes(0, 0, rows.length, 2);
        range.values = rows;

        const header = sheet.getRangeByIndexes(0, 0, 1, 2);
        header.format.fill.color = "#0078d4";
        header.format.font.color = "white";
        header.format.font.bold = true;

        // Highlight the favorite color cell with that color.
        const colorCell = sheet.getRangeByIndexes(8, 1, 1, 1);
        if (model.color) {
            colorCell.format.fill.color = model.color;
        }

        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
        sheet.activate();

        await context.sync();
    });
}
