/**
 * Read text from selected Excel cells.
 */
export async function getSelectedText() {
    return Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load("text");
        await context.sync();
        return range.text.join("\n");
    });
}

/**
 * Write value into selected Excel cell.
 */
export async function writeToSelectedCell(value) {
    return Excel.run(async (context) => {
        const cell = context.workbook.getSelectedRange();
        cell.values = [[value]];
        await context.sync();
    });
}
