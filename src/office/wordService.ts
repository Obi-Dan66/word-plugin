/* global Word */

export async function getSelectionText(): Promise<string> {
  return Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("text");
    await context.sync();
    return range.text;
  });
}

export async function replaceSelectionText(newText: string): Promise<void> {
  await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.insertText(newText, Word.InsertLocation.replace);
    await context.sync();
  });
}
