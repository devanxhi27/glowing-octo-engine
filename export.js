function exportData() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getDataRange();
  const values = range.getValues();
  const richTexts = range.getRichTextValues();

  const output = [];

  for (let r = 0; r < values.length; r++) {
    const row = [];

    for (let c = 0; c < values[r].length; c++) {
      const rich = richTexts[r][c];

      // Check if the cell contains a link
      if (rich && rich.getLinkUrl()) {
        row.push({
          text: rich.getText(),      // display text
          url: rich.getLinkUrl()     // hyperlink
        });
      } else {
        row.push(values[r][c]);       // normal value
      }
    }

    output.push(row);
  }

  console.log(output);
}
