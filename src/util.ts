function fillRows(table: unknown[][]): unknown[][] {
  const longestLength = table.reduce(
    (currentMax, row) => (row.length > currentMax ? row.length : currentMax),
    0
  );

  // Fill rows not long enough with empty strings
  return table.map((row) =>
    row.length < longestLength
      ? row.fill("", row.length, longestLength - 1)
      : row
  );
}

export async function getDataFromSheet(
  sheetId: string,
  range: string
): Promise<unknown[][]> {
  const data = (
    await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range,
    })
  ).result;
  console.log(data.values);

  if (data.values === undefined) {
    return [];
  }

  return fillRows(data.values);
}
