/**
 * Make rows in 2D array all have the same length by filling rows that are
 * not long enough with empty string.
 *
 * @param table - The input 2D array
 * @returns The filled 2D array
 */
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
  let data;
  try {
    data = (
      await gapi.client.sheets.spreadsheets.values.get({
        spreadsheetId: sheetId,
        range,
      })
    ).result;
  } catch (e) {
    throw e.result.error;
  }
  console.log(data.values);

  if (data.values === undefined) {
    return [];
  }

  return fillRows(data.values);
}

export function createPicker(
  token: string,
  callback: (r: google.picker.ResponseObject) => void
) {
  const picker = new google.picker.PickerBuilder()
    .setOAuthToken(token)
    .addView(google.picker.ViewId.SPREADSHEETS)
    .enableFeature(google.picker.Feature.NAV_HIDDEN)
    .hideTitleBar()
    .setCallback(callback)
    .build();
  picker.setVisible(true);
}
