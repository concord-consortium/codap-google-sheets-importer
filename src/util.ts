import { Dataset } from "codap-phone";

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

  if (data.values === undefined) {
    return [];
  }

  return fillRows(data.values);
}

export function createPicker(
  token: string,
  callback: (r: google.picker.ResponseObject) => void
) {
  const view = new google.picker.DocsView(google.picker.ViewId.SPREADSHEETS);
  view.setMode(google.picker.DocsViewMode.LIST);

  const picker = new google.picker.PickerBuilder()
    .setOAuthToken(token)
    .addView(view)
    .enableFeature(google.picker.Feature.NAV_HIDDEN)
    .hideTitleBar()
    .setCallback(callback)
    .build();
  picker.setVisible(true);
}

export function makeDataset(
  attributeNames: [number, string][],
  dataRows: unknown[][]
): Dataset {
  const attributes = attributeNames.map(([, name]) => ({ name }));
  const records = dataRows.map((row) =>
    attributeNames.reduce(
      (acc: Record<string, unknown>, [index, name]: [number, string]) => {
        acc[name] = row[index];
        return acc;
      },
      {}
    )
  );
  return {
    collections: [
      {
        name: "Cases",
        labels: {},
        attrs: attributes,
      },
    ],
    records,
  };
}

export function formatRange(
  sheetName: string,
  customRange: string,
  useCustomRange = true
) {
  return useCustomRange ? `${sheetName}!${customRange}` : sheetName;
}

export function parseRange(range: string): [string, string] {
  const splitByColon = range.split(":");
  if (splitByColon.length !== 2) {
    throw new Error(`Malformed range ${range}`);
  }

  // Safe cast because we checked that the result has two elements
  return splitByColon as [string, string];
}

export function firstRowOfCustomRange(range: string): string {
  const [start, end] = parseRange(range);
  const startRow = start.replace(/[A-Z]/g, "");
  const endColumn = end.replace(/[0-9]/g, "");
  const newEnd = endColumn + startRow;
  return `${start}:${newEnd}`;
}
