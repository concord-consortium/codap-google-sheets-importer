import React, { useState, useEffect, useCallback } from "react";
import "./codap.css";
import "./styles.css";
import { initializePlugin, createTableWithDataset } from "codap-phone";
import { useInput } from "./hooks";
import {
  createPicker,
  getDataFromSheet,
  makeDataset,
  formatRange,
  firstRowOfCustomRange,
} from "./util";
import Select from "react-select";
import { customStyles } from "./selectStyles";

// This identifies us to Google APIs. Not a secret.
const clientId =
  "756054504415-rf57dsh2mt5vqk1sovptbpopcacctred.apps.googleusercontent.com";

const discoveryDocs = [
  "https://sheets.googleapis.com/$discovery/rest?version=v4",
];

// The drive.file scope lets us use the picker, but the spreadsheets scope is
// needed to access the spreadsheets' data.
const scopes = [
  "https://www.googleapis.com/auth/drive.file",
  "https://www.googleapis.com/auth/spreadsheets.readonly",
];
const scope = scopes.join(" ");

const PLUGIN_TITLE = "Google Sheets Importer";
const PLUGIN_WIDTH = 500;
const PLUGIN_HEIGHT = 700;

export default function Importer() {
  const [error, setError] = useState<string>("");
  const [chosenSpreadsheet, setChosenSpreadsheet] =
    useState<Required<gapi.client.sheets.Spreadsheet> | null>(null);
  const [chosenSheet, chosenSheetChange, setChosenSheet] = useInput<
    string,
    HTMLSelectElement
  >("", () => setError(""));
  const [useHeader, setUseHeader] = useState<boolean>(true);
  const [useCustomRange, setUseCustomRange] = useState<boolean>(false);
  const [customRange, customRangeChange, setCustomRange] = useInput<
    string,
    HTMLInputElement
  >("", () => setError(""));
  const [useAllColumns, setUseAllColumns] = useState<boolean>(true);
  const [columns, setColumns] = useState<string[]>([]);
  const [chosenColumns, setChosenColumns] = useState<string[]>([]);

  function resetState() {
    setError("");
    setChosenSpreadsheet(null);
    setChosenSheet("");
    setUseHeader(true);
    setUseCustomRange(false);
    setCustomRange("");
    setUseAllColumns(true);
    setColumns([]);
    setChosenColumns([]);
  }

  useEffect(() => {
    (async () => {
      if (
        chosenSpreadsheet === null ||
        chosenSheet === "" ||
        (useCustomRange && customRange === "")
      ) {
        setUseAllColumns(true);
        setColumns([]);
        return;
      }

      try {
        let firstRow;
        if (!useCustomRange) {
          firstRow = "1:1";
        } else {
          firstRow = firstRowOfCustomRange(customRange);
        }

        const data = await getDataFromSheet(
          chosenSpreadsheet.spreadsheetId,
          firstRow
        );

        if (data.length === 0) {
          setUseAllColumns(true);
          setColumns([]);
          return;
        }

        setColumns(data[0].map(String));
      } catch (e) {
        setUseAllColumns(true);
        setColumns([]);
      }
    })();
  }, [chosenSpreadsheet, chosenSheet, useCustomRange, customRange]);

  const loginAndCreatePicker = useCallback(async () => {
    function makePickerCallback(token: string) {
      return async (response: google.picker.ResponseObject) => {
        if (
          response[google.picker.Response.ACTION] ===
          google.picker.Action.PICKED
        ) {
          const doc = response[google.picker.Response.DOCUMENTS][0];
          const docId = doc[google.picker.Document.ID];

          let sheet;

          try {
            sheet = (
              await gapi.client.sheets.spreadsheets.get({
                spreadsheetId: docId,
              })
            ).result;
          } catch (e) {
            setError(e.result.error.message);
            return;
          }

          // Cast so that fields are not undefined. We know this spreadsheet
          // exists because the user has picked it.
          setChosenSpreadsheet(
            sheet as Required<gapi.client.sheets.Spreadsheet>
          );

          // Set first sheet as chosen
          if (sheet.sheets && sheet.sheets.length > 0) {
            setChosenSheet(sheet.sheets[0].properties?.title as string);
          }
        } else if (
          response[google.picker.Response.ACTION] ===
          google.picker.Action.CANCEL
        ) {
          // Show picker again if cancelled
          loginAndCreatePicker();
        }
      };
    }

    const GoogleAuth = gapi.auth2.getAuthInstance();

    // Authenticate user so we can read their spreadsheets. This will pop up
    // a Google login window. If signed in, use the stored information.
    const currentUser = GoogleAuth.isSignedIn.get()
      ? GoogleAuth.currentUser.get()
      : await GoogleAuth.signIn();
    const token = currentUser.getAuthResponse().access_token;
    createPicker(token, makePickerCallback(token));
  }, [setChosenSheet]);

  const onClientLoad = useCallback(async () => {
    gapi.client.init({
      discoveryDocs,
      clientId,
      scope,
    });

    await loginAndCreatePicker();
  }, [loginAndCreatePicker]);

  // Load Google APIs upon mounting
  useEffect(() => {
    (async () => {
      try {
        await initializePlugin(PLUGIN_TITLE, PLUGIN_WIDTH, PLUGIN_HEIGHT);
      } catch (e) {
        setError("This plugin must be used within CODAP.");
        return;
      }
      gapi.load("client:auth2:picker", onClientLoad);
    })();
  }, [onClientLoad]);

  async function importSheet() {
    if (chosenSpreadsheet === null) {
      setError("Please choose a spreadsheet.");
      return;
    }

    if (useCustomRange && customRange === "") {
      setError("Please select a valid range.");
      return;
    }

    const range = formatRange(chosenSheet, customRange, useCustomRange);

    let data;

    try {
      data = await getDataFromSheet(chosenSpreadsheet.spreadsheetId, range);
    } catch (e) {
      setError(e.message);
      return;
    }

    if (data.length === 0) {
      setError("Specified range is empty.");
      return;
    }

    // The first element of the tuple will store the column index
    let attributeNames: [number, string][];
    let dataRows: unknown[][];
    if (useHeader) {
      attributeNames = data[0].map((name, index) => [index, String(name)]);

      // Use a filter to preserve original order
      if (!useAllColumns) {
        attributeNames = attributeNames.filter(([, name]) =>
          chosenColumns.includes(name)
        );
      }

      dataRows = data.slice(1);
    } else {
      attributeNames = data[0].map((_value, index) => [
        index,
        `Column ${index}`,
      ]);
      dataRows = data;
    }
    await createTableWithDataset(
      makeDataset(attributeNames, dataRows),
      chosenSpreadsheet.properties.title
    );
    resetState();

    // Show the picker again
    loginAndCreatePicker();
  }

  function cancelImport() {
    resetState();
    loginAndCreatePicker();
  }

  function toggleHeader() {
    setUseHeader(!useHeader);
  }

  function useCustomColumns() {
    if (columns.length === 0) {
      return;
    }
    clearErrorAnd(() => setUseAllColumns(false))();
  }

  function clearErrorAnd(f: () => void) {
    return () => {
      setError("");
      f();
    };
  }

  return (
    <>
      {error !== "" && (
        <div className="error">
          <p>{error}</p>
        </div>
      )}
      {chosenSpreadsheet !== null ? (
        <>
          <div className="input-group">
            <h3>Select a Sheet</h3>
            <select value={chosenSheet} onChange={chosenSheetChange}>
              {chosenSpreadsheet.sheets.map((sheet) => (
                <option
                  key={sheet.properties?.index}
                  value={sheet.properties?.title}
                >
                  {sheet.properties?.title}
                </option>
              ))}
            </select>
          </div>

          <div className="input-group">
            <h3>Column Names</h3>
            <input
              type="checkbox"
              id="useHeader"
              onChange={toggleHeader}
              checked={useHeader}
            />
            <label htmlFor="useHeader">Use first row as column names</label>
          </div>

          <div className="input-group">
            <h3>Range to Import</h3>
            <input
              type="radio"
              id="allValues"
              checked={!useCustomRange}
              onClick={clearErrorAnd(() => setUseCustomRange(false))}
            />
            <label htmlFor="allValues">All values</label>
            <br />
            <input
              type="radio"
              checked={useCustomRange}
              onClick={clearErrorAnd(() => setUseCustomRange(true))}
            />
            <input
              type="text"
              placeholder="A1:C6"
              value={customRange}
              onFocus={clearErrorAnd(() => setUseCustomRange(true))}
              onChange={customRangeChange}
            />
          </div>

          {useHeader && (
            <div className="input-group">
              <h3>Columns</h3>
              <input
                type="radio"
                id="allColumns"
                checked={useAllColumns}
                onClick={clearErrorAnd(() => setUseAllColumns(true))}
              />
              <label htmlFor="allColumns">All columns</label>
              <br />
              <div id="column-selector-row">
                <input
                  type="radio"
                  checked={!useAllColumns}
                  disabled={columns.length === 0}
                  onChange={useCustomColumns}
                />
                <Select
                  styles={customStyles}
                  isMulti
                  isDisabled={columns.length === 0}
                  options={columns.map((n) => ({ value: n, label: n }))}
                  onChange={(selected) => {
                    setChosenColumns(selected.map((s) => s.value));
                  }}
                  onFocus={useCustomColumns}
                />
              </div>
            </div>
          )}

          <div id="submit-buttons" className="input-group">
            <button onClick={importSheet}>Import</button>
            <button onClick={cancelImport}>Cancel</button>
          </div>
        </>
      ) : (
        error !== "" && (
          <div className="input-group">
            <button
              onClick={() => {
                resetState();
                loginAndCreatePicker();
              }}
            >
              Reset
            </button>
          </div>
        )
      )}
    </>
  );
}
