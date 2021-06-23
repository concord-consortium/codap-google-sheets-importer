import React, { useState, useEffect, useCallback } from "react";
import "./codap.css";
import "./App.css";
import { initializePlugin, createTableWithDataset } from "codap-phone";
import { useInput } from "./hooks";
import { createPicker, getDataFromSheet } from "./util";

// This identifies us to Google APIs. Not a secret.
const CLIENT_ID =
  "756054504415-rf57dsh2mt5vqk1sovptbpopcacctred.apps.googleusercontent.com";
const DISCOVERY_DOCS = [
  "https://sheets.googleapis.com/$discovery/rest?version=v4",
];

// Lets us see all of the user's spreadsheets, which means that all thumbnails
// will be available in the picker.
const scope = "https://www.googleapis.com/auth/spreadsheets.readonly";

const PLUGIN_TITLE = "Google Sheets Importer";
const PLUGIN_WIDTH = 500;
const PLUGIN_HEIGHT = 700;

export default function App() {
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

  const makePickerCallback = useCallback(
    (token: string) => {
      return async (response: google.picker.ResponseObject) => {
        if (
          response[google.picker.Response.ACTION] ===
          google.picker.Action.PICKED
        ) {
          const doc = response[google.picker.Response.DOCUMENTS][0];
          console.log(doc);
          const sheet = (
            await gapi.client.sheets.spreadsheets.get({ spreadsheetId: doc.id })
          ).result;
          console.log(sheet);
          setChosenSpreadsheet(
            sheet as Required<gapi.client.sheets.Spreadsheet>
          );

          // Set first sheet as chosen
          if (sheet.sheets && sheet.sheets.length > 0) {
            setChosenSheet(sheet.sheets[0].properties?.title as string);
          }
        }
      };
    },
    [setChosenSheet]
  );

  const loginAndCreatePicker = useCallback(async () => {
    const GoogleAuth = gapi.auth2.getAuthInstance();
    const currentUser = GoogleAuth.isSignedIn.get()
      ? GoogleAuth.currentUser.get()
      : await GoogleAuth.signIn();
    const token = currentUser.getAuthResponse().access_token;
    createPicker(token, makePickerCallback(token));
  }, [makePickerCallback]);

  // Authenticate user so we can read their spreadsheets. This will pop up
  // a Google login window.
  const onClientLoad = useCallback(async () => {
    gapi.client.init({
      discoveryDocs: DISCOVERY_DOCS,
      clientId: CLIENT_ID,
      scope,
    });

    await loginAndCreatePicker();
  }, [loginAndCreatePicker]);

  // Load Google APIs upon mounting
  useEffect(() => {
    initializePlugin(PLUGIN_TITLE, PLUGIN_WIDTH, PLUGIN_HEIGHT);
    gapi.load("client:auth2:picker", onClientLoad);
  }, [onClientLoad]);

  function resetState() {
    setChosenSpreadsheet(null);
    setChosenSheet("");
    setUseHeader(false);
    setUseCustomRange(false);
    setCustomRange("");
  }

  async function importSheet() {
    if (chosenSpreadsheet === null) {
      setError("Please choose a spreadsheet.");
      return;
    }

    if (useCustomRange && customRange === "") {
      setError("Please select a valid range.");
      return;
    }

    const range = useCustomRange
      ? `${chosenSheet}!${customRange}`
      : chosenSheet;

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

    let attributeNames: string[];
    let dataRows: unknown[][];
    if (useHeader) {
      attributeNames = data[0].map(String);
      dataRows = data.slice(1);
    } else {
      attributeNames = data[0].map((_value, index) => `Column ${index}`);
      dataRows = data;
    }
    const attributes = attributeNames.map((name) => ({ name }));
    const records = dataRows.map((row) =>
      attributeNames.reduce(
        (acc: Record<string, unknown>, name: string, i: number) => {
          acc[name] = row[i];
          return acc;
        },
        {}
      )
    );
    await createTableWithDataset(
      {
        collections: [
          {
            name: "Cases",
            labels: {},
            attrs: attributes,
          },
        ],
        records,
      },
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

  function clearErrorAnd(f: () => void) {
    return () => {
      setError("");
      f();
    };
  }

  return chosenSpreadsheet !== null ? (
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
        <h3>Range to Import</h3>
        <input
          type="radio"
          id="all"
          checked={!useCustomRange}
          onClick={clearErrorAnd(() => setUseCustomRange(false))}
        />
        <label htmlFor="all">All values</label>
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

      <div id="submit-buttons" className="input-group">
        <button onClick={importSheet}>Import</button>
        <button onClick={cancelImport}>Cancel</button>
      </div>

      {error !== "" && (
        <div className="error">
          <p>{error}</p>
        </div>
      )}
    </>
  ) : (
    <></>
  );
}
