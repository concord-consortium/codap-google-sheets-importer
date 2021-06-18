import React, { useState, useEffect } from "react";
import "./App.css";
import { initializePlugin, createTableWithDataset } from "codap-phone";
import { useInput } from "./hooks";
import { getDataFromSheet } from "./util";

const CLIENT_ID =
  "756054504415-rf57dsh2mt5vqk1sovptbpopcacctred.apps.googleusercontent.com";
const DISCOVERY_DOCS = [
  "https://sheets.googleapis.com/$discovery/rest?version=v4",
];
const scope = "https://www.googleapis.com/auth/spreadsheets.readonly";

const PLUGIN_WIDTH = 500;
const PLUGIN_HEIGHT = 700;

export default function App() {
  // Load Google APIs upon mounting
  useEffect(() => {
    initializePlugin("Google Sheets Importer", PLUGIN_WIDTH, PLUGIN_HEIGHT);
    gapi.load("client:auth2:picker", onClientLoad);
  }, []);

  const [chosenSpreadsheet, setChosenSpreadsheet] =
    useState<Required<gapi.client.sheets.Spreadsheet> | null>(null);
  const [chosenSheet, chosenSheetChange, setChosenSheet] =
    useInput<string, HTMLSelectElement>("");
  const [useHeader, setUseHeader] = useState<boolean>(false);
  const [useCustomRange, setUseCustomRange] = useState<boolean>(false);
  const [customRange, customRangeChange, setCustomRange] =
    useInput<string, HTMLInputElement>("");

  function resetState() {
    setChosenSpreadsheet(null);
    setChosenSheet("");
    setUseHeader(false);
    setUseCustomRange(false);
    setCustomRange("");
  }

  // Authenticate user so we can read their spreadsheets. This will pop up
  // a Google login window.
  async function onClientLoad() {
    gapi.client.init({
      discoveryDocs: DISCOVERY_DOCS,
      clientId: CLIENT_ID,
      scope,
    });

    await loginAndCreatePicker();
  }

  async function loginAndCreatePicker() {
    const GoogleAuth = gapi.auth2.getAuthInstance();
    if (GoogleAuth.isSignedIn.get()) {
      const currentUser = GoogleAuth.currentUser.get();
      createPicker(currentUser.getAuthResponse().access_token);
    } else {
      const response = await GoogleAuth.signIn();
      createPicker(response.getAuthResponse().access_token);
    }
  }

  function createPicker(token: string) {
    const picker = new google.picker.PickerBuilder()
      .setOAuthToken(token)
      .addView(google.picker.ViewId.SPREADSHEETS)
      .enableFeature(google.picker.Feature.NAV_HIDDEN)
      .hideTitleBar()
      .setCallback(makePickerCallback(token))
      .build();
    picker.setVisible(true);
  }

  function makePickerCallback(token: string) {
    return async (response: google.picker.ResponseObject) => {
      if (
        response[google.picker.Response.ACTION] === google.picker.Action.PICKED
      ) {
        const doc = response[google.picker.Response.DOCUMENTS][0];
        console.log(doc);
        const sheet = (
          await gapi.client.sheets.spreadsheets.get({ spreadsheetId: doc.id })
        ).result;
        console.log(sheet);
        setChosenSpreadsheet(sheet as Required<gapi.client.sheets.Spreadsheet>);

        // Set first sheet as chosen
        if (sheet.sheets && sheet.sheets.length > 0) {
          setChosenSheet(sheet.sheets[0].properties?.title as string);
        }
      }
    };
  }

  async function importSheet() {
    if (chosenSpreadsheet === null) {
      console.log("No chosen spreadsheet!");
      return;
    }

    const range = useCustomRange ? customRange : chosenSheet;

    if (range === "") {
      console.log("No range");
      console.log(customRange);
      console.log(chosenSheet);
      return;
    }

    const data = await getDataFromSheet(chosenSpreadsheet.spreadsheetId, range);

    if (data.length === 0) {
      console.log("No data");
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
          onClick={() => setUseCustomRange(false)}
        />
        <label htmlFor="all">All values</label>
        <br />
        <input type="radio" checked={useCustomRange} />
        <input
          type="text"
          value={customRange}
          onFocus={() => setUseCustomRange(true)}
          onChange={customRangeChange}
        />
      </div>

      <div className="input-group">
        <input
          type="checkbox"
          id="useHeader"
          onChange={toggleHeader}
          checked={useHeader}
        />
        <label htmlFor="useHeader">Use first row as column names</label>
      </div>
      <button onClick={importSheet}>Import</button>
      <button onClick={cancelImport}>Cancel</button>
    </>
  ) : (
    <></>
  );
}
