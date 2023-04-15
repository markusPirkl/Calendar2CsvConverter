const fs = require("fs").promises;
const path = require("path");
const process = require("process");
const { authenticate } = require("@google-cloud/local-auth");
const { google } = require("googleapis");

// If modifying these scopes, delete token.json.
const SCOPES = ["https://www.googleapis.com/auth/calendar.readonly"];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = path.join(process.cwd(), "token.json");
const CREDENTIALS_PATH = path.join(process.cwd(), "credentials.json");

/**
 * Reads previously authorized credentials from the save file.
 *
 * @return {Promise<OAuth2Client|null>}
 */
async function loadSavedCredentialsIfExist() {
  try {
    const content = await fs.readFile(TOKEN_PATH);
    const credentials = JSON.parse(content);
    return google.auth.fromJSON(credentials);
  } catch (err) {
    return null;
  }
}

/**
 * Serializes credentials to a file compatible with GoogleAUth.fromJSON.
 *
 * @param {OAuth2Client} client
 * @return {Promise<void>}
 */
async function saveCredentials(client) {
  const content = await fs.readFile(CREDENTIALS_PATH);
  const keys = JSON.parse(content);
  const key = keys.installed || keys.web;
  const payload = JSON.stringify({
    type: "authorized_user",
    client_id: key.client_id,
    client_secret: key.client_secret,
    refresh_token: client.credentials.refresh_token,
  });
  await fs.writeFile(TOKEN_PATH, payload);
}

/**
 * Load or request or authorization to call APIs.
 *
 */
async function authorize() {
  let client = await loadSavedCredentialsIfExist();
  if (client) {
    return client;
  }
  client = await authenticate({
    scopes: SCOPES,
    keyfilePath: CREDENTIALS_PATH,
  });
  if (client.credentials) {
    await saveCredentials(client);
  }
  return client;
}

/**
 * Lists the next 10 events on the user's primary calendar.
 * @param {google.auth.OAuth2} auth An authorized OAuth2 client.
 */
async function listEvents(auth) {
  const calendar = google.calendar({ version: "v3", auth });
  // Events
  var events = [];
  var currentEventsCall = await calendar.events.list({
    calendarId: "7jr69qmahh0ubolt3uplckufcg@group.calendar.google.com",
    timeMin: new Date("2022-03-31").toISOString(),
    timeMax: new Date("2023-04-01").toISOString(),
    singleEvents: true,
    orderBy: "starttime",
  });
  console.log("no events prior filter", events.length);
  events = currentEventsCall.data.items.filter(isValidEvent);
  console.log("no events post filter", events.length);
  var nextPageToken = currentEventsCall.data.nextPageToken;
  while (nextPageToken != undefined) {
    currentEventsCall = await calendar.events.list({
      calendarId: "7jr69qmahh0ubolt3uplckufcg@group.calendar.google.com",
      timeMin: new Date("2022-03-31").toISOString(),
      timeMax: new Date("2023-04-01").toISOString(),
      pageToken: nextPageToken,
      singleEvents: true,
      orderBy: "starttime",
    });
    nextPageToken = currentEventsCall.data.nextPageToken;
    events.push(...currentEventsCall.data.items.filter(isValidEvent));
  }
  console.log(
    "No of Events For tax year 2022: Elisa und Markus Calendar",
    events.length
  );

  // Write csv file
  const fs = require("fs"),
    csv = require("csv-stringify");

  // (C) CREATE CSV FILE
  csv.stringify(
    events,
    {
      header: true,
      columns: {
        summary: "Summary",
        description: "Description",
        location: "Location",
        start: "Start",
        end: "End",
      },
    },
    (err, output) => {
      fs.writeFileSync("events_elisa.csv", output);
      console.log("OK");
    }
  );
}

function isValidEvent(event) {
  if (event.summary?.includes("Geburtstag")) {
    return false;
  } else if (event.summary?.includes("Birthday")) {
    return false;
  }

  console.log("creator email", event.creator?.email);
  if (event.creator?.email !== "elisa.mussemann@googlemail.com") {
    return false;
  }

  return true;
}

authorize().then(listEvents).catch(console.error);
