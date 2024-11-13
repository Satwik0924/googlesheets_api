import { google } from 'googleapis';
import express from 'express';
import path from 'path';
import fs from 'fs';
import bodyParser from 'body-parser';
import cors from 'cors'; // Import CORS middleware

// Get the current directory path using import.meta.url
const __dirname = path.dirname(new URL(import.meta.url).pathname);

// Load OAuth2 client credentials from the credentials.json file
const credentialsPath = path.join(__dirname, 'credentials.json');
const credentials = JSON.parse(fs.readFileSync(credentialsPath));

// Extract client ID, client secret, and redirect URIs from credentials.json
const { client_id, client_secret, redirect_uris } = credentials.web;

// Initialize OAuth2 client using values from credentials.json
const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

// Load tokens from tokens.json if they exist
const tokenPath = path.join(__dirname, 'tokens.json');
let tokensData = {};
if (fs.existsSync(tokenPath)) {
  tokensData = JSON.parse(fs.readFileSync(tokenPath));
  console.log('Tokens loaded successfully from tokens.json!');
} else {
  console.log('Tokens file not found. Please authenticate first.');
}

// Initialize Express app
const app = express();

// Enable CORS
app.use(cors());

// Middleware to parse JSON data
app.use(bodyParser.json());

// Route to initiate the OAuth flow
app.get('/auth', (req, res) => {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: [
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/drive',
    ],
  });
  res.redirect(authUrl); // Redirect the user to Google's OAuth consent screen
});

// Route to handle the OAuth 2.0 callback
app.get('/oauth2callback', async (req, res) => {
  const code = req.query.code;

  try {
    const { tokens } = await oAuth2Client.getToken(code);
    oAuth2Client.setCredentials(tokens);

    // Retrieve the authenticated user's email
    const oauth2 = google.oauth2({ version: 'v2', auth: oAuth2Client });
    const userInfo = await oauth2.userinfo.get();
    const userEmail = userInfo.data.email;

    if (userEmail) {
      // Save tokens under the user's email in tokens.json
      tokensData[userEmail] = tokens;
      fs.writeFileSync(tokenPath, JSON.stringify(tokensData, null, 2));
      console.log(`Tokens saved successfully for user: ${userEmail}`);
    }

    res.send('Authentication successful! You can now append data to your Google Sheets.');
  } catch (error) {
    console.error('Error retrieving access token:', error);
    res.status(500).send('Authentication failed.');
  }
});

// Route to append data to an existing Google Sheet
app.post('/append-data-to-existing-sheet', async (req, res) => {
  const { spreadsheetId, data, additionalFieldsOrder, userEmail } = req.body; // Use the provided userEmail

  // Check if tokens exist for the given userEmail
  if (!userEmail || !tokensData[userEmail]) {
    return res.status(400).json({ message: `No tokens found for user ${userEmail}. Please authenticate first.` });
  }

  // Set the credentials for the OAuth2 client
  oAuth2Client.setCredentials(tokensData[userEmail]);

  const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });

  // Prepare the data to be appended under the headers in the next row
  const mandatoryFields = ['name', 'email', 'phone'];
  const rowData = [];

  // Ensure that the mandatory fields exist in the request body
  for (const field of mandatoryFields) {
    if (!data[field]) {
      return res.status(400).json({ message: `${field} is a mandatory field.` });
    }
    rowData.push(data[field]); // Add mandatory fields to the rowData
  }

  // Handle additional fields in the order specified by the client
  if (Array.isArray(additionalFieldsOrder)) {
    for (const field of additionalFieldsOrder) {
      if (field.startsWith('additional_col') && data[field] !== undefined) {
        rowData.push(data[field]); // Add additional fields in the specified order
      }
    }
  }

  const appendRequest = {
    spreadsheetId: spreadsheetId, // Use the provided sheet ID
    range: 'Sheet1!A2', // Start appending from the second row (assuming headers are in the first row)
    valueInputOption: 'RAW',
    resource: {
      values: [rowData],
    },
  };

  try {
    // Append the data to the sheet
    await sheets.spreadsheets.values.append(appendRequest);
    console.log('Data appended to the sheet.');
    res.json({ message: 'Data appended successfully to the Google Sheet!' });
  } catch (error) {
    console.error('Error appending data:', error);
    res.status(500).json({ message: 'Failed to append data to the Google Sheet.' });
  }
});

// Start the server
app.listen(3000, () => {
  console.log('Server running on http://localhost:3000');
});
