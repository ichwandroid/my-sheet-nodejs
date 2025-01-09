const express = require('express');
const path = require('path');
const { google } = require('googleapis');

const app = express();

const port = parseInt(process.env.PORT) || process.argv[3] || 3000;

app.use(express.static(path.join(__dirname, 'public')))
  .set('views', path.join(__dirname, 'views'))
  .set('view engine', 'ejs');

const sheet = google.sheets('v4');
// Suggested code may be subject to a license. Learn more: ~LicenseLog:85271322.
const auth = new google.auth.GoogleAuth({
  keyFile: '../connect.json',
  scopes: 'https://www.googleapis.com/auth/spreadsheets'
});

app.get('/', (req, res) => {
  res.render('index');
});

app.get('/profile-sekolah', async (req, res) => {
  const authClient = await auth.getClient();
  const spreadsheetId = '1NB38zBlpDdLvtsE_YDzLhXT8vtMQ6a-FUSPeTaec-N0';
  const range = 'SUM PAIBP!A2:R';

  const response = await sheet.spreadsheets.values.get({
    auth: authClient,
    spreadsheetId,
    range,
  });

  const data = response.data.values;
  res.render('profile', { data });
})

app.listen(port, () => {
  console.log(`Listening on http://localhost:${port}`);
})