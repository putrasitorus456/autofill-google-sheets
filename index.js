const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const fs = require('fs');
const { google } = require('googleapis');

const keyFile = './integration-project-425209-ab0777c09e66.json';

const app = express();
app.use(bodyParser.json());
app.use(cors());

async function accessSpreadsheet(values) {
    const auth = new google.auth.GoogleAuth({
        keyFile,
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const client = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: client });

    const spreadsheetId = '1edRT0fpulZz0lfZFgABKqIF07tN4inOknHTtHEb0_M8';

    const getResponse = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: 'Sheet1',
    });
    const numRows = getResponse.data.values ? getResponse.data.values.length : 0;

    const range = `Sheet1!B${numRows + 1}:E${numRows + values.length}`;

    const resource = {
        values,
    };

    try {
        const result = await sheets.spreadsheets.values.update({
            spreadsheetId,
            range,
            valueInputOption: 'RAW',
            resource,
        });
        return result.data;
    } catch (err) {
        console.error(err);
        throw err;
    }
}

app.post('/updateSpreadsheet', async (req, res) => {
    try {
        const values = req.body.values;
        const result = await accessSpreadsheet(values);
        res.json(result);
    } catch (err) {
        res.status(500).send(err.toString());
    }
});

app.listen(9090, () => {
    console.log('Server is listening on port 9090');
});