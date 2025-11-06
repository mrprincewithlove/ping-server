

const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');


const credentials = JSON.parse(process.env.CREDENTIALS);


const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

async function mergeAndWriteSheet(spreadsheetId, sheetTitle, newData) {
    const client = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: client });
    let isSheetExists = false;
    console.log(spreadsheetId);
    console.log(sheetTitle);
    // 1. Add new sheet with today's date
    try {
        // Get existing sheet names
        const res = await sheets.spreadsheets.get({ spreadsheetId });
        const sheetExists = res.data.sheets.some(
            s => s.properties.title === sheetTitle
        );
        if (!sheetExists) {
            isSheetExists = false;
            await sheets.spreadsheets.batchUpdate({
                spreadsheetId,
                requestBody: {
                    requests: [{
                        addSheet: {
                            properties: {
                                title: sheetTitle,
                            },
                        },
                    }],
                },
            });
            console.log(`Sheet "${sheetTitle}" created`);
        } else {
            isSheetExists = true;
            console.log(`Sheet "${sheetTitle}" already exists`);
        }
    } catch (err) {
        if (err.errors?.[0]?.reason === 'duplicate') {
            console.warn(`Sheet '${sheetTitle}' already exists.`);
        } else {
            console.error("Error creating sheet:", err.message);
        }
    }

    try {
        // 1. Try to get existing data
        let existingData = [];
        try {
            const res = await sheets.spreadsheets.values.get({
                spreadsheetId,
                range: `${sheetTitle}!A:J`,
            });
            existingData = res.data.values || [];
        } catch (err) {
            if (err.code !== 404) throw err;
            // Sheet doesn't exist yet - that's fine
        }

        // 2. Process data
        const headers = ["ID", "Key", "Hostname", "IP", "Port", "Info", "Info2", "Country", "Short", "Location"];

        // Create Map of existing data (Key -> Full Row)
        const existingMap = new Map();
        if (existingData.length > 0) {
            // Skip header row if exists
            const startRow = existingData[0][0] === "ID" ? 1 : 0;
            for (let i = startRow; i < existingData.length; i++) {
                const row = existingData[i];
                if (row[1]) existingMap.set(row[1], row); // row[1] = Key column
            }
        }

        // Merge with new data (new entries overwrite existing ones)
        newData.forEach(d => {
            const key = `${d.ip}:${d.port}`;
            existingMap.set(key, [
                d.id || '',
                key,
                d.hostname || '',
                d.ip || '',
                d.port || '',
                d.info || '',
                d.info2 || '',
                d.location?.country || '',
                d.location?.short || '',
                d.location?.name || ''
            ]);
        });

        // Convert back to array
        const mergedData = [headers, ...Array.from(existingMap.values())];

        // 3. Clear and rewrite entire sheet (more efficient than individual updates)
        await sheets.spreadsheets.values.clear({
            spreadsheetId,
            range: `${sheetTitle}!A:J`,
        });

        await sheets.spreadsheets.values.update({
            spreadsheetId,
            range: `${sheetTitle}!A1`,
            valueInputOption: 'RAW',
            requestBody: { values: mergedData }
        });

        console.log(`Merged ${newData.length} new records into ${existingMap.size} total records`);

    } catch (err) {
        console.error("Error in mergeAndWriteSheet:", err.message);
        throw err;
    }
}


module.exports = { mergeAndWriteSheet };

