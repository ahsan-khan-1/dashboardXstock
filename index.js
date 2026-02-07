const express = require("express"); // Tells node go find the Express package inside node_modules and load it

const { google } = require("googleapis"); // Loading GoogleApis

const bodyparser = require("body-parser"); // Loading body parser

const app = express() // Calls the express function we just imported and creates an Express application instance. This 'app' is your server, which you can use to define routes, handle requests, and start listening for connections.

app.use(express.static("public"));

// For form submissions
app.use(bodyparser.urlencoded({ extended:true, limit: "10mb" }));
// Make sure you can parse JSON
app.use(express.json({ limit: "10mb" }));
app.set("view engine", "ejs") // Tells Express that we want to use EJS as our template engine. Now, when we call res.render("file"), Express will look for a file named "file.ejs" in the "views" folder and render it as HTML.

// ---------- Google Sheets setup (GLOBAL, runs once) ----------
const credentials = {
    type: process.env.type,
    project_id: process.env.project_id,
    private_key_id: process.env.private_key_id,
    private_key: process.env.private_key.replace(/\\n/g, '\n'),
    client_email: process.env.client_email,
    client_id: process.env.client_id,
    auth_uri: process.env.auth_uri,
    token_uri: process.env.token_uri,
    auth_provider_x509_cert_url: process.env.auth_provider_x509_cert_url,
    client_x509_cert_url: process.env.client_x509_cert_url,
    universe_domain: process.env.universe_domain,
};

const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

let googleSheets;
const spreadsheetId = "1WoWMm6XHkq0VVW6r5OFRqglN5atdPpCQ5fKxnjcDB7s";

(async () => {
    const client = await auth.getClient();
    googleSheets = google.sheets({ version: "v4", auth: client });
})();

app.get ("/", async (req, res) => {
/** Tells Express: "When someone visits the '/' route (the homepage), run this function."
 * 'req' is the request object (info about the visitor's request)
 * 'res' is the response object (used to send something back to the visitor)
 * 'async' means we can use 'await' inside this function to wait for things like database or Google Sheets data before sending a response.
*/
    /* Read rows from spreadsheet from New_Unsorted_Items
    const getNewItems = await googleSheets.spreadsheets.values.get({
        auth,
        spreadsheetId,
        range: "New_Unsorted_Items!A:M", // get all from this sheet from col A to M
    });*/

    // There is better way to do API calls in which all are done parallel
    const sheets = [
        "New_Unsorted_Items!A:G",
        "Need_To_Make!A:G",
        "Checked_Items!A:G",
        "Returning_Items!A:G",
        "Parts_Items!A:M",
        "AI_Gen_Listing!A:M",
    ];

    const [newItemsData, needToMakeData, checkedData, 
        returnedData, partsData,aiGenListingData] = await Promise.all(
            sheets.map(range => googleSheets.spreadsheets.values.get({
                auth, // basically saying auth: auth 
                spreadsheetId, // spreadsheetId: spreadsheetId
                range // range: range
            }))
        )
    

    const newItems = newItemsData.data.values ; //Get the items or else null this prevents the code from breaking
    const needToMake = needToMakeData.data.values || [];
    const checked = checkedData.data.values || [];
    const returned = returnedData.data.values || [];
    const aiGenListing = aiGenListingData.data.values || [];
    const partItems = partsData.data.values || [];
    
    
    // Write row(s) to spreadsheet
    /*await googleSheets.spreadsheets.values.append({
        auth,
        spreadsheetId,
        range: "Practice1!A:B"
    })*/

    res.render("index", {newItems, needToMake, checked, returned, partItems,aiGenListing});
});


/*
await googleSheets.spreadsheets.values.append({
    auth,
    spreadsheetId,
    range: "Sheet1!D4",
    valueInputOption: "RAW",
    resource: {
        values: [

        ]
    }
})
*/
app.post("/update/new-items", async (req, res) => {
    try {
        // req.body = [ { range: "Sheet!A2", value: "Apple" }, ... ]
        const updates = req.body.map(u => ({
            range: u.range,
            values: [[u.value]],
        }));

        await googleSheets.spreadsheets.values.batchUpdate({
            spreadsheetId,
            resource: {
                valueInputOption: "RAW",
                data: updates,
            },
        });

        res.json({ status: "ok" });
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: "Something went wrong" });
    }
});

app.post("/push/new-items", async (req, res) => {
    const rows = req.body; // array of { rowIndex, data, status }

    try {
        // First, sort rows descending by rowIndex so deleting rows doesn't shift indexes
        const sortedRows = rows
            .filter(r => r.status !== "select") // skip unselected
            .sort((a, b) => b.rowIndex - a.rowIndex);

        for (const row of sortedRows) {
            let targetSheet;
            switch(row.status) {
                case "Need To Make":
                    targetSheet = "Need_To_Make";
                    break;
                case "Return":
                    targetSheet = "Returning_Items";
                    break;
                case "Parts":
                    targetSheet = "Parts_Items";
                    break;
                case "Brand New":
                case "Used As New":
                case "Used":
                    targetSheet = "Checked_Items";
                    break;
                default:
                    continue;
            }

            // 1️⃣ Append row to target sheet
            await googleSheets.spreadsheets.values.append({
                auth,
                spreadsheetId,
                range: `${targetSheet}!A:Z`,
                valueInputOption: "RAW",
                resource: { values: [row.data] }
            });

            // 2️⃣ Delete row from original sheet
            // Note: `sheetId` is numeric, not sheet name. Replace with your sheet's gid
            const sheetId = 0; 

            await googleSheets.spreadsheets.batchUpdate({
                auth,
                spreadsheetId,
                resource: {
                    requests: [
                        {
                            deleteDimension: {
                                range: {
                                    sheetId: sheetId,
                                    dimension: "ROWS",
                                    startIndex: row.rowIndex - 1, // zero-based
                                    endIndex: row.rowIndex        // exclusive
                                }
                            }
                        }
                    ]
                }
            });
        }

        res.json({ success: true });
    } catch (err) {
        console.error(err);
        res.status(500).json({ success: false, error: err.message });
    }
});
const PORT = process.env.PORT || 8080;

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});


