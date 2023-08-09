const express = require('express');
const msal = require('@azure/msal-node');
const cors = require('cors');

const app = express();
const port = 3002;

app.use(cors());
app.use(express.json());

let credentials = {};
let siteName = {};

// Get the credentials of ArcGIS server
app.post('/set-credentials', (req, res) => {
  credentials = req.body;
  res.send({ status: 'Credentials set' });
});

// Get the site Name of ArcGIS server
app.post('/set-siteName', (req, res) => {
  siteName = req.body;
  res.send({ status: 'Name set' });
});

// Get token
app.get('/token', async (req, res) => {
  try {
    const cca = new msal.ConfidentialClientApplication({
      auth: {
        clientId: credentials.client_id,
        authority: "https://login.microsoftonline.com/" + credentials.tenant_id,
        clientSecret: credentials.client_secret,
      },
    });

    const response = await cca.acquireTokenByClientCredential({
      scopes: ["https://graph.microsoft.com/.default"],
    });
    console.log("Token Generated Succesfully");
    res.send(response);
  } catch (err) {
    console.log(err);
    res.status(500).send(err);
  }
});

// Get data
app.get('/getSites', async (req, res) => {
  let sitesData = [];
  let siteId;
  try {
    const token = req.headers.authorization;

    const sitesResponse = await fetch('https://graph.microsoft.com/v1.0/sites', {
      headers: {
        'Authorization': token
      }
    });

    const data = await sitesResponse.json();
    sitesData = data.value;

    sitesData.forEach(site => {
      if (siteName.site_name == site.name) {
        siteId = site.id;
        siteWebUrl = site.webUrl;
      }
    });
    if (siteId != undefined) {
      res.json({ siteId, siteWebUrl });
    } else {
      res.send(null);
    }

  } catch (err) {
    console.log(err);
    res.status(500).send(err);
  }
});

function getValueInsideBraces(str) {
  const match = str.match(/{(.*?)}/);
  if (match) {
    return match[1];
  }
  return null;
}

app.get('/display-ff', async (req, res) => {
  try {
    const token = req.headers.authorization;
    const siteId = req.headers.siteid;

    const filesResponse = await fetch('https://graph.microsoft.com/v1.0/sites/' + siteId + '/lists/Documents/items?expand=fields', {
      headers: {
        'Content-Type': 'application/json',
        'Prefer': 'apiversion=2.1',
        'Authorization': token
      }
    });
    const data = await filesResponse.json();
    res.send(data);

  } catch (err) {
    console.log(err);
    res.status(500).send(err);
  }
});

app.listen(port, () => {
  console.log(`Server running on http://localhost:${port}`);
});
