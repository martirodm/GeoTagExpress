const express = require('express');
const msal = require('@azure/msal-node');
const cors = require('cors');

const app = express();
const port = 3002;

app.use(cors());
app.use(express.json());

let credentials = {};

// Get the credentials of ArcGIS server
app.post('/set-credentials', (req, res) => {
  credentials = req.body;
  res.send({ status: 'Credentials set' });
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

    res.send(response);
  } catch (err) {
    console.log(err);
    res.status(500).send(err);
  }
});

// Get data
app.get('/data', async (req, res) => {
  try {
    const token = req.headers.authorization;
    console.log(token)

    const apiResponse = await fetch('https://graph.microsoft.com/v1.0/sites', {
      headers: {
        'Authorization': token
      }
    });

    const jsonData = await apiResponse.json();
    res.send(jsonData);

  } catch (err) {
    console.log(err);
    res.status(500).send(err);
  }
});

app.listen(port, () => {
  console.log(`Server running on http://localhost:${port}`);
});
