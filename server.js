const express = require('express');
const msal = require('@azure/msal-node');
const cors = require('cors');

const app = express();
const port = 3002;

app.use(cors());
app.use(express.json());

let credentials = {};
let siteName = {};

// Get the credentials of ArcGIS server.
app.post('/set-credentials', (req, res) => {
  credentials = req.body;
  res.send({ status: 'Credentials set' });
});

// Get the site Name of ArcGIS server.
app.post('/set-siteName', (req, res) => {
  siteName = req.body;
  res.send({ status: 'Name set' });
});

// Function to check if a token is expired.
function isTokenExpired(token) {
  const currentTime = new Date();
  return currentTime >= token.expiresOn;
}

// Initialize cachedToken variable.
let cachedToken = null;

// Function to get a valid token.
async function getValidToken() {
  if (!cachedToken || isTokenExpired(cachedToken)) {
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

    cachedToken = response;
  }

  return cachedToken;
}

async function termMiddleware(req, res, next) {
  try {
    const token = await getValidToken();
    const siteId = req.headers.siteid;
    let termGroupId;
    let termSetId;

    //------------------------TERM GROUP---------------------------

    const termGroupsResponse = await fetch('https://graph.microsoft.com/v1.0/sites/' + siteId + '/termStore/groups', {
      headers: {
        'Content-Type': 'application/json',
        'Prefer': 'apiversion=2.0',
        'Authorization': token.accessToken
      }
    });
    const dataGroups = await termGroupsResponse.json();

    const foundGroup = dataGroups.value.find(termGroup => termGroup.displayName === "GeoTag");

    if (foundGroup) {
      console.log("Found TermGroup GeoTag!");
      termGroupId = foundGroup.id;
    } else {
      console.log("Creating TermGroup GeoTag...");
      const urlencoded = new URLSearchParams();
      urlencoded.append("displayName", "GeoTag");

      const createTermGroupsResponse = await fetch('https://graph.microsoft.com/v1.0/sites/' + siteId + '/termStore/groups', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          'Prefer': 'apiversion=2.0',
          'Authorization': token.accessToken
        },
        body: urlencoded,
        redirect: 'follow'
      });

      const dataCreateGroups = await createTermGroupsResponse.json();
      termGroupId = dataCreateGroups.id;
    }

    //------------------------TERM SET---------------------------

    const termSetsResponse = await fetch('https://graph.microsoft.com/v1.0/sites/' + siteId + '/termStore/groups/' + termGroupId + '/sets', {
      headers: {
        'Content-Type': 'application/json',
        'Prefer': 'apiversion=2.0',
        'Authorization': token.accessToken
      }
    });
    const dataSets = await termSetsResponse.json();

    const foundSet = dataSets.value.find(termSet =>
      termSet.localizedNames.some(localizedName => localizedName.name === "GeoTag")
    );

    if (foundSet) {
      console.log("Found TermSet GeoTag!");
      termSetId = foundSet.id;
    } else {
      console.log("Creating TermSet GeoTag...");

      const createTermSetsResponse = await fetch('https://graph.microsoft.com/v1.0/sites/' + siteId + '/termStore/sets', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Prefer': 'apiversion=2.0',
          'Authorization': token.accessToken
        },
        body: JSON.stringify({
          "parentGroup": {
            "id": termGroupId
          },
          "description": "GeoTag",
          "localizedNames": [
            {
              "languageTag": "en-US",
              "name": "GeoTag"
            }
          ]
        }),
        redirect: 'follow'
      });

      const dataCreateSets = await createTermSetsResponse.json();
      termSetId = dataCreateSets.id;
    }

    req.termData = {
      termGroupId: termGroupId,
      termSetId: termSetId
    };

    next();
  } catch (err) {
    console.log(err);
    res.status(500).send(err);
  }
}

// Get token.
app.get('/token', async (req, res) => {
  try {
    const token = await getValidToken();
    res.send(token);
  } catch (err) {
    console.log(err);
    res.status(500).send(err);
  }
});

// Get data.
app.get('/getSites', async (req, res) => {
  let sitesData = [];
  let siteId;
  try {
    const token = await getValidToken();

    const sitesResponse = await fetch('https://graph.microsoft.com/v1.0/sites', {
      headers: {
        'Authorization': token.accessToken,
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

app.get('/display-ff', async (req, res) => {
  try {
    const token = await getValidToken();
    const siteId = req.headers.siteid;
    let folderId = req.headers.folderid;

    if (folderId === "null") {
      folderId = "root";
    }

    const filesResponse = await fetch('https://graph.microsoft.com/v1.0/sites/' + siteId + '/drive/items/' + folderId + '/children?&select=id,eTag,package&expand=listitem(expand=fields(select=FileLeafRef,DocIcon,GeoTag,ContentType))', {
      headers: {
        'Content-Type': 'application/json',
        'Prefer': 'apiversion=2.0',
        'Authorization': token.accessToken
      }
    });
    const data = await filesResponse.json();
    res.send(data);
  } catch (err) {
    console.log(err);
    res.status(500).send(err);
  }
});

app.patch('/addTag', termMiddleware, async (req, res) => {
  try {
    const token = await getValidToken();
    const siteId = req.headers.siteid;
    const tag = req.body.tag;
    const fileTags = req.body.fileTags;
    const fileId = req.body.fileId;

    //------------------------TERM---------------------------

    const termResponse = await fetch('https://graph.microsoft.com/v1.0/sites/' + siteId + '/termStore/sets/' + req.termData.termSetId + '/terms', {
      headers: {
        'Content-Type': 'application/json',
        'Prefer': 'apiversion=2.0',
        'Authorization': token.accessToken
      }
    });
    const dataTerms = await termResponse.json();

    const foundTerm = dataTerms.value.find(term =>
      term.labels.some(label => label.name.toLowerCase() === tag.toLowerCase())
    );

    if (foundTerm) {
      console.log("Found Term " + tag + "!");
      termId = foundTerm.id;
    } else {
      console.log("Creating Term " + tag + "...");

      const createTermResponse = await fetch('https://graph.microsoft.com/v1.0/sites/' + siteId + '/termStore/sets/' + req.termData.termSetId + '/children', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Prefer': 'apiversion=2.0',
          'Authorization': token.accessToken
        },
        body: JSON.stringify({
          "labels": [
            {
              "languageTag": "en-US",
              "name": tag,
              "isDefault": true
            }
          ]
        }),
        redirect: 'follow'
      });

      const dataCreateTerms = await createTermResponse.json();
      termId = dataCreateTerms.id;
    }

    //---------------------ADD TAG---------------------------

    const oldTags = fileTags.map(tag => tag.label + "|" + tag.termGuid + ";");

    const createTagResponse = await fetch('https://graph.microsoft.com/v1.0/sites/' + siteId + '/lists/Documents/items/' + fileId + '/fields', {
      method: 'PATCH',
      headers: {
        'Content-Type': 'application/json',
        'Prefer': 'apiversion=2.0',
        'Authorization': token.accessToken
      },
      body: JSON.stringify({
        "o5c3b196e2d0422495d173d6e391d21f": oldTags + tag + "|" + termId
      }),
      redirect: 'follow'
    });

    const dataCreateTag = await createTagResponse.json();
    //console.log(dataCreateTag)

    res.send({ label: tag, termGuid: termId });

  } catch (err) {
    console.log(err);
    res.status(500).send(err);
  }
});

app.get('/seeTaggedFiles', termMiddleware, async (req, res) => {

  try {
    const token = await getValidToken();
    const siteId = req.headers.siteid;
    const nameTag = req.headers.nametag;

    const termResponse = await fetch('https://graph.microsoft.com/v1.0/sites/' + siteId + '/termStore/sets/' + req.termData.termSetId + '/terms', {
      headers: {
        'Content-Type': 'application/json',
        'Prefer': 'apiversion=2.0',
        'Authorization': token.accessToken
      }
    });
    const dataTerms = await termResponse.json();
    //console.log(dataTerms)

    const foundTerm = dataTerms.value.find(term =>
      term.labels.some(label => label.name.toLowerCase() === nameTag.toLowerCase())
    );

    if (foundTerm) {
      console.log("Found Term " + nameTag + "!");
      termId = foundTerm.id;
      console.log(termId)
      const taggedFilesResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root/search(q='${termId}')`, {
        headers: {
          'Content-Type': 'application/json',
          'Prefer': 'apiversion=2.0',
          'Authorization': token.accessToken
        }
      });
      const dataTaggedFiles = await taggedFilesResponse.json();
      //console.log(dataTaggedFiles);
      res.send(dataTaggedFiles);

    } else {
      console.log(nameTag + " not found!");
      res.status(404).send({ message: "Term not found!" });
    }

  } catch (err) {
    console.log(err);
    res.status(500).send(err);
  }

});

app.get('/seeDataTaggedFile', async (req, res) => {
  try {
    const token = await getValidToken();
    const siteId = req.headers.siteid;
    let fileId = req.headers.fileid;

    const filesResponse = await fetch('https://graph.microsoft.com/v1.0/sites/' + siteId + '/drive/items/' + fileId + '?&select=id,name&expand=listitem(expand=fields(select=GeoTag))', {
      headers: {
        'Content-Type': 'application/json',
        'Prefer': 'apiversion=2.0',
        'Authorization': token.accessToken
      }
    });
    const data = await filesResponse.json();
    console.log(data)
    res.send(data);
  } catch (err) {
    console.log(err);
    res.status(500).send(err);
  }
});

app.listen(port, () => {
  console.log(`Server running on http://localhost:${port}`);
});