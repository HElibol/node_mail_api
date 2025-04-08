require('dotenv').config();
const express = require('express');
const axios = require('axios');
const msal = require('@azure/msal-node');
const cors = require('cors');

const app = express();
const PORT = 3000;

app.use(cors());



// MSAL config
const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    clientSecret: process.env.CLIENT_SECRET,
    authority: process.env.AUTHORITY,
  },
};
const cca = new msal.ConfidentialClientApplication(msalConfig);

// Token alma fonksiyonu
async function getToken() {
  const tokenRequest = {
    scopes: ["https://graph.microsoft.com/.default"],
  };
  const response = await cca.acquireTokenByClientCredential(tokenRequest);
  return response.accessToken;
}

// Tüm kullanıcıları getir
app.get('/users', async (req, res) => {
  try {
    const token = await getToken();
    const graphRes = await axios.get('https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail', {
      headers: { Authorization: `Bearer ${token}` },
    });
    res.json(graphRes.data);
  } catch (err) {
    console.error(err.response?.data || err.message);
    res.status(500).send('Kullanıcılar alınamadı.');
  }
});

// Kullanıcının son 10 mailini getir
app.get('/users/:userId/emails', async (req, res) => {
    const userId = req.params.userId;
    try {
      const token = await getToken();
      const url = `https://graph.microsoft.com/v1.0/users/${userId}/mailFolders/Inbox/messages?$top=10&$orderby=receivedDateTime desc&$select=subject,from,receivedDateTime,bodyPreview`;
      const graphRes = await axios.get(url, {
        headers: { Authorization: `Bearer ${token}` },
      });
  


      const formatted = graphRes.data.value
      .filter(msg => msg.from?.emailAddress?.address !== "noreply@emeaemail.teams.microsoft.com")
      .map((msg) => ({
        fromName: msg.from?.emailAddress?.name,
        fromEmail: msg.from?.emailAddress?.address,
        subject: msg.subject,
      }));
  
      res.json(formatted);
    } catch (err) {
      console.error(err.response?.data || err.message);
      res.status(500).send('Mailler okunamadı.');
    }
  });

app.listen(PORT, () => {
  console.log(`API çalışıyor: http://localhost:${PORT}`);
});
