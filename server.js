require('dotenv').config();
const express = require('express');
const fetch = require('node-fetch');
const cors = require('cors');
const PORT = process.env.PORT || 3000;

const app = express();
app.use(cors());
app.use(express.json());

const {
  CLIENT_ID,
  CLIENT_SECRET,
  REDIRECT_URI,
  USER_ID_ML
} = process.env;

let ACCESS_TOKEN = null;

app.get('/auth/url', (req, res) => {
  if (!CLIENT_ID || !REDIRECT_URI) {
    return res.status(500).json({ error: 'Configuração do cliente incompleta.' });
  }

  const authUrl = `https://auth.mercadolivre.com.br/authorization?response_type=code&client_id=${CLIENT_ID}&redirect_uri=${REDIRECT_URI}`;
  res.json({ authUrl });
});

// Rota para trocar "code" pelo "access_token"
app.post('/auth/token', async (req, res) => {
  const { code } = req.body;

  if (!code) {
    return res.status(400).json({ error: 'Authorization code não fornecido.' });
  }

  try {
    const params = new URLSearchParams();
    params.append('grant_type', 'authorization_code');
    params.append('client_id', CLIENT_ID);
    params.append('client_secret', CLIENT_SECRET);
    params.append('code', code);
    params.append('redirect_uri', REDIRECT_URI);

    const response = await fetch('https://api.mercadolibre.com/oauth/token', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params.toString()
    });

    if (!response.ok) {
      const errorText = await response.text();
      return res.status(response.status).json({ error: errorText });
    }

    const data = await response.json();
    ACCESS_TOKEN = data.access_token;

    res.json({ access_token: ACCESS_TOKEN });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.post('/pedido/:id', async (req, res) => {
  const orderId = req.params.id;
  const token = req.body.token || ACCESS_TOKEN;

  if (!token) return res.status(400).json({ error: 'Access token não fornecido' });

  try {
    const orderRes = await fetch(`https://api.mercadolibre.com/orders/${orderId}`, {
      headers: { Authorization: `Bearer ${token}` },
    });

    if (!orderRes.ok) {
      const err = await orderRes.text();
      return res.status(orderRes.status).json({ errorPedido: err });
    }

    const orderData = await orderRes.json();

    res.json(orderData);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.post('/liberacoes/:beginDate/:endDate/:offset', async (req, res) => {
  const { beginDate, endDate, offset } = req.params;
  const token = req.body.token || ACCESS_TOKEN;

  if (!token) return res.status(400).json({ error: 'Access token não fornecido' });

  try {
    const url = `https://api.mercadopago.com/v1/payments/search?range=money_release_date&begin_date=${beginDate}T00:00:00.000-03:00&end_date=${endDate}T23:59:59.000-03:00&status=approved&sort=money_release_date&criteria=asc&limit=500&offset=${offset}`;

    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    });

    if (!response.ok) {
      const errorText = await response.text();
      return res.status(response.status).json({ error: errorText });
    }

    const data = await response.json();
    res.json({ pedido: data });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.post('/nfe/:id', async (req, res) => {
  const orderId = req.params.id;
  const token = req.body.token || ACCESS_TOKEN;

  if (!token) return res.status(400).json({ error: 'Access token não fornecido' });

  try {
    const response = await fetch(`https://api.mercadolibre.com/users/${USER_ID_ML}/invoices/orders/${orderId}`, {
      headers: { Authorization: `Bearer ${token}` },
    });

    if (!response.ok) {
      const err = await response.text();
      return res.status(response.status).json({ error: err });
    }

    const data = await response.json();
    res.json(data);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.listen(PORT, () => {
  console.log(`Servidor rodando na porta ${PORT}`);
});