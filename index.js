const express = require('express');
const cors = require('cors');
const axios = require('axios');
const XLSX = require('xlsx');

const app = express();
app.use(cors());

const ICLOUD_FILE_URL = 'https://www.icloud.com/iclouddrive/05c0dV2h126caWv2msr9ehFrg#Logistic_report';

// iCloud link conversion - corrected
function getDirectUrl(iCloudUrl) {
  // Extract the ID from iCloud URL
  // Format: https://www.icloud.com/iclouddrive/05c0dV2h126caWv2msr9ehFrg#Logistic_report
  const match = iCloudUrl.match(/iclouddrive\/([^#]+)/);
  if (!match) throw new Error('Invalid iCloud URL format');
  const id = match[1];
  // Use direct download endpoint
  return `https://p01-icloud.icloud.com/download/${id}/Logistic_report.xlsm`;
}

app.get('/api/orders', async (req, res) => {
  try {
    const url = getDirectUrl(ICLOUD_FILE_URL);
    
    const response = await axios.get(url, {
      responseType: 'arraybuffer',
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
      }
    });

    const workbook = XLSX.read(response.data, { type: 'array' });
    const sheet = workbook.Sheets['EGYPT Customers'];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const colIdx = {
      rif: 0, year: 1, supplier: 4, consignee: 51, goods: 53,
      sc_value: 57, amount: 67, ship_status: 41, coll_status: 63,
      paying_status: 27, balance: 62
    };

    const orders = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const order = {};
      for (const [key, idx] of Object.entries(colIdx)) {
        const val = row[idx];
        if (val != null && val !== '') {
          order[key] = String(val).trim();
        }
      }
      if (order.rif) orders.push(order);
    }

    res.json({ success: true, count: orders.length, orders });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/health', (req, res) => {
  res.json({ status: 'ok' });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`✅ Server running on port ${PORT}`);
});
