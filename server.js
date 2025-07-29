const express = require('express');
const axios = require('axios');
const XLSX = require('xlsx');
const cors = require('cors');

const app = express();
app.use(cors());

app.get('/prices', async (req, res) => {
  try {
    const url = 'https://onedrive.live.com/download?resid=ED7EB020544F6733%21113&authkey=%21ACNac6VLwtHQfZk&em=2';
    const response = await axios.get(url, {
  responseType: 'arraybuffer',
  headers: {
    'User-Agent':
      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.9',
    'Referer': 'https://onedrive.live.com/',
    'Connection': 'keep-alive'
  }
});

    const workbook = XLSX.read(response.data, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Знайти останній непорожній рядок з мінімум 4 колонками
    let row = null;
    for (let i = data.length - 1; i >= 0; i--) {
      if (data[i] && data[i].length >= 4) {
        row = data[i];
        break;
      }
    }

    if (!row) {
      return res.status(400).json({ error: 'Не знайдено дійсного рядка з цінами' });
    }

    console.log('Останній рядок:', row);

    res.json({
      a92: row[0],
      a95: row[1],
      dp: row[2],
      gaz: row[3]
    });
  } catch (err) {
    console.error('Помилка:', err.message);
    res.status(500).json({ error: 'Помилка при обробці Excel' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`✅ Server running on port ${PORT}`));
