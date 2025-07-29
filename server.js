const express = require('express');
const axios = require('axios');
const XLSX = require('xlsx');
const cors = require('cors');

const app = express();
app.use(cors());

app.get('/prices', async (req, res) => {
  try {
    const url = 'https://onedrive.live.com/download?resid=ED7EB020544F6733%21113&authkey=%21ACNac6VLwtHQfZk&em=2';
    const response = await axios.get(url, { responseType: 'arraybuffer' });
    const workbook = XLSX.read(response.data, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
console.log('Content-Type:', response.headers['content-type']);
console.log('Response length:', response.data?.length);

    const row = data[data.length - 1];

    if (!row || row.length < 4) {
      return res.status(400).json({ error: 'Недостатньо даних у Excel' });
    }

    res.json({
      a92: row[0],
      a95: row[1],
      dp: row[2],
      gaz: row[3]
    });
  } catch (err) {
    console.error(err.message);
    res.status(500).json({ error: 'Помилка при обробці Excel' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
