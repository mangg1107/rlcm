const express = require('express');
const path = require('path');
const { createApiApp } = require('./api');

const app = express();
const { app: apiApp, loadSheet, loadRates } = createApiApp();

app.use(express.static(path.join(__dirname, 'public')));
app.use('/api', apiApp);
app.use(apiApp);

Promise.all([loadSheet(), loadRates()])
  .catch((err) => {
    console.error('초기 시트 로드 실패:', err.message || err);
  })
  .finally(() => {
    app.listen(3000, '127.0.0.1', () => {
      console.log('http://127.0.0.1:3000 에서 실행 중');
    });
  });
