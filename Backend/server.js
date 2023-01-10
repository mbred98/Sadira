const express = require('express');
const cors = require('cors');
const app = express();
app.use(cors());
const port = process.env.PORT || 5000;
const testpy = require('./runPostcardPy')

app.get('/api/endpoint', (req, res) => {
  res.send({ data: 'Hello from the server!' });
});

app.get('/api/dailyPost', async (req, res) =>{
    const dataToSend = await testpy.runDailyPostcard();
    res.send({data: dataToSend});
});

app.get('/api/hellopy', async (req, res) => {
    const dataToSend = await testpy.testPythonHello();
    res.send({data: dataToSend});
});

app.listen(port, () => {
  console.log(`Server listening on port ${port}`);
});