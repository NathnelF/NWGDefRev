const express = require('express');
const { exec } = require('child_process');
const cors = require('cors');
const fs = require('fs');
const https = require('https');

const app = express();
const port = 3001;

app.use(cors());

app.get('/', (req, res) => {
  res.send('Hello World!');
});

app.get('/run-script', (req, res) => {
  //const scriptPath = 'C:\\Users\\natef\\OneDrive\\Desktop\\Projects\\NWGDefRev\\backend\\script.py'; //'C:\\Users\\natef\\OneDrive\\Desktop\\Projects\\NWGDefRev\\backend\\script.py'
  exec(`python3 'script.py'`, (err, stdout, stderr) => {
    if (err) {
      console.error(`Error: ${err.message}`);
      return res.status(500).send(`Error: ${err.message}`);
    }
    if (stderr) {
      console.error(`Stderr: ${stderr}`);
      return res.status(500).send(`Stderr: ${stderr}`);
    }
    console.log(`Stdout: ${stdout}`);
    res.send(`Output: ${stdout}`);
  });
});

app.get('/run-sheets', (req, res) => {
  //const activateEnv = 'C:\\Users\\natef\\OneDrive\\Desktop\\Projects\\NWGDefRev\\venv\\Scripts\\activate &&';
  //const scriptPath = 'C:\\Users\\natef\\OneDrive\\Desktop\\Projects\\NWGDefRev\\sheets.py';

  exec(`python3 'sheets.py'`, (err, stdout, stderr) => {
    if (err) {
      console.error(`Error: ${err.message}`);
      return res.status(500).send(`Error: ${err.message}`);
    }
    if (stderr) {
      console.error(`Stderr: ${stderr}`);
      return res.status(500).send(`Stderr: ${stderr}`);
    }
    console.log(`Stdout: ${stdout}`);
    res.send(`Output: ${stdout}`);
  });
});

app.get('/run-qb', (req, res) => {
  exec(`python3 'sandboxqb.py'`, (err,stdout,stder) => {
    if (err) {
      console.error(`Error: ${err.message}`);
      return res.status(500).send(`Error: ${err.message}`);
    }
    if (stderr) {
      console.error(`Stderr: ${stderr}`);
      return res.status(500).send(`Stderr: ${stderr}`);
    }
    console.log(`Stdout: ${stdout}`);
    res.send(`Output: ${stdout}`);
  });
});

const httpsOptions = {
  key: fs.readFileSync('key.pem'),
  cert: fs.readFileSync('cert.pem')
};

https.createServer(httpsOptions, app).listen(port, () => {
  console.log(`HTTPS Server is running on https://localhost:${port}`);
});