const { spawn } = require('child_process');

async function testPythonHello() {
  return new Promise((resolve, reject) => {
    let data = '';
    const python = spawn('python', ['./scripts/script1.py']);

    python.stdout.on('data', function(chunk) {
      data += chunk.toString();
    });

    python.on('close', (code) => {
      if (code !== 0) {
        reject(new Error(`Python process exited with code ${code}`));
      } else {
        console.log(data)
        resolve(data);
      }
    });
  });
}


async function runDailyPostcard() {
    return new Promise((resolve, reject) => {
      let data = '';
      const python = spawn('python', ['./scripts/SalesPostCodeSummary.py']);
  
      python.stdout.on('data', function(chunk) {
        data += chunk.toString();
      });
  
      python.on('close', (code) => {
        if (code !== 0) {
          reject(new Error(`Python process exited with code ${code}`));
        } else {
          //data = data.replace(/\r\n \r\n/g, '<br>');
          //data = data.replace(/\r\n/g, '<br>');
          //data = data.replace(/\r\n\r\n/g, '<br>');
          //data = data.replace(/\n/g, '&#10;').replace(/\r/g, '&#13;'); 
          console.log(data)
          resolve(data);
        }
      });
    });
  }
module.exports = {
  testPythonHello,
  runDailyPostcard
};
