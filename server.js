var fs = require('fs');
var http = require('http');
var https = require('https');
var caKeys = [fs.readFileSync('./certs/ca.crt'), ];
var privateKey  = fs.readFileSync('./certs/server.key', 'utf8');
var certificate = fs.readFileSync('./certs/server.crt', 'utf8');

var credentials = {ca: caKeys, key: privateKey, cert: certificate};
var express = require('express');
var app = express();

// your express configuration here

// the __dirname is the current directory from where the script is running
app.use(express.static(__dirname));

// send the user to index html page inspite of the url
app.get('*', (req, res) => {
  res.sendFile(path.resolve(__dirname, 'index.html'));
});

//var httpServer = http.createServer(app);
var httpsServer = https.createServer(credentials, app);

//httpServer.listen(8080);
httpsServer.listen(8000);
console.log('Node server running on port 8000 over https');  