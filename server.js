const path = require('path');
var fs = require('fs');
var http = require('http');
var https = require('https');
var caKeys = fs.readFileSync(path.resolve(__dirname,'certs/ca.crt'), 'utf8');
var privateKey  = fs.readFileSync(path.resolve(__dirname,'certs/server.key'), 'utf8');
var certificate = fs.readFileSync(path.resolve(__dirname,'certs/server.crt'), 'utf8');

var credentials = {ca: caKeys, key: privateKey, cert: certificate};
var express = require('express');
var app = express();

// your express configuration here

// the __dirname is the current directory from where the script is running
app.use(express.static(__dirname));

// send the user to index html page inspite of the url
app.get('*', (req, res) => {
    console.log('__dirname' + __dirname)
    console.log('Path: ' + path.resolve(__dirname, 'dist/index.html'));  
    res.sendFile(path.resolve(__dirname, 'dist/index.html'));
});

//var httpServer = http.createServer(app);
var httpsServer = https.createServer(credentials, app);

//httpServer.listen(8080);
httpsServer.listen(process.env.PORT ||8000, '0.0.0.0');
console.log("Node server running on port "+ process.env.PORT +" over https");  