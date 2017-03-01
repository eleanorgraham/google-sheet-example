// server.js
// where your node app starts

// init project
var express = require('express');
var app = express();

var request = require('request');
var fs = require('fs');
var tmp = require('tmp');

// http://expressjs.com/en/starter/static-files.html
app.use(express.static('public'));

// http://expressjs.com/en/starter/basic-routing.html
app.get("/", function (request, response) {
  response.sendFile(__dirname + '/views/index.html');
});

app.get("/data", function (_, response) {
    // Use localUrl to host your spreadsheet using Gomix's Assets folder:
    var localUrl = "https://s3.amazonaws.com/hyperweb-editor-assets/us-east-1%3Ac3e8cd26-9ffe-46b3-b130-48f8ac8a781c%2Fgibberish.xlsx";
    
    // Or use something fancier like Google Sheets.
    // > In Google Sheets choose File > Publish to the web... 
    // > Use the second dropdown to choose the xlsx format.
    var googlesheetUrl = "https://docs.google.com/spreadsheets/d/1vOrj0c7E3wSZODczbvPWCkiYjsqIH1jtdttE-UZI1D0/pub?output=xlsx";
    
    // We'll use this Google Sheet url by default.
    var url = googlesheetUrl;
    
    var r = request(url);
      r.on('response', (res) => {
        // Our excel parsing library expects to read from a file,
        // So we do a little bit of temp file dance to wire it up.
        // For the moment it only support .xlsx files,
        // but .csv, .xls, etc., are fairly easy extensions.

        var tmpobj = tmp.fileSync({postfix:'.xlsx'});
        var fileName = tmpobj.name;
        res.pipe(fs.createWriteStream(fileName));

        var parseXlsx = require('excel');
        parseXlsx(fileName, (err, data) => {
          if(err) throw err;

          response.send(data);
        });
      });
     r.on('error', (err) => {
         throw err;
     });
});

// listen for requests :)
var listener = app.listen(process.env.PORT, function () {
  console.log('Your app is listening on port ' + listener.address().port);
});