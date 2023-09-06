const http = require("http");
var XLSX = require("xlsx");

const hostname = "127.0.0.1";
const port = 3000;

const headerRow =
  "Team,T,Fw,Fi,F,Cfw,F/T,Rw,Rd,Rl,R,Crw,Sw,Sw/r,Sl,Sl/r,Sdif,Ce,YK,RK, ";

const testSecondRow = "TeamName,,,,=D3/F3";
const startingDataRow = 3;

var data = [
  { A: "Title", B: "b1" },
  { A: "URL", B: "b2" },
];

//expects teamnames as an array of strings
const buildSpreadsheetRows = (teamnames) => {};

const server = http.createServer((req, res) => {
  res.statusCode = 200;

  var ws = XLSX.utils.json_to_sheet(data, { skipHeader: true });
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "sheet1");
  res.setHeader('Content-Disposition', 'attachment; filename="SheetJS.xlsx"');
  res.end(XLSX.write(wb, {type:"buffer", bookType: "xlsx"}));
});

server.listen(port, hostname, () => {
  console.log(`Server running at http://${hostname}:${port}/`);
});
