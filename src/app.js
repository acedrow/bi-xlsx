const http = require("http");
var XLSX = require("xlsx");

const hostname = "127.0.0.1";
const port = 3000;
const DATA_START_ROW = 3

const row0HeaderCells = {
  A: "",
  B: "Team",
  C: "T",
  D: "Fights",
  E: "",
  F: "",
  G: "",
  H: "",
  I: "Rounds",
  J: "",
  K: "",
  L: "",
  M: "",
  N: "Score",
  O: "",
  P: "",
  Q: "",
  R: "",
  S: "",
  T: "Cards",
  U: "",
  V: "Points",
};

const row1HeaderCells = {
  A: "",
  B: "",
  C: "T",
  D: "Fw",
  E: "Fl",
  F: "F",
  G: "Cfw",
  H: "F/T",
  I: "Rw",
  J: "Rd",
  K: "Rl",
  L: "R",
  M: "Crw",
  N: "Sw",
  O: "Sw/r",
  P: "Sl",
  Q: "Sl/r",
  R: "Sdif",
  S: "Ce",
  T: "YK",
  U: "RK",
  V: "",
};

const headerMerges = [
  { s: { r: 0, c: 1 }, e: { r: 1, c: 1 } },
  { s: { r: 0, c: 2 }, e: { r: 1, c: 2 } },
  { s: { r: 0, c: 3 }, e: { r: 0, c: 7 } },
  { s: { r: 0, c: 8 }, e: { r: 0, c: 12 } },
  { s: { r: 0, c: 13 }, e: { r: 0, c: 18 } },
  { s: { r: 0, c: 19 }, e: { r: 0, c: 20 } },
  { s: { r: 0, c: 21 }, e: { r: 1, c: 21 } },
];

//expects teamnames as an array of strings
const buildTeamRows = (teamnames) => {
  teamnames.map((team, i) => ({
    A: i,
    B: team,
    C: 1,
    D: 0,
    E: 0,
    F: 0,
    G: `=D3/F3`,
    H: "F/T",
    I: "Rw",
    J: "Rd",
    K: "Rl",
    L: "R",
    M: "Crw",
    N: "Sw",
    O: "Sw/r",
    P: "Sl",
    Q: "Sl/r",
    R: "Sdif",
    S: "Ce",
    T: "YK",
    U: "RK",
    V: "",
  }));
};

const server = http.createServer((req, res) => {
  res.statusCode = 200;

  var ws = XLSX.utils.json_to_sheet([row0HeaderCells, row1HeaderCells], {
    skipHeader: true,
  });
  ws["!merges"] = [...headerMerges];
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "sheet1");
  res.setHeader("Content-Disposition", 'attachment; filename="SheetJS.xlsx"');
  res.end(XLSX.write(wb, { type: "buffer", bookType: "xlsx" }));
});

server.listen(port, hostname, () => {
  console.log(`Server running at http://${hostname}:${port}/`);
});
