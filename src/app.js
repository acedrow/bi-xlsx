const http = require("http");
var XLSX = require("xlsx");

const hostname = "127.0.0.1";
const port = 3000;
const DATA_START_ROW = 3;
const CHARCODE_FOR_A = 97;

const padEmptyCells = (arrayOfRowObjects) => {
  return arrayOfRowObjects.map((rowObject) => {
    const rowCopy = { ...rowObject };
    for (let i = 0; i < 26; i++) {
      //allows us to iterate through the alphabet
      const currentLetter = String.fromCharCode(
        CHARCODE_FOR_A + i
      ).toLocaleUpperCase();
      //if we don't have an existing key for a specific letter, fill it with an empty string
      console.log("rowCopy.currentLetter", rowCopy[currentLetter]);
      if (rowCopy[currentLetter] === undefined) {
        rowCopy[currentLetter] = "";
      }
    }
    return rowCopy;
  });
};

//sheet1 column width
const sheet1ColumnWidth = [
  { wch: 18 },
  { wch: 18 },
  { wch: 3 },
  { wch: 3 },
  { wch: 3 },
  { wch: 3 },
  { wch: 3 },
  { wch: 3 },
  { wch: 3 },
  { wch: 3 },
  { wch: 3 },
  { wch: 3 },
  { wch: 3 },
  { wch: 3 },
  { wch: 3 },
  { wch: 3 },
  { wch: 3 },
  { wch: 3 },
  { wch: 3 },
  { wch: 5 },
  { wch: 5 },
  { wch: 5 },
  { wch: 8 },
  { wch: 8 },
  { wch: 8 },
  { wch: 12 },

];

const headerMerges = [
  //Row 1: Buhurt Internation
  { s: { r: 0, c: 0 }, e: { r: 0, c: 25 } },
  //Row 2: Event scoring title
  { s: { r: 1, c: 0 }, e: { r: 1, c: 25 } },
  //Row 3: Event Details
  { s: { r: 2, c: 0 }, e: { r: 2, c: 18 } },
  { s: { r: 2, c: 19 }, e: { r: 2, c: 21 } },
  { s: { r: 2, c: 22 }, e: { r: 2, c: 23 } },
  //Row 4:  Event Location
  { s: { r: 3, c: 0 }, e: { r: 3, c: 18 } },
  { s: { r: 3, c: 19 }, e: { r: 3, c: 21 } },
  { s: { r: 3, c: 22 }, e: { r: 3, c: 23 } },
];

const headerCells = [
  {
    A: "Buhurt International",
  },
  {
    A: "Event Scoring Sheet",
  },
  {
    A: "Event Name:",
    T: "Event Date:",
    Y: "Event Tier:",
  },
  {
    A: "Event Location:",
    T: "Finals Type",
    Y: "Tier Mult.",
  },
];

console.log("padEmptyCells", padEmptyCells(headerCells));

// const row0HeaderCells = {
//   A: "",
//   B: "Team",
//   C: "T",
//   D: "Fights",
//   E: "",
//   F: "",
//   G: "",
//   H: "",
//   I: "Rounds",
//   J: "",
//   K: "",
//   L: "",
//   M: "",
//   N: "Score",
//   O: "",
//   P: "",
//   Q: "",
//   R: "",
//   S: "",
//   T: "Cards",
//   U: "",
//   V: "Points",
// };
// const row1HeaderCells = {
//   A: "",
//   B: "",
//   C: "T",
//   D: "Fw",
//   E: "Fl",
//   F: "F",
//   G: "Cfw",
//   H: "F/T",
//   I: "Rw",
//   J: "Rd",
//   K: "Rl",
//   L: "R",
//   M: "Crw",
//   N: "Sw",
//   O: "Sw/r",
//   P: "Sl",
//   Q: "Sl/r",
//   R: "Sdif",
//   S: "Ce",
//   T: "YK",
//   U: "RK",
//   V: "",
// };

//expects teamnames as an array of strings
const buildTeamRows = (teamnames) => {
  return teamnames.map((team, i) => ({
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

  var ws = XLSX.utils.json_to_sheet([...padEmptyCells(headerCells)], {
    skipHeader: true,
  });

  ws["!cols"] = sheet1ColumnWidth;
  ws["!merges"] = [...headerMerges];
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "sheet1");
  res.setHeader("Content-Disposition", 'attachment; filename="SheetJS.xlsx"');
  res.end(XLSX.write(wb, { type: "buffer", bookType: "xlsx" }));
});

server.listen(port, hostname, () => {
  console.log(`Server running at http://${hostname}:${port}/`);
});
