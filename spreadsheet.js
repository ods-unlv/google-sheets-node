const GoogleSpreadsheet = require("google-spreadsheet");
const { promisify } = require("util");

const creds = require("./client_secret.json");

function printStudent(student) {
  console.log(`Name: ${student.studentname}`);
  console.log(`Major: ${student.major}`);
  console.log(`Home State: ${student.homestate}`);
  console.log("----------------------------");
}

async function accessSpreadsheet() {
  const doc = new GoogleSpreadsheet(
    "15STY1x3NZwU4wdhwj0Iv_ajh3maD7yRW8GVExSZLqXo"
  );
  await promisify(doc.useServiceAccountAuth)(creds);
  const info = await promisify(doc.getInfo)();
  const sheet = info.worksheets[0];
  //console.log(`Title: ${sheet.title}, Rows: ${sheet.rowCount}`);

  //   const rows = await promisify(sheet.getRows)({
  //     // offset: 5, //start row
  //     // limit: 10,
  //     // orderby: "homestate"
  //     query: "homestate = NY"
  //   });

  //console.log(rows);

  //   rows.forEach(row => {
  //     //printStudent(row);
  //     row.homestate = "NV";
  //     row.save();
  //   });

  //   const row = {
  //     studentname: "Dmitrii",
  //     major: "Computer Science",
  //     homestate: "Stavropol krai",
  //     classlevel: "5. Graduated",
  //     extracurricularactivity: "Economics"
  //   };

  //   await promisify(sheet.addRow)(row);

  //   const rows = await promisify(sheet.getRows)({
  //     query: "studentname = Dmitrii"
  //   });

  //   rows[0].del();

  const cells = await promisify(sheet.getCells)({
    "min-row": 1,
    "max-row": 3,
    "min-col": 1,
    "max-col": 2
  });

  for (const cell of cells) {
    console.log(`${cell.row},${cell.col}: ${cell.value}`);
  }

  cells[2].value = "Peneloppa";
  cells[2].save();
}

accessSpreadsheet();
