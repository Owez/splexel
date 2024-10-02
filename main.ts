import excel from "npm:exceljs";
import process from "node:process";
import fs from "node:fs";

function getFileName(): string {
  const fileName = process.argv[2];
  if (fileName == undefined) {
    console.log("Usage: splexel <file.xlsx>");
    process.exit(1);
  } else if (!fs.existsSync(fileName)) {
    console.log("File not found: " + fileName);
    process.exit(1);
  }
  return fileName;
}

async function openWorkbook(fileName: string): Promise<excel.Workbook> {
  const workbook = new excel.Workbook();
  await workbook.xlsx.readFile(fileName);
  return workbook;
}

function splitSheets(workbook: excel.Workbook): [string, excel.Workbook][] {
  const convertedWorkbooks: [string, excel.Workbook][] = [];
  workbook.eachSheet((worksheet) => {
    const convertedWorkbook = new excel.Workbook();
    convertedWorkbook.addWorksheet(worksheet.name);
    convertedWorkbooks.push([worksheet.name, convertedWorkbook]);
  });
  return convertedWorkbooks;
}

function writeWorkbooks(
  workbooks: [string, excel.Workbook][],
  prefix?: string,
) {
  const outputDir = "output";
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir);
  }

  for (const workbook of workbooks) {
    const [name, wb] = workbook;
    wb.xlsx.writeFile(`${outputDir}/${prefix ? `${prefix}-` : ""}${name}.xlsx`)
      .catch(
        (err) => {
          console.log("Error writing file: " + err);
          process.exit(1);
        },
      );
  }
}

const fileName = getFileName();
const workbook = await openWorkbook(fileName);
const workbooks = splitSheets(workbook);
writeWorkbooks(workbooks);
console.log(`Wrote ${workbooks.length} workbook(s)!`);
