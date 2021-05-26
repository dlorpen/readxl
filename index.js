"use strict";

const yargs = require("yargs/yargs");
const { hideBin } = require("yargs/helpers");
const argv = yargs(hideBin(process.argv))
  .usage("Usage: $0 [options] -f <file name>")

  .alias("f", "file")
  .nargs("f", 1)
  .describe("f", "XLSX file to read")
  .demandOption(["f"])

  .describe("sheet", "only print the named sheet")
  .string(["sheet"])
  .describe("index", "only print sheet at the specified index")
  .number(["index"])

  .describe("contents", "print worksheet contents")
  .describe("names", "print Worksheet names")
  .default("names", false)
  .default("contents", true)
  .boolean(["contents", "names"])

  .describe("separator", "CSV separator for output")
  .string("separator")
  .default("separator", ",")

  .conflicts("sheet", "index")

  .check((argv, options) => {
    if (!!argv.f && Array.isArray(argv.f))
      throw new Error("only one file can be processed at a time");
    if (!(argv.index == null) && isNaN(argv.index))
      throw new Error("index must be a valid number");
    return true;
  })

  .help("h")
  .alias("h", "help")
  .example(
    "$0 --no-contents --names -f foo.xlsx",
    "display sheet names of foo.xlsx"
  )
  .example("$0 --index=1 -f foo.xlsx", "print first sheet of foo.xlsx")

  .epilog("Copyright D Orpen 2021").argv;

const os = require("os");
const EOL = os.EOL;

const exceljs = require("exceljs");
const workbook = new exceljs.Workbook();
const csvOptions = {
  delimiter: argv.separator,
};

const printSheet = (workbook, worksheet) => {
  if (argv.names) console.log(worksheet.name);
  if (argv.contents) {
    workbook.csv.write(process.stdout, {
      sheetName: worksheet.name,
      formatterOptions: csvOptions,
    });
    process.stdout.write(EOL);
  }
};

const chalk = require("chalk");

workbook.xlsx
  .readFile(argv.f)
  .then((workbook) => {
    if (argv.sheet) {
      const worksheet = workbook.getWorksheet(argv.sheet);
      printSheet(workbook, worksheet);
    } else if (argv.index) {
      const worksheet = workbook.worksheets[argv.index - 1];
      printSheet(workbook, worksheet);
    } else {
      workbook.eachSheet((worksheet, sheetId) => {
        printSheet(workbook, worksheet);
      });
    }
  })
  .catch((err) => {
    console.error(chalk.red(err.message));
    process.exit(1);
  });
