# readxl

Excel file reader to dump Excel worksheets to the commandline in CSV format

## Installation

Download the appropriate binary for your platform from the latest release and place it on your $PATH.

## Usage

Use the `--help` option at the command-line for more detailed instructions, e.g. `readxl-linux --help`

This is a very simple tool, all it does is read an excel file, convert it to CSV and put the results into stdout. From there you can redirect it via a pipe or to a file.

By default the tool will convert all worksheets of the Excel file, but you can control that by selecting a worksheet by name or number. There is also an option to just print out the worksheet names.

You can also change the character used as a field separator in the CSV output, for example if the data contains commas, you can use some other character as a field separator. After that you can e.g. select the columns and rows you want using `awk`.

## Examples

These are the same examples that are given with the `--help` option at the command line, but they should give a general idea:

Display the sheet-names of foo.xlsx
```
  readxl-linux --no-contents --names -f foo.xlsx
```

Print first sheet of foo.xlsx
```
  readxl-linux --index=1 -f foo.xlsx        

```