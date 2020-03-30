[![Coverage Status](https://codecov.io/gh/dariober/excelToCsv/branch/master/graph/badge.svg)](https://codecov.io/gh/dariober/excelToCsv/branch/master)
[![Build Status](https://travis-ci.com/dariober/excelToCsv.svg?branch=master)](https://travis-ci.com/dariober/excelToCsv)
[![Language](http://img.shields.io/badge/language-java-brightgreen.svg)](https://www.java.com/)
[![License](http://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/dariober/excelToCsv)

<!-- vim-markdown-toc GFM -->

* [Description & Usage](#description--usage)
    * [Options](#options)
* [Installation](#installation)
    * [Similar programs](#similar-programs)
* [Developer](#developer)
    * [Cut new release](#cut-new-release)

<!-- vim-markdown-toc -->

Description & Usage
===========

Every time you use Excel to share data a kitten dies. However, if Excel is what
you have been given, you have to deal with it and one of the first things you
may want to do is to export the data to Comma-Separated Values
([CSV](https://en.wikipedia.org/wiki/Comma-separated_values)) files.

`excelToCsv` is a command-line exporter of MS Excel files to CSV format. It
supports **xlsx** and **xls** Excel files. All sheets in each input workbook are
exported and concatenated to stdout. You can use the first three columns to
extract specific spreadsheets.

Unless option `--no-prefix` is set, the first three columns of the output CSV
are always:

* Source file name

* Index of the exported spreadsheet (1-based)

* Name of the exported spreadsheet

So the actual data starts at column 4.

Options
-------

```
--delimiter DELIMITER, -d DELIMITER
                       Column delimiter (default: \t)
--na-string NA_STRING, -na NA_STRING
                       String for missing values (empty cells) (default: )
--quote QUOTE, -q QUOTE
                       Character for quoting (default: ")
--sheet [SHEET [SHEET ...]], -s [SHEET [SHEET ...]]
                       Optional list of sheet names or sheet indexes to
                       export (first sheet has index 1)
--drop-empty-rows, -r  Skip rows with only empty cells (default: false)
--drop-empty-cols, -c  Skip columns with only empty cells (default: false)
--date-as-iso, -i      Convert dates to ISO 8601 format and UTC standard. E.g
                       2020-03-28T11:40:10Z (default: false)
--no-prefix, -p        Do not prefix rows with filename, sheet index, sheet
                       name (default: false)
```

Example usage:

```
excelToCsv -i in1.xlsx in2.xlsx
excelToCsv -i in1.xlsx | awk '$3 == "Sheet1"'
```

Installation
============

`excelToCsv` requires only Java 8 or later and no installation is needed. 

* Download and unzip the latest [release](https://github.com/dariober/excelToCsv/releases/) 

* On Linux/Unix simply execute `excelToCsv [OPTS]`, on other system use `java -jar excelToCsv.jar [OPTS]`

That is:

```
curl -O https://github.com/dariober/excelToCsv/releases/download/vX.Y.Z/excelToCsv-x.y.z.zip
unzip excelToCsv-x.y.z.zip

cd excelToCsv-x.y.z/
chmod a+x excelToCsv
cp excelToCsv /usr/local/bin/     # Or else in your PATH e.g. ~/bin/
```

Similar programs
----------------

There are a number of Excel-to-CSV exporters. I found this
[excel2csv](https://github.com/informationsea/excel2csv) when I already wrote
mine also based on the Java POI package which seems pretty good. My solution
may be a bit more versatile for example when converting multiple files and
sheets.

I think converters based on Python packages like pandas, xlrd or openpyxl cannot
faithfully convert the content of MS Excel (see for example this
[question](https://stackoverflow.com/questions/60802014/how-to-consistently-handle-excel-boolean-with-pandas)
of mine).

Developer
=========

Prepare the graddle wrapper (assuming
[gradle](https://github.com/gradle/gradle) is installed)

```
gradle wrapper
```

Run tests and build project:

```
./gradlew build
```

Coverage report:

```
./gradlew jacocoTestReport
```

Inspect coverage by opening `build/reports/jacoco/test/html/index.html`

Cut new release
---------------

Prepare a zip file containing the jar file, the helper bash script and other
files the user may find useful (*e.g.* this README file). Upload this zip file
to GitHub as a new release.

```
cd ~/git_repos/excelToCsv ## Or wherever the latest local dir is

./gradlew build

VERSION='0.1.0' # To match ArgParse.VERSION

mkdir excelToCsv-${VERSION}

## Copy helper script and jar file to future zip dir
cat excelToCsv.stub build/libs/excelToCsv.jar > excelToCsv-${VERSION}/excelToCsv && chmod a+x excelToCsv-${VERSION}/excelToCsv

excelToCsv-${VERSION}/excelToCsv -h ## Check it works ok

cp build/libs/excelToCsv.jar excelToCsv-${VERSION}/
cp README.md excelToCsv-${VERSION}/

## Zip up
zip -r excelToCsv-${VERSION}.zip excelToCsv-${VERSION}
rm -r excelToCsv-${VERSION}
```
