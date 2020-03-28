[![Coverage Status](https://codecov.io/gh/dariober/excel2csv/branch/master/graph/badge.svg)](https://codecov.io/gh/dariober/excel2csv/branch/master)
[![Build Status](https://travis-ci.com/dariober/excel2csv.svg?branch=master)](https://travis-ci.com/dariober/excel2csv)
[![Language](http://img.shields.io/badge/language-java-brightgreen.svg)](https://www.java.com/)
[![License](http://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/dariober/excel2csv)

<!-- vim-markdown-toc GFM -->

* [Description & Usage](#description--usage)
    * [Options](#options)
* [Installation](#installation)
* [Developer](#developer)
    * [Cut new release](#cut-new-release)

<!-- vim-markdown-toc -->

Description & Usage
===========

Every time you use Excel to share data a kitten dies. However, if Excel is what
you have been given, you have to deal with it and one of the first things you
may want to do is to export the data to Comma-Separated Values
([CSV](https://en.wikipedia.org/wiki/Comma-separated_values)) files.

`excel2csv` is a command-line exporter of MS Excel files to CSV format. It
supports **xlsx** and **xls** Excel files. All sheets in each input workbook are
exported and concatenated to stdout. You can use the first three columns to
extract specific spreadsheets.

The first three columns of the output CSV are always:

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
                     Character for quoting or an empty string for no quoting
                     (default: ")
--drop-empty-rows, -r  Skip rows with only empty cells (default: false)
--drop-empty-cols, -c  Skip columns with only empty cells (default: false)
--date-as-iso, -i      Convert dates to ISO 8601 format and UTC standard. E.g
                       2020-03-28T11:40:10Z (default: false)
```

Example usage:

```
excel2csv in1.xlsx in2.xlsx
excel2csv in1.xlsx | awk '$3 == "Sheet1"'
```

Installation
============

`excel2csv` requires only Java 8 or later.

```
curl -O https://github.com/dariober/excel2csv/releases/download/vX.Y.Z/excel2csv-x.y.z.zip
unzip excel2csv-x.y.z.zip

cd excel2csv-x.y.z/
chmod a+x excel2csv
cp excel2csv.jar /usr/local/bin/ # Or else in your PATH e.g. ~/bin/
cp excel2csv /usr/local/bin/     # Or else in your PATH e.g. ~/bin/
```

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
cd ~/git_repos/excel2csv ## Or wherever the latest local dir is

./gradlew build

VERSION='0.1.0' # To match ArgParse.VERSION

mkdir excel2csv-${VERSION}

## Copy helper script and jar file to future zip dir
cp excel2csv excel2csv-${VERSION}/
cp build/libs/excel2csv.jar excel2csv-${VERSION}/
cp README.md excel2csv-${VERSION}/

## Zip up
zip -r excel2csv-${VERSION}.zip excel2csv-${VERSION}
rm -r excel2csv-${VERSION}
```
