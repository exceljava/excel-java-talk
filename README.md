# Code for LJC Talk "Tools for Working with Excel in Java"

There are three projects in this repo corresponding to the three sections of the talk.

A PDF of the slides is also included in this repo.

## apache-poi-example

Code showing how Excel workbooks can be written using Java and Apache POI, including formatting and formulas.

Requires [Apache-POI](https://poi.apache.org/).

## com4j-example

Example of how Excel can be automated from Java using the Excel COM interface and Com4j.

Requires [Jinx-COM4J](https://github.com/exceljava/jinx-com4j)

## jinx-example

Demo of extending Excel using the Jinx Excel add-in the covers the following:

- Embedding the previous com4j example in Excel as a ribbon menu
- Writing custom worksheet functions
- Passing Java objects between Excel functions

Requires [Jinx](https://exceljava.com)
