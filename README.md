# Sage300-TimecardGenerator

A simple timecard generator for Sage300


## Index

* [Introduction](#introduction)
* [Download](#download)
* [Setup](#Setup)
* [Import Template](#import-template)


## Introduction

Sage300 contains an import/export spreadsheet feature when processing employee payroll. Data entry is time consuming, and that time grows exponentially with every new employee to process. Sage300-TimecardGenerator aims to solve this problem by acting as a free tool to generate importable timecards for payroll processing.


## Download

All versions can be found in the [releases](https://github.com/TBPixel/Sage300-TimecardGenerator/releases) section of the github repository.

Select the latest release zip and download. Unzip the file to a folder location you'll remember. The README.txt is there if ever you need a reference on setting up the application.


## Setup

1. Setup
	* Rename the "user-settings-EXAMPLE.ini" to "user-settings.ini"
	* Configure each field of "user-settings.ini" to match your local Sage300 database connection.

2. Run the Sage300-TimecardGenerator.exe file
	* Click file->Open Spreadsheet to select your employee hour records template spreadsheet
	* Fill out the payperiod field with your current payperiod
	* Click run

Now simply import the newly created "GENERATED-TIMECARDS.xlsx" spreadsheet file into the Sage300 Timecard system


## Import Template

The following template file shows the essential data that any template hours should follow. Make sure to read through the cell types. Dates should all set to a date format, hours should be set to a time format, not plain text.

[Download Import Template Example](https://github.com/TBPixel/Sage300-TimecardGenerator/raw/master/docs/Import%20Template%20Example.xlsx)
