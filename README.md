# assignment-tool
`assignment-tool` is a python tool to produce feedback PDF files for grading
student assignments.

<p align="center">
<img src="img/header.png?raw=true" alt="Header Image"/>
</p>

## Installation

You can install `assignment-tool` by running

    pip install git+https://github.com/KohlbacherLab/assignment-tool.git

## Usage

The user interface of `assignment-tool` looks as follows:

	$ assignment-tool -h
	usage: assignment-tool [-h] [--pdflatex <pdflatexpath>]
			       <sheetpath> <texpath> <tutorname> sheet

	positional arguments:
	  <sheetpath>           Path to the Excel file
	  <texpath>             Path to the LaTeX template
	  <tutorname>           Name of the correcting tutor
	  sheet                 Sheet number to process

	optional arguments:
	  -h, --help            show this help message and exit
	  --pdflatex <pdflatexpath>
				pdflatex command to use

 * The Excel file must contain at least three spreadsheets following the
   specifications described below. It contains information about the course
   participants, the task sheets and the scores the participants achieved.
 * The LaTeX template for the student feedback files. An example can be found
   in the [examples](/examples) folder.
 * The name of the tutor that corrected the task sheets. With the template
   provided in the [examples](/examples) folder, it will show on the feedback
   sheet.
 * The number of the task sheet for which the feedback PDF files should be generated.
 * Optionally, using `--pdflatex` a path to the `pdflatex` binary can be
   provided. Otherwise, the `pdflatex` binary provided by the `PATH` environment will
   be invoked.

## Example

Assuming you have cloned the repository, you have installed `assignment-tool` and your current working directory is the [examples](/examples) folder, you can invoke

    assignment-tool ExampleSheet.xlsx template.tex 'Max Mustermann' 1

to generate three PDF files, one for each example participant listed in the Excel Sheet.

## Excel File Specification

The Excel file read by the two has to contain at least the following three sheets:

 * *Participants*

   Used columns: `Name`, `Username`

 * *Sheets*

   Used columns: `Sheet`, `Task`, `Subtask`, `MaxScore`

 * *Grading*

   Used columns: `Username`, `Sheet`, `Task`, `Subtask`, `Type`, `Value`

### The *Participants* Sheet

<p align="center">
  <img src="img/sheet_participants.png?raw=true" alt="Screenshot Participants Sheet"/>
</p>

The participants sheet contains information about the course participants. The
required columns are `Name` and `Username`. `Name` will be shown on the
feedback PDF and `Username` is used to join to the `Grading` sheet.

### The *Sheets* Sheet

<p align="center">
  <img src="img/sheet_sheets.png?raw=true" alt="Screenshot Sheets Sheet"/>
</p>

The *Sheets* sheet contains information about the maximum number of points that can be achieved for each task on a task sheet.

### The *Grading* Sheet

<p align="center">
  <img src="img/sheet_grading.png?raw=true" alt="Screenshot Grading Sheet"/>
</p>

The *Grading* sheet contains information about the number of points each
student achieved in each task. Additionally, comments can be added which will
be printed along the points on the feedback PDF file.

The order of the rows does not matter, tasks and subtasks will always be shown
on the feedback PDF file in numerical order. If multiple comments are added for
the same subtask, the order in which they are specified in the Excel sheet is
maintained on the feedback PDF file.

### The *Summary* Sheet

<p align="center">
  <img src="img/sheet_summary.png?raw=true" alt="Screenshot Summary Sheet"/>
</p>

The *Summary* sheet shown here is not used by `assignment-tool`. It can easily
be created using the Pivot Table feature of Excel to have an overview of the
points scored by each student in every task / sheet etc. An example is included
in the Excel sheet provided in the [examples](/examples) folder.
