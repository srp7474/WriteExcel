WriteExcel Distribution
=======================

This zip contains the complete source, executables, javadocs and 
build commands for the WriteExcel package.

You will need Excel to verify the results.

It assumes that the envionment variable JAVA_HOME points at
a valid Java installation.

Only the JRE is needed for execution.  The SDK is required for rebuilding.

Installation
============

Unzip this into a folder of your choice.

Execution
=========

Start a CMD.exe window and CDD to the install directory

The root directory contains the .bat and .cmd
files to execute as follows. Just type the bat file name
(without the .bat). It will run the test of the specified
name. The output is in the run\out directory.

  basic
  hello
  cloner
  chart
  reader
  regress

The runner.cmd is used by the test commands.

Building
========

The build.bat rebuilds the executables. Just type build

Javadocs
========
Are available at gael-home.appspot.com or by clicking on
javadocs\index.html

The javadoc.bat rebuilds the executables. Just type javadoc




