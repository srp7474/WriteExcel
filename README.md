

### WriteExcel Purpose ####

The **WriteExcel** package is a Java wrapper for the [Apache POI Excel library](https://poi.apache.org/apidocs/index.html).

A complete description of the package's capabilities can be found at [WriteExcel - POI Excel Utilities](https://stevepritchard.ca/home/WriteExcel/overview.htm)

### Brief Description ####


**WriteExcel** creates a workbook using a simple API. One or more <code>Area</code> class instances are used to specify the portion of a Sheet to populate and with what.

The interface uses formatted strings as the primary way of passing information to the support methods which interpret the strings and issue the necessary POI method calls.

In addition, an existing Workbook can be used as a template source so that sheets can be copied and then left intact, modified and/or supplemented.

The creation of Workbooks containing charts is supported by using an existing Workbook as a template that contains one or more charts and using WriteExcel to modify the data that the chart refers to.

The **ReadExcelFile** component of the package can be used to selectively iterate across existing Workbooks (or Workbooks under construction).

### Repository Contents ####

* Complete source code for package

* Build and test scripts (windows)

* POI library .jar files used for testing and building

* Javadoc files

* Auxillary files and classes used in the test harness

