// com.rexg.excel.ReadExcelRecord -

// Copyright (c) 2000-2002 AltCode Systems Inc, All Rights Reserved.
// Copyright (c) 2018 Rexcel Syste Inc, All Rights Reserved.
//
/*
 @license
 Copyright (c) 2019 by Steve Pritchard of Rexcel Systems Inc.
 This file is made available under the terms of the Creative Commons Attribution-ShareAlike 3.0 license
 http://creativecommons.org/licenses/by-sa/3.0/.
 Contact: public.pritchard@gmail.com
*/
package com.psec.excel;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Cell;


/**
  The ReadExcelRecord class in combination with the
  {@link ReadExcelFile} class
    method
  {@link ReadExcelFile#readExcel readExcel}
  or
  {@link ReadExcelFile#readExcelSmart readExcelSmart}
  is used to iterate through a Workbook sheet.
  <p>

  The class is extended with a specific class that is passed into the
  {@link ReadExcelFile#readExcel readExcel}
  or
  {@link ReadExcelFile#readExcelSmart readExcelSmart}
  method.
  <p>
  As each cell of each row that is traversed is processed the
  {@link ReadExcelRecord#canAccept canAccept} method is called.  If accepted
  the record is added, along with referenced columns translated into fields contained in the extended class definition,
  to the output set returned by
  the
  {@link ReadExcelFile#readExcel readExcel}
  or
  {@link ReadExcelFile#readExcelSmart readExcelSmart}
  method.

  @see ReadExcelFile ReadExcelFile class
  @see ReadExcelFile#readExcel readExcel method
  @see ReadExcelFile#readExcelSmart readExcelSmart method

  */

public abstract class ReadExcelRecord {
  /** The column map for record set */
  public ReadExcelFile.ColMap oCM;
  /** The Row containing the columns.*/
  public Row oRow;
  /** The ColMap string definition is returned.
  <p>
  The ColMap is an array of strings.  Each string describes a group of cells that are mapped into the fields
  of the extended ReadExcelRecord instance.
  <p>
  The format of each string is
    <p>
    n1=field1;n2=field2;....;nn=fieldn
    <p>
    where
    <ul>
    <li>nx - the absolute column number relative to 0 </li>
    <li>fieldx - the field within ReadExcelRecord to populate with the value</li>
    <li>; - separates each fild definition</li>
    </ul>
    Typically each set of fields is separated by a "/" so the example creates 3 columns (1, 7, 13) on each row where values are inspected.
    <p>
    <code>
    "1=fldA;2=fldB;4=fldC/5=fldA;6=fldB;8=fldC/13=fldA;14=fldB;16=fldC".split("/");
    </code>
  */
  public abstract String[] getColMap();
  /** This function can be changed to be selective as to what rows are included in the Row set.*/
  public boolean canAccept() throws Exception {
    return true;
  }

  // Used to copy comment to target sheet from source sheet
  /**
    Used internally in the WriteExcel clone operation
  */
  protected static class Note {
    public Cell    oCell;
    public Comment oNote;
  }

}