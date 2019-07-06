// com.rexg.excel.ReadExcelFile -

// Copyright (c) 2000-2002 AltCode Systems Inc, All Rights Reserved.
// Copyright (c) 2018 Rexcel Syste Inc, All Rights Reserved.

/*  History:

      EC9504 - changed relativity of ColMap from 1 to 0 (for consistency);
*/

/*
 @license
 Copyright (c) 2019 by Steve Pritchard of Rexcel Systems Inc.
 This file is made available under the terms of the Creative Commons Attribution-ShareAlike 3.0 license
 http://creativecommons.org/licenses/by-sa/3.0/.
 Contact: public.pritchard@gmail.com
*/
package com.psec.excel;
import java.util.ArrayList;
import java.io.FileInputStream;
import java.lang.reflect.Field;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFRow;
//import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;


//import org.apache.poi.ss.usermodel.CellValue;

//import com.psec.util.CU;

/**
  The ReadExcelFile class opens an existing Workbook with one or more Sheets.
  <p>
  Methods are provided that can return selected rows of data from each sheet.  The {@link ReadExcelRecord} class
  is used to provide selection capabilities as to what rows are to be returned and what content of each row is to be returned.
  <p>
  This selection capability is used to create a regression test facility for the WriteExcel set of methods.
  <p>
  The ReadClassFile class is also used to point to a sheet that can be cloned using the {@Link WriteExcel#clone} method.

  <h5>Example</h5>
  This reads the file <i>demo-excel-inp.xlsx</i> in the demo package.
  <p>
  <pre style="font-size:90%;">
  public static class OpenBalRec extends ReadExcelRecord {
    static Matcher oM = Pattern.compile("^[ALS][0-9]{3}$").matcher("");
    public String[] getColMap() {
      return "3=sLab;2=sName;4=dVal".split("/");
    }
    public String sLab;
    public String sName;
    public double dVal;

    &#64;Override
    public boolean canAccept() throws Exception {
      if (this.sLab == null) return false;
      if (oM.reset(this.sLab).find()) return true;
      return false;
    }
  }

  private void demoReader() throws Exception {
    oREF = new ReadExcelFile();
    oREF.openFile(oOpt.sInp);
    ReadExcelRecord[] oRows = oREF.readExcelSmart("sample-bal-template",OpenBalRec.class,1,true);
    for(ReadExcelRecord oRER:oRows) {
      OpenBalRec oOBR = (OpenBalRec)oRER;
      log(String.format("%s %10.2f %s",oOBR.sLab,oOBR.dVal,oOBR.sName));
    }
    oREF.closeFile();
  }
  </pre>

  @see <a href={@docRoot}overview-summary.html#WriteExcel-desc>WriteExcel description</a>
  @see <a href={@docRoot}com/psec/excel/ReadExcelRecord.html>ReadExcelRecord</a>

*/
public class ReadExcelFile {
  Workbook oWB;
  FormulaEvaluator oFE;
  String sWorkbookFile;

  private static void log(String sMsg) {System.out.println(sMsg);}
  private static Exception e(String s) {return new Exception(s); }

  private static class RowColSet {
    ArrayList<ColMap[]> oCMs = new ArrayList<>();
  }


  /**
    The ColMap class describes a single column used internally as part of the
    {@link ReadExcelFile#readExcel readExcel} method processing.

    @see ReadExcelRecord
  */
  public static class ColMap {
    /** The field that can be used to populate the value using Java reflection in the extended ReadExcelRecord class*/
    public Field oFld;
    /** The name of the field in the extended ReadExcelRecord class*/
    public String sName;
    /** The column number relative to 0 on the Workbook sheet where the value is to be obtained*/
    public int nCol;
    /** The constructor
      @param nCol The column number relative to 0
      @param sName The name of the field (column)
      @param oFld The field that can be used to extract the value using Java reflection
    */
    public ColMap(int nCol,String sName,Field oFld) {
      this.nCol  = nCol;
      this.sName = sName;
      this.oFld  = oFld;
    }
  }



  /**
    Opens the Workbook.
    @param sFile the name of the file.

    @throws Exception file does not exist or is not a valid Workbook
  */
  public void openFile(String sFile) throws Exception {
    log("Opening "+sFile);
    oWB = new XSSFWorkbook(new FileInputStream(sFile));
    if (oWB == null) throw e("oWB is null");
    sWorkbookFile = sFile;
    oFE = oWB.getCreationHelper().createFormulaEvaluator();
  }

  /**
    Attaches the Workbook to a ReadExcelFile instance so all its methods can be used. This method is used
    instead of the {@link ReadExcelFile#openFile openFile} method when a Workbook is known (such as in an ongoing {@link WriteExcel} operation.
  */
  public void attachWorkbook(Workbook oWB) {
    this.oWB = oWB;
    sWorkbookFile = null;
    oFE = oWB.getCreationHelper().createFormulaEvaluator();
  }


  /**
    Closes the Workbook.  It is present for comlpeteness but
    does nothing apart from releasing memory.
    <p>
    It is ignored if the Workbook is part of
    an active {@link WriteExcel} operation attached to this instance via {@link ReadExcelFile#attachWorkbook attachWorkbook} method.

    @throws Exception An internal operation throws an Exception.
  */
  public void closeFile() throws Exception {
    if ((oWB != null) && (sWorkbookFile != null)) {
      oWB.close();
    }
  }



  /**
    Returns the named Sheet.  It allows raw access to to Workbook objects.

    @Param sSheet The name of the sheet to return.
    @return org.apache.poi.ss.usermodel.Sheet or null if the named sheet does not exist.
  */
  public Sheet getSheet(String sSheet) /*throws Exception*/ {
    return oWB.getSheet(sSheet);
  }

  /**
    Returns the String version of a Sheet print area.  It is obtained by calling Workbook.getPrintArea for the specified sheet.

    @Param sSheet The name of the sheet to use.
    @return The internal string representation of the print area.
    @throws Exception The sheet is not found
  */

  public String getPrintArea(String sSheet) throws Exception {
    Sheet oS = oWB.getSheet(sSheet);
    if (oS == null) return null;
    return oWB.getPrintArea(oWB.getSheetIndex(oS));
  }

  /**
    Iterate across a sheet and selectively create a set of
    {@link ReadExcelRecord readExcelRecord} instances, one for each row selected.
    <p>
    It is implemented by calling {@link ReadExcelFile#readExcelSmart readExcelSmart} with a value of 0 for nStartRow and the value of
    bSkipBlank set to false.

    @param sSheet the sheet name to process
    @param oRecCls The class type used to instantiate each row {@link ReadExcelRecord readExcelRecord} instance.

    @return The array of ReadExcelRecord[] objects.

    @throws Exception The sheet is not found
    */

  public ReadExcelRecord[] readExcel(String sSheet,Class<ReadExcelRecord> oRecCls) throws Exception {
    RowColSet oRCS = makeColMap(oRecCls);
    if (oWB == null) throw e("oWB is null");
    Sheet oSheet    = oWB.getSheet(sSheet);
    //XSSFSheet oSheet    = oWB.getSheetAt(3);
    if (oSheet == null) throw e("Sheet is null "+sSheet);
    ArrayList<ReadExcelRecord> oLst = new ArrayList<>();
    int nMaxRow = 0;
    for(ColMap[] oCMs:oRCS.oCMs) {
      int nRows = 1;
      for(int i=2,iMax=999; i<iMax; i++) {
        Row oRow        = oSheet.getRow(i);
        if (oRow == null) break;
        Object oVal = readCell(oRow,(oCMs[0].nCol-1),false);
        if (oVal == null) break;
        processRow(oLst,nRows++,oRow,oCMs,oRecCls);
        if (nMaxRow < i) nMaxRow = i;
      }
      log("Read rows:"+nRows);
    }
    ReadExcelRecord[] oRs = (ReadExcelRecord[])oLst.toArray(new ReadExcelRecord[oLst.size()]);
    log("Read Max rows:"+nMaxRow+" produced "+oRs.length+" records");
    return oRs;
  }

  /**
    Iterate across a sheet and selectively create a set of
    {@link ReadExcelRecord readExcelRecord} instances, one for each row selected.
    <p>
    The <code>nStartRow</code> parameter is used to indicate the first row to start processing at.  Its primary use is to
    skip headers frequently found at the top of sheets.
    <p>
    The <code>bSkipBlank</code> will cause rows to be skipped in which the first field of the {@link ReadExcelFile.ColMap ColMap} object has no Cell object.

    @param sSheet the sheet name to process
    @param oRecCls The class type used to instantiate each row {@link ReadExcelRecord readExcelRecord} instance.
    @param nStartRow The first row to process relative to 0;  This allows the
    @param bSkipBlank Rows where the first field is null are skipped.

    @return The array of ReadExcelRecord[] objects.

    @throws Exception The sheet is not found
    */

  public ReadExcelRecord[] readExcelSmart(String sSheet,Class<?> oRecCls,int nStartRow,boolean bSkipBlank) throws Exception {
    RowColSet oRCS = makeColMap(oRecCls);
    if (oWB == null) throw e("oWB is null");
    Sheet oSheet    = oWB.getSheet(sSheet);
    if (oSheet == null) throw e("Sheet is null "+sSheet);
    ArrayList<ReadExcelRecord> oLst = new ArrayList<>();
    int nActRows = 0;
    int nMaxRow = 0;
    for(ColMap[] oCMs:oRCS.oCMs) {
      int nRows = 0;
      for(int i=nStartRow,iMax=oSheet.getLastRowNum()+1; i<iMax; i++) {
        nRows += 1;
        Row oRow        = oSheet.getRow(i);
        if (oRow == null) continue;
        Object oVal = readCell(oRow,(oCMs[0].nCol),false);
        if (oVal == null) continue;
        nActRows += 1;
        processRow(oLst,nRows++,oRow,oCMs,oRecCls);
        if (nMaxRow < i) nMaxRow = i;
      }
      log("Read rows set:"+nRows+" sum.actrows="+nActRows);
    }
    ReadExcelRecord[] oRs = (ReadExcelRecord[])oLst.toArray(new ReadExcelRecord[oLst.size()]);
    log(String.format("Read Max rows: %d   Active Rows: %d   Produced %d rows rowsets=%d",nMaxRow,nActRows,oRs.length,oRCS.oCMs.size()));
    return oRs;
  }

  RowColSet makeColMap(Class<?> oRecCls) throws Exception {
    RowColSet oRCS = new RowColSet();
    ReadExcelRecord oRER = (ReadExcelRecord)oRecCls.newInstance();
    String[] sColMaps = oRER.getColMap();
    for(String sColMap:sColMaps) {
      String[] sCols = makeTokenArray(sColMap,';');
      Pattern oPDef = Pattern.compile("^([a-zA-Z][a-zA-Z0-9]+)$");
      Pattern oPFld = Pattern.compile("(\\d+)=(.*)");
      ColMap[] oCMs = new ColMap[sCols.length];
      int nColNum = 0;
      for(int i=0,iMax=sCols.length; i<iMax; i++) {
        nColNum += 1;
        int nCol = 0;
        String sName = null;
        Matcher oM = oPDef.matcher(sCols[i]);
        if (oM.matches()) {
          nCol = nColNum - 1;
          sName = oM.group(1);
        } else {
          oM = oPFld.matcher(sCols[i]);
          if (!oM.matches()) throw e("Blew up at "+sCols[i]);
          String sGrp = oM.group(1);
          sName = oM.group(2);
          //log("Process "+sGrp+" "+sName);
          nCol = Integer.parseInt(sGrp);
        }
        Field oFld = oRecCls.getField(sName);
        if (oFld == null) throw e("Lost "+sName);
        oCMs[i] = new ColMap(nCol,sName,oFld);
      }
      oRCS.oCMs.add(oCMs);
    }
    return oRCS;
  }

  private String[] makeTokenArray(String sToks,char cDelim) {
    int nAt = 0;
    if (sToks.length() == 0) return new String[0];
    int nRows = 1;
    while(nAt < sToks.length()) {
      int n = sToks.indexOf(cDelim,nAt);
      if (n < 0) break;
      nAt = n+1;
      nRows++;
    }
    String[] sR = new String[nRows];
    nAt = 0;
    nRows = 0;
    while(nAt < sToks.length()) {
      int n = sToks.indexOf(cDelim,nAt);
      if (n < 0) break;
      sR[nRows++] = sToks.substring(nAt,n);
      nAt = n+1;
    }
    sR[nRows] = "";
    if (nAt < sToks.length()) sR[nRows] = sToks.substring(nAt);
    return sR;
  }


  void processRow(ArrayList<ReadExcelRecord> oLst,int nRow,Row oRow,ColMap[] oCMs,Class<?> oRecCls) throws Exception {
    ReadExcelRecord oRER = (ReadExcelRecord)oRecCls.newInstance();
    for(int i=0,iMax=oCMs.length; i<iMax; i++) {
      ColMap oCM = oCMs[i];
      if (oCM.oFld.getType().getName().endsWith("ReadExcelRecord$Note")) {
        ReadExcelRecord.Note oNote = new ReadExcelRecord.Note();
        oCM.oFld.set(oRER,oNote);
        oNote.oCell = getColCell(oRow,oCM);
        if (oNote.oCell != null) oNote.oNote = oNote.oCell.getCellComment();
        log("ColMap "+oCM.sName+" "+oCM.oFld.getType().getName()+" "+oNote.oNote);
      } else {
        Object oVal  = readCell(oRow,oCM.nCol,false);
        //if (oCM.sName.equals("mTranID") && (oVal != null)) log("Got "+oCM.sName+" val <"+oVal+"> "+oVal.getClass().getName()+" "+oCM.oFld.getType().getName());
        try {
          if (oVal.getClass().getName().endsWith(".Double") && oCM.oFld.getType().getName().equals("long")) {
            oCM.oFld.setLong(oRER,((Double)oVal).longValue());
          } else if (oVal.getClass().getName().endsWith(".Double") && oCM.oFld.getType().getName().equals("int")) {
            oCM.oFld.setInt(oRER,((Double)oVal).intValue());
          } else {
            oCM.oFld.set(oRER,oVal);
          }
        } catch (Exception e) {}
      }
      oRER.oCM  = oCM;
      oRER.oRow = oRow;
    }
    if (oRER.canAccept()) {
      oLst.add(oRER);
    }
  }

  Object readCell(Row oRow,int nCol,boolean bReqd) throws Exception {
    return readCell(oRow,nCol,bReqd,false);
  }

  String readCellStr(Row oRow,int nCol,boolean bReqd) throws Exception {
    //L("Reading cell row "+oRow.getRowNum()+" col "+nCol);
    Cell oCell     = oRow.getCell((short)nCol);
    if (oCell == null) return null;
   return getCellAsStr(oCell);
  }

  Object readCell(Row oRow,int nCol,boolean bReqd,boolean bDebug) throws Exception {
    //L("Reading cell row "+oRow.getRowNum()+" col "+nCol);
    Cell oCell     = oRow.getCell((short)nCol);
    //log("readcell "+oRow.getSheet().getSheetName()+" r="+oRow.getRowNum()+" c="+nCol);
    if (oCell == null) return null;
    CellType oCT = oCell.getCellType();
    if (oCT == CellType.STRING) {
      String sVal = oCell.getStringCellValue();
      if (sVal == null) return null;
      sVal = sVal.trim();
      sVal = dequote(sVal);
      if (sVal == null) return null;
      if (sVal.length() == 0) return null;
      return sVal;
    } else if (oCT == CellType.BLANK) {
      return null;
    } else if (oCT == CellType.NUMERIC) {
      return oCell.getNumericCellValue();
    } else if (oCT == CellType.FORMULA) {
      DataFormatter formatter = new DataFormatter();
      String s = null;
      try {
        s = formatter.formatCellValue(oCell,oFE);
      } catch(Exception e) {
        s = formatter.formatCellValue(oCell);
        log("Exception "+e+" "+s);
      }
      if (s.trim().length() == 0) s = "0.00";
      return Double.parseDouble(s);
    } else {
      throw e("Cannot handle "+oCT);
    }
    //int nType = 0; //4.0 oCell.getCellType();
    //if (bDebug) log("Cell Type "+nType+" row "+/*oRow.getRowNum()*/"?"+" col "+nCol);
    //switch (nType) {
      /*4.0 case XSSFCell.CELL_TYPE_BLANK: return null;
      case XSSFCell.CELL_TYPE_STRING:
        String sVal = oCell.getStringCellValue();
        if (sVal == null) return null;
        sVal = sVal.trim();
        sVal = dequote(sVal);
        if (sVal == null) return null;
        if (sVal.length() == 0) return null;
        return sVal;
      case XSSFCell.CELL_TYPE_NUMERIC:
        double oVal = oCell.getNumericCellValue();
        return ""+oVal;
      case XSSFCell.CELL_TYPE_FORMULA:
        //L("CellType.2 "+oCell+" "+oCell.getNumericCellValue());
        return ""+oCell.getNumericCellValue();*/
      //default: throw e("Bad Cell type:"+nType+" Col="+nCol);
    //}
    //return getCellAsStr(oCell);
  }

  /**
  Return a stringized version of the Cell at the specified location. The Cell represented by the
  {@link ReadExcelFile.ColMap} parameter within the oRow parameter is the one processed.

  @param oRow The row to process which can be obtained via the {@link ReadExcelFile#getSheet getSheet} method.
  @param oCM   The column to select.

  @return a stringized version of the selected Cell.
  */

  public String getColAsStr(Row oRow,ColMap oCM) throws Exception {
    Cell oCell     = oRow.getCell((short)oCM.nCol);
    if (oCell == null) return null;
    return getCellAsStr(oCell);
  }

  /**
  Return Cell at the specified location.

  The Cell represented by the
  {@link ReadExcelFile.ColMap} parameter within the oRow parameter is the returned.

  @param oRow The row to process which can be obtained via the {@link ReadExcelFile#getSheet getSheet} method.
  @param oCM   The column to select.

  @return The selected Cell or null.
  */
  public Cell getColCell(Row oRow,ColMap oCM) throws Exception {
    Cell oCell     = oRow.getCell((short)oCM.nCol);
    return oCell;
  }


  private String getCellAsStr(Cell oC) throws Exception {
    if (oC == null) return null;
    CellType oType = oC.getCellType(); // javadoc is wrong here
    //if (bDebug) log("Cell Type "+nType+" row "+/*oRow.getRowNum()* /"?"+" col "+nCol);
    switch (oType) {
      case BLANK: return null;
      case STRING:
        String sVal = oC.getStringCellValue();
        if (sVal == null) return null;
        sVal = sVal.trim();
        sVal = dequote(sVal);
        if (sVal == null) return null;
        if (sVal.length() == 0) return null;
        return sVal;
      case NUMERIC:
        double oVal = oC.getNumericCellValue();
        return ""+oVal;
      case FORMULA:
        DataFormatter formatter = new DataFormatter();
        String s = null;
        try {
          s = formatter.formatCellValue(oC,oFE);
        } catch(Exception e) {
          s = formatter.formatCellValue(oC);
          log("Exception "+e+" "+s);
        }
        log("Formulat "+s);
        return s;
      default: throw e("Bad Cell type:"+oType+" Col="+oC.getColumnIndex());
    }
  }


  String retNull(boolean bReqd) {
    return (bReqd?"\"?\"":"NULL");
  }

  String dequote(String sD) {
    sD = sD.trim();
    while(sD.startsWith("\"")) sD = sD.substring(1);
    while(sD.endsWith("\"")) sD = sD.substring(0,sD.length()-1);
    sD = sD.trim();
    if (sD.length() == 0) return null;
    return sD;
  }

  // not reuired atpresent
  /*public*/ String lookup(String sSheet,int nRow,String sRowKey,int nCol,String sColKey) throws Exception {
    Sheet oSheet    = oWB.getSheet(sSheet);
    if (oSheet == null) throw e("Where is "+sSheet);
    Row oRow        = oSheet.getRow(nRow);
    int nHitCol = -1;
    int nHitRow = -1;
    sRowKey = sRowKey.toUpperCase();
    sColKey = sColKey.toUpperCase();
    // Find Col
    for(int i=0,iMax=99; i<iMax; i++) {
      Object oVal = readCell(oRow,i,false);
      if ((oVal != null) && sRowKey.equals(((String)oVal).toUpperCase())) {
        nHitCol = i;
        break;
      }
    }
    if (nHitCol == -1) throw e("Lost "+sRowKey);
    //L("Col for "+sRowKey+" is "+nHitCol);
    // Find Row
    boolean bHadNull = false;
    for(int i=nRow+1,iMax=9999; i<iMax; i++) {
      Row oRow0  = oSheet.getRow(i);
      if (oRow0 == null) {
        if (bHadNull) throw e("Lost col "+sColKey+" "+i);
        bHadNull = true;
        continue;
      }
      bHadNull = false;
      Object oVal = readCell(oRow0,nCol,false);
      if ((oVal != null) && sColKey.equals(((String)oVal).toUpperCase())) {
        nHitRow = i;
        break;
      }
    }
    if (nHitRow == -1) throw e("Lost "+sColKey);
    // Find Value at crossroads
    oRow        = oSheet.getRow(nHitRow);
    if (oRow == null) throw e("Lost row "+nHitRow);
    Object oVal = readCell(oRow,nHitCol,false,false);
    //if (sVal == null) L("Expected value at Sheet "+sSheet+" row "+nHitRow+" col "+nHitCol);
    return ""+oVal;
  }

  // not reuired atpresent
  /*public*/ String[][] readSheetAsStrMatrix(String sSheet,int nFirstRow,int nLastRow,int nLastCol) throws Exception {
    Sheet oSheet    = oWB.getSheet(sSheet);
    if (oSheet == null) throw e("Sheet "+sSheet+" not defined");
    ArrayList<String[]> oLst = new ArrayList<>();
    for(int i=nFirstRow,iMax=nLastRow; i<iMax; i++) {
      Row oRow = oSheet.getRow(i);
      if (oRow == null) continue;
      String[] oRowStr = new String[nLastCol+1];
      for(int j=0,jMax=nLastCol+1; j<jMax; j++) {
        String sStr = readCellStr(oRow,j,false);
        oRowStr[j] = sStr;
      }
      oLst.add(oRowStr);
    }
    String[][] oXs = (String[][])oLst.toArray(new String[oLst.size()][]);
    return oXs;
  }

  /**
    Gets the <code>Workbook</code> being accessed.
    <p>
    This allows inspection or modifications to be made to the Workbook using the POI library directly.
    @return Workbook.
  */
  public Workbook getWorkbook() throws Exception {
    return oWB;
  }


}
