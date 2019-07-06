// WriteExcel - Write Excel WorkSheet (.xlsx)

// Copyright (c) 2017 Rexcel Systems Inc, All Rights Reserved.

/*

 Change History:
   EC9418 - Original from WriteExcel
          - added createArea notion, eliminated SheetData
          - Made all Styles into HashMap
          - e
*/

/*
 @license
 Copyright (c) 2019 by Steve Pritchard of Rexcel Systems Inc.
 This file is made available under the terms of the Creative Commons Attribution-ShareAlike 3.0 license
 http://creativecommons.org/licenses/by-sa/3.0/.
 Contact: public.pritchard@gmail.com
*/

package com.psec.excel;
import java.io.FileOutputStream;
import java.io.FileInputStream;
import java.util.ArrayList;
//import java.util.List;
import java.util.HashMap;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
//import  com.rexg.util.CU;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFDrawing;
//import org.apache.poi.xssf.usermodel.XSSFChart;
//import org.apache.poi.xddf.usermodel.chart.XDDFChart;
//import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
//import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.CellType;

import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PrintSetup;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.IndexedColorMap;
import org.apache.poi.common.usermodel.HyperlinkType;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.BorderStyle;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellAddress;

/**
  The WriteExcel class creates a Workbook with one or more Sheets.  The {@link Area}
  object is used to provide the columns of data and rows of columns with which to
  populate the data.  <code>CellStyles</code> are provided by default, by column specification
  or by row.
<p>
  Multiple Areas can be applied to one sheet and multiple Sheets can be written using
  multiple Areas.
<p>
  <code>CellStyles</code> are created using the {@link WriteExcel#addStyle addStyle} or {@link WriteExcel#addStyles addStyles} methods.
  <code>WriteExcel</code> manages the combinations so only the active combinations are written to the the output Workbook CellStyle registry.

 @see <a href={@docRoot}overview-summary.html#WriteExcel-desc>WriteExcel description</a>

*/

public class WriteExcel {

  /** Used to specify the contents of a portion of a Sheet.
  */
  private static class HdrCol {
    int     nHdrIX;
    int     nMaxStr;
    int     nWidthMult;  //Factor to calc column width
    int     nMerge;
    String  sText;
    //String  sRawText;
    String  sHdrFmt;
    boolean bPlain;
  }

  /** Used to manage DataFmts which is the DataFormat string
    we store with a CellStyle.
  */
  private static class DataFmt {
    String  sName;
    String  sDataFmt;
    Matcher oM;
    boolean bInteger;
  }

  /** Used to extract specific StyleStrs prefixed to data
    as {mm.xxx}
  */
  private static class SpecFmt {
    String  sName;
    int     nMerge;
    String  sAnonFmt;
    String  sData;  // remaining data
    boolean bPlain;
    DataFmt oDF;
//
  }


  /** Used to manage Styles so that only those in play are added
    to the style sheet. Registered styles have oStyle and oDF populated.
  */

  private static class StyleDef {
    String     sName;     // internal name.  Names starting with "#" are builtins
    String     sStyStr;   // String to make style with.  see makeStyleParse and StyleAttrs
    StyleAttrs oSA;       // resulting attribute string
  }

  /** Used to manage Styles so that only those in play are added
    to the style sheet. Registered styles have oStyle and oDF populated.
  */

  /** Used to store parsed Styles strings
  */
  private static class StyleAttrs {
    boolean bBackOnly;          // Only BG color specified
    boolean bNeedFont;          // any cmd requiring font
    //                          ---- related command(s) -----
    String  sFontFamily;        // f - fixed or FF(name)
    String  sCustExit;          // CE(str)
    boolean bBold;              // b
    boolean bItalic;            // i
    boolean bStrike;            // s
    double  dPoints;            // n.m or n
    boolean bLink;              // l (link)
    byte[]  oForeRGB;           // FG(r,g,b) or FG(red}blue|green| etc See poi IndexedColors. Added gray)
    byte[]  oBackRGB;           // BG(r,g,b) or BG(red}blue|green| etc               ,, )
    short   nForeIX;            // Indexes when standard color
    //short   nBackIX;

    char    cHorzAlign = 0;     // C or L or R (L is default for strings, R for numerics and integers);
    char    cVertAlign = 0;     // T or B or M (M is default for all);
    char    cTot       = 0;     // = or - or ~ as totals indicators;

  }

  /** Used to specify the contents of a portion of a Sheet.
  */
  public static class Area {
    WriteExcel            oWE;
    int                   nBaseRow;
    int                   nBaseCol;
    ArrayList<HdrCol[]>   oHdrs    = new ArrayList<>();
    ArrayList<String[]>   oRows    = new ArrayList<>();
    //ArrayList<Integer>    nStripes;
    ArrayList<String>     sColFmts;
    String                sCurSheet;
    // computed values
    private int           nMaxCol;

    /**
      Class constructor.
    */
    public Area(){}

    /**
      Creates a Data filter line as row 0.
      <p>
      If the top rows of a sheet has merged values, the Excel Data filter is impeded.  Calling this method ensures
      row 0 of a sheet has an empty cell for every column in the sheet.  This makes the Data filter function work properly.
      @returns Area for chaining purposes.
    */
    public Area addDataFilterLine() {
      Sheet oS = oWE.oWB.getSheet(sCurSheet);
      Row oRow = oS.getRow(0);
      if (oRow == null) oRow = oS.createRow(0);
      for(int i=0,iMax=this.nMaxCol; i<iMax; i++) {
        Cell oC = oRow.getCell(i);
        if (oC == null) {
          oC = oRow.createCell(i);
          oC.setCellValue("");
        }
      }
      return this;
    }

    /**
      Adds a row of columns to the Area row array.
      <p>
      Each column string can contain a formatting mark which is parsed according to <a href={@docRoot}/overview-summary.html#format-spec>format specifier</a>.
      @returns Area for chaining purposes.
    */
    public Area addRow(ArrayList<String> oRow) throws Exception {
      oRows.add((String[])oRow.toArray(new String[oRow.size()]));
      return this;
    }

    /**
      Adds a row  of columns to the Area row array and adds a stripe mark.
      <p>
      Each column string can contain a formatting mark which is parsed according to <a href={@docRoot}/overview-summary.html#format-spec>format specifier</a>.
      @param nStripe Stripe option. 0 - no stripe, odd number - #odd background, Even number - #evn background
      @returns Area for chaining purposes.
    */
    public Area addRow(ArrayList<String> oRow,int nStripe) throws Exception {
      oRows.add((String[])oRow.toArray(new String[oRow.size()]));
      addStripe(nStripe);
      return this;
    }

    /**
      Adds a row  of columns to the Area row array along with a row format specifier.
      <p>
      Each column string can contain a formatting mark which is parsed according to <a href={@docRoot}/overview-summary.html#format-spec>format specifier</a>.
      @param sRowFmt with a row format specifier of sRowFmt (a defined style)
      @returns Area for chaining purposes.
    */
    public Area addRow(ArrayList<String> oRow,String sRowFmt) throws Exception {
      oRows.add((String[])oRow.toArray(new String[oRow.size()]));
      addColFmt(sRowFmt);
      return this;
    }

    /**
      Adds a row  of columns to the Area row array.
      <p>
      Each column string can contain a formatting mark which is parsed according to <a href={@docRoot}/overview-summary.html#format-spec>format specifier</a>.
      @returns Area for chaining purposes.
    */
    public Area addRow(String[] sRows) throws Exception {
      oRows.add(sRows);
      return this;
    }

    /**
      Adds a row  of columns to the Area row array and adds a stripe mark.
      <p>
      Each column string can contain a formatting mark which is parsed according to <a href={@docRoot}/overview-summary.html#format-spec>format specifier</a>.
      @param nStripe Stripe option. 0 - no stripe, odd number - #odd background, Even number - #evn background
      @returns Area for chaining purposes.
    */
    public Area addRow(String[] sRows,int nStripe) throws Exception {
      oRows.add(sRows);
      addStripe(nStripe);
      return this;
    }

    /**
      Adds a row  of columns to the Area row array along with a row format specifier.
      <p>
      Each column string can contain a formatting mark which is parsed according to <a href={@docRoot}/overview-summary.html#format-spec>format specifier</a>.
      @param sRowFmt with a row format specifier of sRowFmt (a defined style)
      @returns Area for chaining purposes.
    */
    public Area addRow(String[] sRows,String sRowFmt) throws Exception {
      oRows.add(sRows);
      addColFmt(sRowFmt);
      return this;
    }

    private void addStripe(int nStripe) {
      if (nStripe != 0) {
        addColFmt((nStripe % 2 == 0)?"#evn":"#odd");
      } else {
        if (sColFmts != null) {
          addColFmt(null);
        }
      }
    }

    /**
      Sets the width of a set of columns to nChars.
      <p>
      This should be called after {@link WriteExcel.Area#writeArea writeArea} is called. It sets the specifed inclusive range of columns
      (0 basde within Area) to the
      nChars width.
      @returns Area for chaining purposes.
    */
    public Area colWidth(int n1stCol,int nLastCol,int nChars) throws Exception {
      for(int i=n1stCol,iMax=nLastCol; i<=iMax; i++) {
        colWidth(i,nChars);
      }
      return this;
    }

    /**
      Sets the width of the specified column to nChars.
      <p>
      This should be called after {@link WriteExcel.Area#writeArea writeArea} is called. It sets the specifed column (0 based within Area) to the
      nChars width.
      @returns Area for chaining purposes.
    */
    public Area colWidth(int nCol,int nChars) throws Exception {
      int nAbsCol = nCol + nBaseCol;
      if ((nAbsCol >= 0) && (sCurSheet != null)) {
        Sheet oS = oWE.oWB.getSheet(sCurSheet);
        if (oS != null) oS.setColumnWidth(nAbsCol,nChars * 256);
      }
      return this;
    }

    /**
      Return next absolute row.
      @return sum of Row count plus header count plus Base row.
    */
    public int           getAbsRow()   { return getRowCount()+getHdrCount()+getBaseRow();}

    /**
      Return absolute column of start of Area on sheet.
      @return value of nCol given to the {@link WriteExcel#createArea createArea} method call that created the area
    */
    public int           getBaseCol()  { return nBaseCol;}

    /**
      Return absolute row of start of Area on sheet.
      @return value of nRow given to the {@link WriteExcel#createArea createArea} method call that created the area
    */
    public int           getBaseRow()  { return nBaseRow;}

    /**
      Return absolute row of first data row.
      @return Sum of header count plus Base row.
    */
    public int           getDataRow()  { return nBaseRow+oHdrs.size();}

    /**
      Return size of Header array.
      @return size of Rows array which is increased by {@link WriteExcel.Area#header header} method calls
    */
    public int           getHdrCount() { return oHdrs.size();}

    /**
      Return size of Rows array.
      @return size of Rows array which is increased by {@link WriteExcel.Area#addRow addRow} method calls
    */
    public int           getRowCount() { return oRows.size();}

    /**
      Gets the current row array contents.
      @returns the row array.
    */
    public ArrayList<String[]> getRows() throws Exception {
      return oRows;
    }

    /**
      Return Sheet name area.
      @return Sheet name area given to the {@link WriteExcel#createArea createArea} method call that created the area.
    */
    public String        getSheet()    { return sCurSheet;}

    /**
      Return WriteExcel for this area.
      @return WriteExcel for this area.
    */
    public WriteExcel    getWriter()   { return oWE;}

    /**
      Adds a set of columns to the header array.
      <p>
      The string sCols is split using the '/' character and then each column is stored as a value after storing any formatting
      or merge marks parsed according to <a href={@docRoot}/overview-summary.html#format-spec>format specifier</a>.
      @param sCols The array of columns, each separated by the '/' character,
      @returns Area for chaining purposes.
    */
    public Area header(String sCols) throws Exception {
      oHdrs.add(oWE.parseHeader(this,sCols,"#hdr"));
      return this;
    }

    /**
      Adds a set of columns to the header array and applies a default row format specifier.
      <p>
      The string sCols is split using the '/' character and then each column is stored as a value after storing any formatting
      or merge marks parsed according to <a href={@docRoot}/overview-summary.html#format-spec>format specifier</a>.
      @param sCols The array of columns, each separated by the '/' character,
      @param sHdrFmt The name of a defined style.
      @returns Area for chaining purposes.
    */
    public Area header(String sCols,String sHdrFmt) throws Exception {
      oHdrs.add(oWE.parseHeader(this,sCols,sHdrFmt));
      return this;
    }

    /**
      Writes the Area to sCurrent sheet.
      <p>
      After the cells are written, size calculations are performed on each column based on the data.
      @returns Area for chaining purposes.
    */
    public Area writeArea() throws Exception {
      oWE.writeArea(this,this.sCurSheet);
      return this;
    }

    /**
      Changes the column text for nHdr and nCol.
      <p>
      This should be called <b>before</b> {@link WriteExcel.Area#writeArea writeArea} is called. It changes the column text for the nRow and nCol within the Area (both 0 based).
      The text can have a format specification parsed according to <a href={@docRoot}/overview-summary.html#format-spec>format specifier</a>.
      @returns Area for chaining purposes.
    */
    public Area zapColText(int nRow,int nCol,String sText) throws Exception {
      if (nRow >= oRows.size()) return this;
      String[] sRow = oRows.get(nRow);
      if (nCol >= sRow.length) return this;
      sRow[nCol] = sText;
      return this;
    }

    /**
      Changes the header format specifier for the specified nHdr and nCol.
      <p>
      This should be called <b>before</b> {@link WriteExcel.Area#writeArea writeArea} is called. It changes the header format specifier for the nHdr and nCol within the Area (both 0 based).
      @returns Area for chaining purposes.
    */
    public Area zapHdrFmt(int nHdr,int nCol,String sFmt) throws Exception {
      HdrCol oHC = getHdrCol(nHdr,nCol);
      if (oHC == null) return this;
      oHC.sHdrFmt = sFmt;
      return this;
    }

    /**
      Changes the header text for the specified nHdr and nCol.
      <p>
      This should be called <b>before</b> {@link WriteExcel.Area#writeArea writeArea} is called. It changes the header value for the nHdr and nCol within the Area (both 0 based).
      @returns Area for chaining purposes.
    */
    public Area zapHdrText(int nHdr,int nCol,String sText) throws Exception {
      HdrCol oHC = getHdrCol(nHdr,nCol);
      if (oHC == null) return this;
      oHC.sText = sText;
      return this;
    }

    // apply rowfmt to column
    private void addColFmt(String sRowFmt) {
      if ((sRowFmt == null) && (sColFmts == null)) return;
      if (sColFmts == null) sColFmts = new ArrayList<>();
      while(sColFmts.size() < oRows.size()-1) sColFmts.add(null);
      sColFmts.add(sRowFmt);
    }

    private HdrCol getHdrCol(int nHdr,int nCol) {
      if (nHdr >= oHdrs.size()) return null;
      HdrCol[] oHCs = oHdrs.get(nHdr);
      if (nCol >= oHCs.length) return null;
      return oHCs[nCol];
    }
  }

  /** Convenient method to generate logging information written to stdout.
  */
  private static void log(String sMsg) {System.out.println(sMsg);}
  /** Convenient method to create Exception class.
  */
  private static Exception e(String s) {return new Exception(s); }
  //                                          1              2     3                4
  static Matcher oFmt  = Pattern.compile("^\\{(?<mrg>[0-9]+)?([.])?(?<fmt>[^}]*)?\\}(?<rest>.*)").matcher("");



  // -------------- Globals ----------------
  String                    sFileName;
  Workbook                  oWB;
  String                    sNegFmt;
  boolean                   bShowNegAsRed = false;
  boolean                   bDidInitStyles = false;
  FormulaEvaluator          oFE;
  HashMap<String,CellStyle> oStyMap = new HashMap<>();      // Cloned styles
  HashMap<String,CellStyle> oStyColMap = new HashMap<>();   // Color variation on styles
  //HashMap<String,CellStyle> oStyles = new HashMap<>();    // Basic styles
  Font                      oFntFix  = null;//TEMP
  String                    sFntFix  = "Courier New";       // Default Font
  String                    sFntProp = null;                // Default proportional

  HashMap<String,StyleDef>  oStyDefs  = new HashMap<>();    // Style pool
  HashMap<String,CellStyle> oStyRegs  = new HashMap<>();    // Styles registered
  ArrayList<DataFmt>        oDataFmts = new ArrayList<>();  // Data formats we support

  TreeMap<String,CellStyle> oStyleCache;                    // For regression testing
  TreeMap<String,Font>      oFontCache;

  /**
    Adds a Comment to the specified Cell using a fixed Font.
    The size of the window is calculated based on the text.
    <p>
    This method call should be after the {@link Area#writeArea} has been issued.
    @param sSheet The sheet name containing th cell.
    @param nRow The base 0 row index on the Sheet.
    @param nCol The base 0 column index on the Row.
    @param sText The comment text.
      Use a "\r\n" to cause a line break.
    @return WriteExcel for chaining purposes.
  */
  public WriteExcel addCellComment(String sSheet,int nRow,int nCol,String sText) throws Exception {
    return addCellComment(sSheet,nRow,nCol,sText,true);
  }

  /**
    Adds a Comment to the specified Cell with specified Font.
    The size of the window is calculated based on the text.
    <p>
    This method call should be after the {@link Area#writeArea} has been issued.
    @param sSheet The sheet name containing th cell.
    @param nRow The base 0 row index on the Sheet.
    @param nCol The base 0 column index on the Row.
    @param sText The comment text.
      Use a "\r\n" to cause a line break.
    @param bFont The text Font will be fixed if true, otherwise the default proportional font.
    @return WriteExcel for chaining purposes.
  */
  public WriteExcel addCellComment(String sSheet,int nRow,int nCol,String sText,boolean bFixed) throws Exception {
    Sheet oS = oWB.getSheet(sSheet);
    if (oS == null) return this;
    Row oRow = oS.getRow(nRow);
    if (oRow == null) return this;
    Font oFont = bFixed?oFntFix:null;
    addCellComment(oRow,nCol,sText,oFont);
    return this;
  }

  /**
    Adds a DataFormat type. DataFormats are matched when data is written to a cell. The value selected is used as the DataFormat seen in an Excel Cell
    when right clicking on the Cell and selecting <code>Format Cells...</code>.
    <p>
    This method call(s) should be made before any <code>Areas</code> are created using the {@link WriteExcel#createArea createArea} methods.

    <p>
    <b>Examples</b>
    <p>
    The next line of code will display with no decimal points. The <code>ND</code> is the signal.
    It will match any number of decimals but there must be at least one.
    <p>
    <code>addDataFormat("nodec","^ND-?[0-9]+[.][0-9]+$","0",)</code>
    <p>
    A value of <code>"ND10.333"</code> or <code>"ND-40.1"</code> would select this format;
    <p>
    The next line of code will display with two decimal points and negative coloring.
    It will match any number of decimals but there must be at least one.
    <p>
    <code>addDataFormat(neg-nodec","^NND-?[0-9]+[.][0-9]+$","0.00;[Red]0.00")</code>
    <p>
    A value of <code>"NND10.333"</code> or <code>"NND-40.1"</code> would select this format;

    @param sName The DataFormat name.  Do not use names starting with "@" as these are internal name.
    These names are only used in the logs created as style combinations
    are registered.
    @param sMatcher The Regex string to match the value being written to the Cell. It should match the entire value.
    @param sDataFmt The value used as the DataFormat for the CellStyle.
    <p>
    When processing the pattern match, they are executed in the sequence of the calls to method and before any standard DataFormats
    seen at <a href={@docRoot}/overview-summary.html#generic-types>Standard Types</a>.
    <p>
    The matched Cell data is assumed to be a double or integer with formatting aids such as $ and , puncuation marks and an optional prefix signal..
    These are removed when the parsing for the double or integer value takes place.
    <p>Thus an alphabetic 'signal' string prefix can be used to coerce the match to select a particular formatting style.

    @return WriteExcel for chaining purposes.

    @see
    <a href={@docRoot}/overview-summary.html#generic-types>Standard Types</a>.
    @see
    <a href={@docRoot}/overview-summary.html#style-choice>Choosing a Cell Style</a>.

  */
  public WriteExcel addDataFormat(String sName,String sMatcher,String sDataFmt) throws Exception {
    createDataFormat(sName,sMatcher,sDataFmt);
    return this;
  }

  /**
    This clones the source Sheet into the Workbook being constructed as a Sheet named <code>sSheet</code>.
    <p>
    This is primarily used to create new Sheets from a template. The source Sheet can be the same Workbook or from a Workbook
    opened with {@link com.sec.excel.ReadExcel ReadExcel}.
    <p>
    The cloning process prevserves the original column widths, merged cells and formatting.
    <p>
    The <code>Print Area</code>, when specified, will be adjusted to refelect the new Sheet name.  All Print Properties in the
    original Sheet are also transcribed.
  */
  public void addExternalSheet(String sSheet,Sheet oS,String sPrintArea) throws Exception {
    Sheet oNewS = oWB.getSheet(sSheet);
    if (oNewS != null) throw e("Sheet "+sSheet+" already exists");
    oNewS = oWB.createSheet(sSheet);
    int nMaxCol = 0;
    for(int i=oS.getFirstRowNum(),iMax = oS.getLastRowNum(); i <= iMax; i++) {
      Row oRow = oS.getRow(i);
      if (oRow != null) {
        nMaxCol = copyRow(oNewS,i,oRow,nMaxCol);
      }
    }

    for(int i=0,iMax = nMaxCol; i <= iMax; i++) {
      oNewS.setColumnWidth(i,oS.getColumnWidth(i));
    }

    for(CellRangeAddress oCRA:oS.getMergedRegions()) {
      oNewS.addMergedRegion(oCRA);
    }

    clonePrintSetup(oNewS,oS,sPrintArea);
    cloneChartObject(oNewS,oS);
  }

  /**
    Registers a named style. The named style once registered can be referenced as described in
    <a href={@docRoot}/overview-summary.html#style-choice>Choosing a Cell Style</a>.
    <p>
    Names starting with '#' are builtin Style Names and should not be used unless purposefully overriding
    a builtin Style Name. They are listed at <a href={@docRoot}/overview-summary.html#builtin-names>Builtin Style Names</a>.
    <p>
    Names are case sensitive and can contain alphanumeric characters as well as the - and _ character.
    <p>
    Must be called before calls {@link WriteExcel#createArea createArea} that reference the Style.  Changes to a Style definition will be ignored
    once it is referenced when a Cell using the style is written to the Workbook.

  @param sName The name to register.
  @param sStyStr The cell style orders as described in <a href={@docRoot}/overview-summary.html#cell-orders>Cell Style Orders Syntax</a>.

    @see
    <a href={@docRoot}/overview-summary.html#builtin-names>Builtin Style Names</a>.
    @see
    <a href={@docRoot}/overview-summary.html#style-choice>Choosing a Cell Style</a>.
  */
  public WriteExcel addStyleDefn(String sName,String sStyStr) throws Exception {
    insertStyleDef(sName,sStyStr);
    return this;
  }

  /**
    Creates an empty Sheet called sName.
    <p>
    It can happen that the processing sequence is different than the preferred Sheet sequence in the resulting Workbook.
    Calling this method allows the Sheet to be created in the desired sequence.
    @return WriteExcel for chaining purposes.
  */
  public WriteExcel bookSheet(String sName) throws Exception {
    oWB.createSheet(sName);
    return this;
  }

  /**
    Return  a stringized summary of the Cell. It caches the Style and Font indexes so they can be dumped later.
    The data value and data format are returned as strings. Used for creating regression test logs.
    @param oC Cell to examine
    @return a summary of the Cell. If oC is null returns null.
  */
  public String cellSummary(Cell oC) throws Exception {
    if (oC == null) return null;
    if (oStyleCache == null) {
      oStyleCache = new TreeMap<>();
      oFontCache  = new TreeMap<>();
    }
    return cellAsString(oC);
  }

  /**
    Return  a stringized summary of the CellStyle cache. Created by the cellSummary method calls and this should be called
    as the last step.
    @return a summary of the CellStyle cache, one for each CellStyle index value.
  */
  public String[] dumpCellStyleCache() throws Exception {
    if (oStyleCache == null) return new String[0];
    String[] oArr = new String[oStyleCache.size()];
    int nIX = -1;
    for(String sKey:oStyleCache.keySet().toArray(new String[oStyleCache.size()])) {
      nIX += 1;
      CellStyle oCS =  oStyleCache.get(sKey);
      VerticalAlignment oVA = oCS.getVerticalAlignment();
      String sVA = "v";
      String sHA = "h";
      if (oVA != null) {
        switch(oVA) {
          case BOTTOM:      sVA = "B"; break;
          case CENTER:      sVA = "C"; break;
          case DISTRIBUTED: sVA = "D"; break;
          case JUSTIFY:     sVA = "J"; break;
          case TOP:         sVA = "T"; break;
        }
      }
      HorizontalAlignment oHA = oCS.getAlignment();
      if (oHA != null) {
        switch(oHA) {
          case CENTER:           sHA = "C"; break;
          case CENTER_SELECTION: sHA = "S"; break;
          case DISTRIBUTED:      sHA = "D"; break;
          case FILL:             sHA = "F"; break;
          case GENERAL:          sHA = "G"; break;
          case JUSTIFY:          sHA = "J"; break;
          case LEFT:             sHA = "L"; break;
          case RIGHT:            sHA = "R"; break;
        }
      }
      oArr[nIX] = String.format("%5d %s %s tblr(%s) fFG(%s,%s,%s)",
        Integer.parseInt(sKey),sHA,sVA,getBorders(oCS),""+oCS.getFillPattern(),colorMap(oCS.getFillForegroundColor()),colorMap(oCS.getFillBackgroundColor()));
    }
    return oArr;
  }

  private String getBorders(CellStyle oCS) {
    StringBuilder oSB = new StringBuilder();
    oSB.append(bsAsStr(oCS.getBorderTop())       +"."+colorMap(oCS.getTopBorderColor()));
    oSB.append(","+bsAsStr(oCS.getBorderBottom())+"."+colorMap(oCS.getBottomBorderColor()));
    oSB.append(","+bsAsStr(oCS.getBorderLeft())  +"."+colorMap(oCS.getLeftBorderColor()));
    oSB.append(","+bsAsStr(oCS.getBorderRight()) +"."+colorMap(oCS.getRightBorderColor()));
    return ""+oSB;
  }

  private String bsAsStr(BorderStyle oBS) {
    if (oBS == null) return "no";
    switch(oBS) {
      case DASH_DOT:               return "D1";
      case DASH_DOT_DOT:           return "D2";
      case DASHED:                 return "D3";
      case DOTTED:                 return "T1";
      case DOUBLE:                 return "DB";
      case HAIR:                   return "HR";
      case MEDIUM:                 return "M1";
      case MEDIUM_DASH_DOT:        return "M2";
      case MEDIUM_DASH_DOT_DOT:    return "M3";
      case MEDIUM_DASHED:          return "M4";
      case NONE:                   return "NO";
      case SLANTED_DASH_DOT:       return "S1";
      case THICK:                  return "K1";
      case THIN:                   return "N1";
    }
    return "??";
  }

  private String colorMap(short nCol) {
    IndexedColorMap oICM = ((XSSFWorkbook)oWB).getStylesSource().getIndexedColors();
    byte[] oB = oICM.getRGB(nCol);
    StringBuilder oSB = new StringBuilder();
    if (oB == null) {
      oSB.append("------");
    } else {
      for(int i=0,iMax=oB.length; i<iMax; i++) {
        oSB.append(String.format("%02x",((int)oB[i]) & 0xFF));
      }
    }
    return String.format("%04d.%s",nCol,""+oSB);
  }

  /**
    Allow custom modifications to CellStyle. This exit is called if the
    <a href={@docRoot}/overview-summary.html#cell-orders>CellStyle orders</a> contains a <code>CE(...)</code> order.
    <p>
    It allows custom CellStyle changes to be made to the CellStyle being constructed that are not covered by the order set.  It is called just prior to caching
    the CellStyle after all other orders have been processed.
    @param oCS <code>CellStyle</code> being constructed.
    @param sStr The string contained in the <code>CE(...)</code> order.  This allows different settings based on the string passed in.
  */
  public void customExit(CellStyle oCS,String sStr) {
  }

  /**
    Return  a stringized summary of the CellStyle Font cache. Created by the cellSummary method calls and this should be called
    as the last step.
    @return a summary of the CellStyle font cache, one for each CellStyle font index value.
  */
  public String[] dumpCellFontCache() throws Exception {
    if (oFontCache == null) return new String[0];
    String[] oArr = new String[oFontCache.size()];
    int nIX = -1;
    for(String sKey:oFontCache.keySet().toArray(new String[oFontCache.size()])) {
      nIX += 1;
      Font oFont =  oFontCache.get(sKey);
      oArr[nIX] = String.format("%5d %-20s %5d %s%s%s u=%02d %s",Integer.parseInt(sKey),oFont.getFontName(),oFont.getFontHeight()
        ,oFont.getBold()?"b":"."
        ,oFont.getItalic()?"i":"."
        ,oFont.getStrikeout()?"s":"."
        ,(int)oFont.getUnderline()
        ,colorMap(oFont.getColor()));
    }
    return oArr;
  }

  /**
    Close the Workbook. The Workbook is written and the <code>FileOutputStream</code> closed. No further changes can be made.
  */
  public void close() throws Exception {
    FileOutputStream oFOS = new FileOutputStream(sFileName);
    oWB.write(oFOS);
    oWB.close();
  }

  // -------------- Instantiators ----------------
  /** Creates an instance of <code>WriteExcel</code> that will write sFileName.
     @param oWE the parent instance that subclasses WriteExcel;
     @param sFileName The fully qualified file path and name suitable for use in a FileOutputStrem.
     @return The created instance.
  */
  public static WriteExcel create(WriteExcel oWE,String sFileName) throws Exception {
    return WriteExcel.create(oWE,sFileName,null);
  }

  /**
    Create an instance of <code>WriteExcel</code> that will write sFileName and uses sSrcFile as a template file.
    @param oWE the parent instance that subclasses WriteExcel;
    @param sFileName The fully qualified file path and name suitable for use in a FileOutputStrem.
    @param sSrcName The input .xlsx file that is to be used as a template.
    @return The created instance.
  */

  public static WriteExcel create(WriteExcel oWE,String sFileName,String sSrcName) throws Exception {
    oWE.sFileName = sFileName;
    if (sSrcName == null) {
      oWE.oWB = new XSSFWorkbook();
    } else {
      oWE.oWB = new XSSFWorkbook(new FileInputStream(sSrcName));
    }
    oWE.oFE = oWE.oWB.getCreationHelper().createFormulaEvaluator();
    return oWE;
  }

  /**
    Internal access to Cell manipulation routines. Used in regression testing.
    @param oWB The Workbook we are accessing
    @return WriteExcel in read only mode
  */
  public static WriteExcel create(Workbook oWB) {
    WriteExcel oWE = new WriteExcel();
    oWE.oWB = oWB;
    oWE.oFE = oWB.getCreationHelper().createFormulaEvaluator();
    return oWE;
  }

  /**
    Creates an <code>Area</code> with a base Row and Column of 0 on the specified Sheet.  The Sheet is created if it does not exist.
    @param sSheet The Sheet name for the <code>Area</code>.
    @return The created <code>Area</code>.
  */
  public Area createArea(String sSheet) throws Exception {
    return createArea(sSheet,0,0);
  }

  /**
    Creates an <code>Area</code> with the specified base Row and Column on the specified Sheet.  The Sheet is created if it does not exist.
    @param sSheet The Sheet name for the <code>Area</code>.
    @param nRow The 0 based row index.
    @param nCol The 0 based column index.
    @return The created <code>Area</code>.
  */
  public Area createArea(String sSheet,int nRow,int nCol) throws Exception {
    if (!bDidInitStyles) {
      bDidInitStyles = true;
      createStandardStylePods();
    }
    Area oA = new Area();
    oA.oWE = this;
    oA.sCurSheet = sSheet;
    oA.nBaseRow  = nRow;
    oA.nBaseCol  = nCol;
    return oA;
  }

  /**
    Gets the standard header builtins styles.
    <p>
    Each String returned is of the format <code>#name:stystr</code> where
    <ul>
      <li><code>name</code> is the style name prefixed with a <code>#</code> to mark it as standard style</li>
      <li><code>stystr</code> is the string orders as defined in <a href={@docRoot}/overview-summary.html#cell-orders>CellStyle orders</a>.</li>
    </ul>
  */
  public String[] getHdrExtras() {
    return "#hdr:bC/#title:b16C/#hdrBlue:bCBG(pale-blue)".split("/");
  }

  /**
    Gets the Row count on the specified Sheet.
    @param sSheet The sheet name to inspect.
    @return The count.
  */
  public int getRowCount(String sSheet) {
    Sheet oS = oWB.getSheet(sSheet);
    if (oS == null) return 0;
    return oS.getLastRowNum() + 1;
  }

  /**
    Gets the standard builtins styles.
    <p>
    Each String returned is of the format <code>#name:stystr</code> where
    <ul>
      <li><code>name</code> is the style name prefixed with a <code>#</code> to mark it as standard style</li>
      <li><code>stystr</code> is the string orders as defined in <a href={@docRoot}/overview-summary.html#cell-orders>CellStyle orders</a>.</li>
    </ul>
  */
  public String[] getStdExtras() {
    return "#sub:-/#tot:~/#fin:=/#lnk:l/#lkc:lC".split("/");
  }

  /**
    Gets the String value of the specified Cell.
    @param sSheet The sheet name to inspect.
    @param nRow The base 0 row index on the Sheet.
    @param nCol The base 0 column index on the Row.
    @return The value as stored in the Cell. It could be <code>null</code>.
  */
  public String getStrValue(String sSheet,int nRow,int nCol) throws Exception {
    Sheet oS = oWB.getSheet(sSheet);
    if (oS == null) return null;
    Row oRow = oS.getRow(nRow);
    if (oRow== null) return null;
    Cell oC = oRow.getCell(nCol);
    if (oC == null) return null;
    return getCellAsStr(oC);
  }

  /**
    Gets the <code>Workbook</code> being constructed.
    <p>
    This allows inspection or modifications to be made to the Workbook using the POI library directly. Caution is advised.
    Calling this method allows the Sheet to be created in the desired sequence.
    @return Workbook.
  */
  public Workbook getWorkbook() throws Exception {
    return oWB;
  }

  /**
  Make File link entry. Creates a link entry that when clicked will open the file using the Windows defaults to find a program to open it with.
  <p>
  The sheet must exist that will be linked before this method can be called.
  @param sTargLinkSty Style of index to target link.
  A <code>null</code> value is treated as <code>#lkc</code>.
  <p>
  The builtin <code>#lkc</code> is centered underlined blue.
  The builtin <code>#lnk</code> is non-centered underlined blue.
  <p>
  Other style values may be used provided they have been added with the {@link WriteExcel#addStyleDefn addStyleDefn} method
  @param sSheet Source sheet name on which to create link
  @param sIdxRow Link row on source page.
  @param sIdxCol Link column on source page.
  @param sStr The link text
  @param sFileName The File path passed to the Windows routines.
  */

  public void makeFileLink(String sLinkSty,String sSheet,int nRow,int nCol,String sStr,String sFileName) throws Exception {
    Sheet oS = oWB.getSheet(sSheet);
    if (oS == null) return;
    Row oRow = oS.getRow(nRow);
    if (oRow == null) oRow = oS.createRow(nRow);
    CreationHelper oCH = oWB.getCreationHelper();
    Cell oC = oRow.createCell(nCol);
    oC.setCellValue(sStr);
    oC.setCellStyle(useLinkStyle(sStr,sLinkSty));
    Hyperlink oL = oCH.createHyperlink(HyperlinkType.FILE);
    oL.setAddress(sFileName);
    oC.setHyperlink(oL);
    return;
  }

  /**
  Make index link entry. Creates a bi-directional link entry to and from a page called "index".
  <p>
  This method allows for the creation of an index page which is useful in complicated spreadsheets.
  <p>
  The sheet must exist that will be linked before this method can be called.
  If necessary 'defer' logic should be added to the creating program as sheets are created
  that need links.  The {@link WriteExcel#bookSheet bookSkeet} method can be called to position the "index" sheet as the first sheet in the Workbook.
  <p>
  It creates a bi-directional link to the target page from the index page by calling method
  {@link WriteExcel#makeIndexLink(String,String,String,int,int,String,int,int) makeIndexLink}
  with <code>nIdxLnkRow</code>=1 and <code>nIdxLnkCol</code>=1.
  @param sTargLinkSty Style of index to target link.
  A <code>null</code> value is treated as <code>#lkc</code>.
  <p>
  The builtin <code>#lkc</code> is centered underlined blue.
  The builtin <code>#lnk</code> is non-centered underlined blue.
  <p>
  Other style values may be used provided they have been added with the {@link WriteExcel#addStyleDefn addStyleDefn} method
  @param sTargSheet Sheet name to index
  @param sIdxName Link name on index page.
  @param sIdxRow Link row on index page.
  @param sIdxCol Link column on index page.
  @param sIdxLinkSty Style of target to index link.
  A <code>null</code> value will use the <code>sTargLinkSty</code> value or <code>#lkc</code>.
  <p>
  The builtin <code>#lkc</code> is centered underlined blue.
  The builtin <code>#lnc</code> is non-centered underlined blue.
  <p>
  Other style values may be used provided they have been added with the {@link WriteExcel#addStyleDefn addStyleDefn} method
  */
  public void makeIndexLink(String sTargLinkSty,String sTargSheet,String sIdxName,int nIdxRow,int nIdxCol,String sIdxLinkSty) throws Exception {
    makeIndexLink(sTargLinkSty,sTargSheet,sIdxName,nIdxRow,nIdxCol,sIdxLinkSty,1,1);
  }

  /**
  Make index link entry. Creates a bi-directional link entry to and from a page called "index".
  <p>
  This method allows for the creation of an index page which is useful in complicated spreadsheets.
  <p>
  The sheet must exist that will be linked before this method can be called.
  If necessary 'defer' logic should be added to the creating program as sheets are created
  that need links.  The {@link WriteExcel#bookSheet bookSkeet} method can be called to position the "index" sheet as the first sheet in the Workbook.
  <p>
  It creates a bi-directional link to the target page from the index page.

  It creates the index pointer on the target page at row <code>nIdxLnkRow</code> column <code>nIdxLnkCol</code>.
  @param sTargLinkSty Style of index to target link.
  A <code>null</code> value is treated as <code>#lkc</code>.
  <p>
  The builtin <code>#lkc</code> is centered underlined blue.
  The builtin <code>#lnk</code> is non-centered underlined blue.
  <p>
  Other style values may be used provided they have been added with the {@link WriteExcel#addStyleDefn addStyleDefn} method
  @param sTargSheet Sheet name to index
  @param sIdxName Link name on index page.
  @param sIdxRow Link row on index page.
  @param sIdxCol Link column on index page.
  @param sIdxLinkSty Style of target to index link.
  A <code>null</code> value will use the <code>sTargLinkSty</code> value or <code>#lkc</code>.
  <p>
  The builtin <code>#lkc</code> is centered underlined blue.
  The builtin <code>#lnc</code> is non-centered underlined blue.
  <p>
  Other style values may be used provided they have been added with the {@link WriteExcel#addStyleDefn addStyleDefn} method
  @param sIdxLnkRow Link row on target page.
  @param sIdxLnkCol Link column on target page.

  */

  public void makeIndexLink(String sTargLinkSty,String sTargSheet,String sIdxName,int nIdxRow,int nIdxCol,String sIdxLinkSty,int nIdxLnkRow,int nIdxLnkCol) throws Exception {
    Sheet oIdxSheet = oWB.getSheet("index");
    Row oIdxRow = oIdxSheet.getRow(nIdxRow);
    if (oIdxRow == null) oIdxRow = oIdxSheet.createRow(nIdxRow);
    CreationHelper oCH = oIdxRow.getSheet().getWorkbook().getCreationHelper();
    Cell oIdxCell = oIdxRow.createCell(nIdxCol);
    oIdxCell.setCellValue(sIdxName);
    oIdxCell.setCellStyle(useLinkStyle(sIdxName,sTargLinkSty));
    Hyperlink oTargLnk = oCH.createHyperlink(HyperlinkType.DOCUMENT);
    oTargLnk.setAddress("'"+sTargSheet+"'!"+(new CellAddress(oIdxCell)).toString());
    oIdxCell.setHyperlink(oTargLnk);

    Cell oTargCell = oIdxRow.getSheet().getWorkbook().getSheet(sTargSheet).getRow(nIdxLnkRow).createCell(nIdxLnkCol);
    Hyperlink oIdxLnk = oCH.createHyperlink(HyperlinkType.DOCUMENT);
    oIdxLnk.setAddress("'index'!"+(new CellAddress(oIdxCell)).toString());
    oTargCell.setCellValue("index");
    oTargCell.setCellStyle(useLinkStyle(sIdxName,sIdxLinkSty!=null?sIdxLinkSty:sTargLinkSty));
    oTargCell.setHyperlink(oIdxLnk);

  }

  /**
  Make standard link entry. Creates a bi-directional link entry between the source and target sheets where The link value is the other sheet name.
  <p>
  The sheets must exist that will be linked before this method can be called.
  If necessary 'defer' logic should be added to the creating program as sheets are created
  that need links.
  <p>
  The target and source parameters in this method are commutative.
  @param sTargLinkSty Style of target to source link.
  A <code>null</code> value is treated as <code>#lkc</code>.
  <p>
  The builtin <code>#lkc</code> is centered underlined blue.
  The builtin <code>#lnk</code> is non-centered underlined blue.
  <p>
  Other style values may be used provided they have been added with the {@link WriteExcel#addStyleDefn addStyleDefn} method
  @param sTargSheet Target Sheet name
  @param sTargRow Link row on target page.
  @param sTargCol Link column on target page.
  @param sSrcLinkSty Style of source to link.
  A <code>null</code> value is treated as <code>#lkc</code>.
  <p>
  The builtin <code>#lkc</code> is centered underlined blue.
  The builtin <code>#lnk</code> is non-centered underlined blue.
  <p>
  Other style values may be used provided they have been added with the {@link WriteExcel#addStyleDefn addStyleDefn} method
  @param sSrcSheet Source Sheet name
  @param sSrcRow Link row on source page.
  @param sSrcCol Link column on source page.

  */
  public void makeStdLink(String sTargLinkSty,String sTargSheet,int nTargRow,int nTargCol,String sSrcLinkSty,String sSrcSheet,int nSrcRow,int nSrcCol) throws Exception {
    Sheet oSrcSheet = oWB.getSheet(sSrcSheet);
    Row oSrcRow = oSrcSheet.getRow(nSrcRow);
    if (oSrcRow == null) oSrcRow = oSrcSheet.createRow(nSrcRow);
    CreationHelper oCH = oSrcRow.getSheet().getWorkbook().getCreationHelper();
    Cell oSrcCell = oSrcRow.createCell(nSrcCol);
    oSrcCell.setCellValue(sTargSheet);
    oSrcCell.setCellStyle(useLinkStyle(sTargSheet,sTargLinkSty));
    Hyperlink oTargLnk = oCH.createHyperlink(HyperlinkType.DOCUMENT);
    oTargLnk.setAddress("'"+sTargSheet+"'!"+"ABCDEFGHIJKLMNOPQRSTUVXYZ".substring(nTargCol,nTargCol+1)+(nTargRow+1));
    oSrcCell.setHyperlink(oTargLnk);

    Cell oTargCell = oSrcRow.getSheet().getWorkbook().getSheet(sTargSheet).getRow(nTargRow).createCell(nTargCol);
    Hyperlink oSrcLnk = oCH.createHyperlink(HyperlinkType.DOCUMENT);
    oSrcLnk.setAddress("'"+sSrcSheet+"'!"+"ABCDEFGHIJKLMNOPQRSTUVXYZ".substring(nSrcCol,nSrcCol+1)+(nSrcRow+1));
    oTargCell.setCellValue(sSrcSheet);
    oTargCell.setCellStyle(useLinkStyle(sSrcSheet,sSrcLinkSty));
    oTargCell.setHyperlink(oSrcLnk);
  }

  /**
  Make a one directional link from the source location to the target location.
  <p>
  The sheets must exist that will be linked before this method can be called.
  If necessary 'defer' logic should be added to the creating program as sheets are created
  that need links.
  <p>
  Implemented by calling {@link WriteExcel#makeUniLink(String,String,int,int,String,int,int,String,int) makeUniLink} with an <code>nRows</code> value of <code>1</code>.
  @param sLinkSty Style of link.
  A <code>null</code> value is treated as <code>#lkc</code>.
  <p>
  The builtin <code>#lkc</code> is centered underlined blue.
  The builtin <code>#lnk</code> is non-centered underlined blue.
  <p>
  Other style values may be used provided they have been added with the {@link WriteExcel#addStyleDefn addStyleDefn} method
  @param sTargSheet Target Sheet name
  @param sTargRow Link row on target page.
  @param sTargCol Link column on target page.
  @param sSrcSheet Source Sheet name
  @param sSrcRow Link row on source page.
  @param sSrcCol Link column on source page.
  @param sSrcText Text value to put in link text on the source sheet.
  */
  public void makeUniLink(String sLinkSty,String sTargSheet,int nTargRow,int nTargCol,String sSrcSheet,int nSrcRow,int nSrcCol,String sSrcText) throws Exception {
    makeUniLink(sLinkSty,sTargSheet,nTargRow,nTargCol,sSrcSheet,nSrcRow,nSrcCol,sSrcText,1);
  }
  /**
  Make a one directional link from the source location to the target location.
  <p>
  The sheets must exist that will be linked before this method can be called.
  If necessary 'defer' logic should be added to the creating program as sheets are created
  that need links.
  <p>
  When the link is pressed the target location is selected.  Multiple rows are selected if the parameter nRows is greated than 1.
  @param sLinkSty Style of link.
  A <code>null</code> value is treated as <code>#lkc</code>.
  <p>
  The builtin <code>#lkc</code> is centered underlined blue.
  The builtin <code>#lnk</code> is non-centered underlined blue.
  <p>
  Other style values may be used provided they have been added with the {@link WriteExcel#addStyleDefn addStyleDefn} method
  @param sTargSheet Target Sheet name
  @param sTargRow Link row on target page.
  @param sTargCol Link column on target page.
  @param sSrcSheet Source Sheet name
  @param sSrcRow Link row on source page.
  @param sSrcCol Link column on source page.
  @param sSrcText Text value to put in link text on the source sheet.
  @param nRows number of rows to select in target location.
  */
  public void makeUniLink(String sLinkSty,String sTargSheet,int nTargRow,int nTargCol,String sSrcSheet,int nSrcRow,int nSrcCol,String sSrcText,int nRows) throws Exception {
    Sheet oSrcSheet = oWB.getSheet(sSrcSheet);
    Row oSrcRow = oSrcSheet.getRow(nSrcRow);
    if (oSrcRow == null) oSrcRow = oSrcSheet.createRow(nSrcRow);
    CreationHelper oCH = oSrcRow.getSheet().getWorkbook().getCreationHelper();
    Cell oSrcCell = oSrcRow.createCell(nSrcCol);
    oSrcCell.setCellValue(sSrcText);
    oSrcCell.setCellStyle(useLinkStyle(sSrcSheet,sLinkSty));
    Hyperlink oTargLnk = oCH.createHyperlink(HyperlinkType.DOCUMENT);
    String sAddr = "'"+sTargSheet+"'!"+"ABCDEFGHIJKLMNOPQRSTUVXYZ".substring(nTargCol,nTargCol+1)+(nTargRow+1);
    if (nRows > 1) {
      sAddr += ":"+"ABCDEFGHIJKLMNOPQRSTUVXYZ".substring(nTargCol,nTargCol+1)+(nTargRow+nRows);
    }
    oTargLnk.setAddress(sAddr);
    oSrcCell.setHyperlink(oTargLnk);
  }

  /**
  Make URL link entry. Creates a link entry that when clicked will open the default browser with the URL as the target address.
  <p>
  The sheet must exist that will be linked before this method can be called.
  @param sTargLinkSty Style of index to target link.
  A <code>null</code> value is treated as <code>#lkc</code>.
  <p>
  The builtin <code>#lkc</code> is centered underlined blue.
  The builtin <code>#lnk</code> is non-centered underlined blue.
  <p>
  Other style values may be used provided they have been added with the {@link WriteExcel#addStyleDefn addStyleDefn} method
  @param sSheet Source sheet name on which to create link
  @param sIdxRow Link row on source page.
  @param sIdxCol Link column on source page.
  @param sStr The link text
  @param sUrlName The URL passed to the browser
  */
  public void makeUrlLink(String sLinkSty,String sSheet,int nRow,int nCol,String sStr,String sUrlName) throws Exception {
    Sheet oS = oWB.getSheet(sSheet);
    if (oS == null) return;
    Row oRow = oS.getRow(nRow);
    if (oRow == null) oRow = oS.createRow(nRow);
    CreationHelper oCH = oWB.getCreationHelper();
    Cell oC = oRow.createCell(nCol);
    oC.setCellValue(sStr);
    oC.setCellStyle(useLinkStyle(sStr,sLinkSty));
    Hyperlink oL = oCH.createHyperlink(HyperlinkType.URL);
    oL.setAddress(sUrlName);
    oC.setHyperlink(oL);
    return;
  }

  /**
    Refresh the calculated value in a specific Cell.
    <p>
    WriteExcel uses the POI library provided FormulaEvaluator to maintain the current values
    of cells containing formulae.  This method invokes the FormulaEvaluator instance for this Workbook
    for the specific cell.
    <p>
    This should be used if referenced values have been changed before obtaining the current value of a cell
    that contains calculated values.
    <p>
    If the refenced Cell does not exist this call is ignored.
    @param sSheet Sheet name
    @param nRow   Row number relative to 0
    @param nCol   Column number relative to 0
  */
  public void refreshCell(String sSheet,int nRow,int nCol) {
    Sheet oS = oWB.getSheet(sSheet);
    if (oS == null) return;
    Row oRow = oS.getRow(nRow);
    if (oRow== null) return;
    Cell oC = oRow.getCell(nCol);
    if (oC == null) return;
    oFE.evaluateInCell(oC);
  }

  /**
    Refresh the calculated values of all cells.
    <p>
    WriteExcel uses the POI library provided FormulaEvaluator to maintain the current values
    of cells containing formulae.  This method invokes the FormulaEvaluator instance for this Workbook
    for all cells.
    <p>
    This should be used if referenced values have been changed before obtaining the current value of a cell
    that contains calculated values.
    <p>
  */
  public void refreshCells() {
    oFE.evaluateAll();
  }

  /**
    Set how negative numbers are displayed. Calling this method determines the default for
    how negative numbers are displayed.
    <p>
    This can be changed for a specific format using {@link WriteExcel#addGenericType addGenericType} with
    an appropriate signal prefix as shown in the example.
    <p>
    Must be called before calls to {@link WriteExcel#addStyles addStyles} or {@link WriteExcel#createArea createArea}.
    @param bShowNegAsRed Negative numbers will also be shown in red
    @param sNegFmt Valid combinations as shown below:
    <ul>
    <li> null : default variation (color option ignored)</li>
    <li> ""&nbsp;&nbsp;&nbsp;   : color option with no prefix</li>
    <li>"-"&nbsp;&nbsp;   : - prefix plus color option</li>
    <li>"()"&nbsp;  : Paren wrapper plus color option</li>
    </ul>
  */
  public void setNegativeFormat(boolean bShowNegAsRed,String sNegFmt) throws Exception {
    this.bShowNegAsRed = bShowNegAsRed;
    this.sNegFmt       = sNegFmt;
  }

  /**
    Make changes to specified Cell data with the CellStyle determined according to <code>sVal</code>.
    @param sSheet The sheet name containing the cell.
    @param nRow The base 0 row index on the Sheet.
    @param nCol The base 0 column index on the Row.
    @param sVal The replacement Cell text. This is parsed according to <a href={@docRoot}/overview-summary.html#format-spec>format specifier</a>.
    @return WriteExcel for chaining purposes.
  */
  public void zapCell(String sSheet,int nRow,int nCol,String sVal) throws Exception {
    zapCell(sSheet,nRow,nCol,sVal,false);
  }

  /**
    Make changes to specified Cell data with optional CellStyle preservation.
    @param sSheet The sheet name containing the cell.
    @param nRow The base 0 row index on the Sheet.
    @param nCol The base 0 column index on the Row.
    @param sVal The replacement Cell text. This is parsed according to <a href={@docRoot}/overview-summary.html#format-spec>format specifier</a>.
    This is ignored if <code>bKeepStyle</code> is true.
    @param boolean bKeepStyle The existing style format is preserved if true. Otherwise a new format is determined according to the parsing of <code>sVal</code>.
    @return WriteExcel for chaining purposes.
  */
  public void zapCell(String sSheet,int nRow,int nCol,String sVal,boolean bKeepStyle) throws Exception {
    Sheet oS = oWB.getSheet(sSheet);
    if (oS == null) return;
    Row oRow = oS.getRow(nRow);
    if (oRow== null) oRow = oS.createRow(nRow);
    CellStyle oSty = null;
    if (bKeepStyle) {
      Cell oCell = oRow.getCell(nCol);
      if (oCell != null) oSty = oCell.getCellStyle();
    }
    setCellContent(null,oRow,nCol,""+sVal);
    if (oSty != null) oRow.getCell(nCol).setCellStyle(oSty);
  }

  // -------------- Style Management ----------------

  private void createStandardStylePods() throws Exception {
    for(String sStr:getStdTypes()) {
      String[] sParts = sStr.split(":");
      if (sParts.length > 1) {
        createDataFormat(sParts[0],sParts[2],sParts[1]);
      } else {
        createDataFormat(sParts[0],null,null);
      }
    }
    // other standard formats
    for(String sStr:getHdrExtras()) {
      String[] sParts = sStr.split(":");
      insertStyleDef(sParts[0],sParts[1]);
    }
    // other common variations
    for(String sStr:getStdExtras()) {
      String[] sParts = sStr.split(":");
      insertStyleDef(sParts[0],sParts[1]);
    }
    // Standard Background colors
    for(String sStr:getStdBGs()) {
      String[] sParts = sStr.split(":");
      StyleDef oSD = insertStyleDef(sParts[0],sParts[1]);
      if (!oSD.oSA.bBackOnly) {
        log("BG Style "+sParts[0]+" "+sParts[1]+" has non-color contaminants");
      }
    }
  }

  // name : sFmtStr : pattern / etc
  private String[] getStdTypes() {
    return "@int:0:^-?[0-9]+$/@nm1:0.0:^-?[0-9]+\\.[0-9]$/@num:0.00:^-?[0-9]+\\.[0-9][0-9]$/@nm3:0.000:^-?[0-9]+\\.[0-9]{3}$/@nm4:0.0000:^-?[0-9]+\\.[0-9]{4}$/@str".split("/");
  }

  private String[] getStdBGs() {
    return "#evn:BG(245,255,245)/#odd:BG(245,245,255)/#qtr:BG(255,200,145)/#TOT:BG(255,168,80)".split("/");
  }

  private StyleDef insertStyleDef(String sName,String sStyStr) throws Exception {
    log("InsertDef "+sName+" "+sStyStr);
    StyleDef oSD = new StyleDef();
    oSD.sName = sName;
    oSD.sStyStr = sStyStr;
    oSD.oSA = parseStyleAttrs(oSD);
    oStyDefs.put(oSD.sName,oSD);
    return oSD;
  }

  private void createDataFormat(String sName,String sMatcher,String sDataFmt) throws Exception {
    DataFmt oDF = new DataFmt();
    oDF.sName = sName;
    if (sMatcher != null) {
      oDF.oM = Pattern.compile(sMatcher).matcher("");
    }
    oDF.sDataFmt = sDataFmt;
    if (sDataFmt != null) {
      oDF.bInteger = !(sDataFmt.contains("."));
      if ((!sDataFmt.contains(";")) && (bShowNegAsRed || (sNegFmt != null))) {
        if (sNegFmt != null) {
          switch(sNegFmt) {
            case "":   oDF.sDataFmt += ";"+(bShowNegAsRed?"[Red]":"")+oDF.sDataFmt;          break; // no negative indicator
            case "-":  oDF.sDataFmt += ";"+(bShowNegAsRed?"[Red]":"")+"-"+oDF.sDataFmt;      break; //- sign
            case "()": oDF.sDataFmt += ";"+(bShowNegAsRed?"[Red]":"")+"("+oDF.sDataFmt+")";  break; // (Accounting style)
            default:   throw e("Negative format of "+sNegFmt+" not implemented");
          }
        }
      }
    }
    oDataFmts.add(oDF);
  }

  private CellStyle registerStyle(String sKey,CellStyle oCS) {
    CellStyle oPrevCS = oStyRegs.put(sKey,oCS);
    if (oPrevCS != null) log("DUPLICATE style:"+sKey);
    log("Registered style "+sKey);
    return oCS;
  }

  /* Priorities:
   * if (sRowFmt != null)
   *    sRowFmt != back-color  we use that with sStyName for data-fmt.
   *    Use sStyName without Data-fmt and apply BG color
   *    Use sStyName for Data-fmt and apply BG color
   * else
   *    Use sStyName without Data-fmt
   *    Use sStyName for Data-fmt
   * end-if
   **/
  private CellStyle chooseStyle(SpecFmt oSF,String sRowFmt) throws Exception {
    //log("chooseStyle "+sRowFmt+" /"+oSF.sName+"/ "+oSF.oDF.sDataFmt+" "+oSF.sData);
    if ((sRowFmt == null) && (oSF.oDF.oM == null) && (oSF.sName == null)) return null;

    String sStyKey = oSF.oDF.sName;

    if ((sRowFmt == null) && (oSF.sName == null)) {
      if (oStyRegs.containsKey(sStyKey)) return oStyRegs.get(sStyKey);
      CellStyle oCS = oWB.createCellStyle();
      oCS.setDataFormat(oWB.createDataFormat().getFormat(oSF.oDF.sDataFmt));
      return registerStyle(sStyKey,oCS);
    }

    StyleDef oRowSD = null;
    StyleDef oSD = null;
    if (oSF.sName != null) {
      oSD = oStyDefs.get(oSF.sName);
      if (oSD != null) {
        sStyKey += "/"+oSF.sName;
      } else {
        log("StyleDef ref "+oSF.sName+" ignored");
      }
    }
    if (sRowFmt != null) {
      oRowSD = oStyDefs.get(sRowFmt);
      if (oRowSD != null) {
        if ((oSD == null) || ((oSD.oSA.oBackRGB == null) && (oRowSD.oSA.bBackOnly))) {
          sStyKey += "/"+sRowFmt;
        } else {
          oRowSD = null;
        }
      } else {
        log("StyleDef row ref "+sRowFmt+" ignored");
      }
    }

    CellStyle oCS = oStyRegs.get(sStyKey);
    if (oCS != null) return oCS;

    oCS = performStyleMerge(oSD,oRowSD,oSF.oDF);
    return registerStyle(sStyKey,oCS);
  }

  private CellStyle performStyleMerge(StyleDef oSD,StyleDef oRowSD,DataFmt oDF) throws Exception {
    CellStyle oCS = oWB.createCellStyle();
    if (oSD == null) {  // If RowSD only SD then treat as primary
      oSD = oRowSD;
      oRowSD = null;
    }
    if (oDF.sDataFmt != null) oCS.setDataFormat(oWB.createDataFormat().getFormat(oDF.sDataFmt));
    Font oFont = null;
    if (oSD != null) {
      StyleAttrs oSA = oSD.oSA;
      if (oSA.bNeedFont) {
        oFont = oWB.createFont();
        if (oSA.sFontFamily != null) oFont.setFontName(oSA.sFontFamily);
        if (oSA.bBold) oFont.setBold(true);
        if (oSA.bItalic) oFont.setItalic(true);
        if (oSA.bStrike) oFont.setStrikeout(true);
        if (oSA.dPoints != 0.00) oFont.setFontHeightInPoints(new Double(oSA.dPoints).shortValue());
        if (oSA.bLink && (oSA.oForeRGB == null)) {
          oFont.setUnderline(Font.U_SINGLE);
          oFont.setColor(IndexedColors.BLUE.getIndex());
        }
        if (oSA.oForeRGB != null) {
          if (oSA.nForeIX != -1) {
            oFont.setColor(oSA.nForeIX);
          } else {
            // seemed only way to make this work with custom colors
            XSSFColor oXC = new XSSFColor(oSA.oForeRGB,null);
            ((XSSFFont)oFont).setColor(oXC);
          }
        }
      }
      if (oSA.oBackRGB != null) {
        IndexedColorMap oICM = ((XSSFWorkbook)oWB).getStylesSource().getIndexedColors();
        XSSFColor oColBG = new XSSFColor(oSA.oBackRGB,oICM);
        ((XSSFCellStyle)oCS).setFillForegroundColor(oColBG);
        ((XSSFCellStyle)oCS).setFillBackgroundColor(oColBG);
        ((XSSFCellStyle)oCS).setFillPattern(FillPatternType.FINE_DOTS);
      }
      switch(oSA.cHorzAlign) {
        case 'L': oCS.setAlignment(HorizontalAlignment.LEFT); break;
        case 'C': oCS.setAlignment(HorizontalAlignment.CENTER); break;
        case 'R': oCS.setAlignment(HorizontalAlignment.RIGHT); break;
      }
      switch(oSA.cHorzAlign) {
        case 'T': oCS.setVerticalAlignment(VerticalAlignment.TOP); break;
        case 'M': oCS.setVerticalAlignment(VerticalAlignment.CENTER); break;
        case 'B': oCS.setVerticalAlignment(VerticalAlignment.BOTTOM); break;
      }
      switch(oSA.cTot) {
        case '-': oCS.setBorderTop(BorderStyle.MEDIUM);      break;
        case '~': oCS.setBorderTop(BorderStyle.DOUBLE);      break;
        case '=': oCS.setBorderTop(BorderStyle.DOUBLE);
                  oCS.setBorderBottom(BorderStyle.DOUBLE);   break;
      }
      if (oFont != null) oCS.setFont(oFont);
      if (oDF.sDataFmt != null) {
        if (oSA.cHorzAlign == 0) oCS.setAlignment(HorizontalAlignment.RIGHT);
        if (oSA.cVertAlign == 0) oCS.setVerticalAlignment(VerticalAlignment.CENTER);
      }
    }

    if ((oRowSD != null) && (oRowSD.oSA.bBackOnly) && (oSD.oSA.oBackRGB == null)) {
      IndexedColorMap oICM = ((XSSFWorkbook)oWB).getStylesSource().getIndexedColors();
      XSSFColor oColBG = new XSSFColor(oRowSD.oSA.oBackRGB,oICM);
      ((XSSFCellStyle)oCS).setFillForegroundColor(oColBG);
      ((XSSFCellStyle)oCS).setFillBackgroundColor(oColBG);
      ((XSSFCellStyle)oCS).setFillPattern(FillPatternType.FINE_DOTS);
    }
    if ((oSD != null) && (oSD.oSA.sCustExit != null)) {
      customExit(oCS,oSD.oSA.sCustExit);
    }
    return oCS;
  }

  private CellStyle getLinkStyle(String sCellData,String sName) throws Exception {
    SpecFmt oSF = parseCellData(sCellData);
    oSF.oDF = getDefaultFmt();
    CellStyle oSty = chooseStyle(oSF,sName);
    return oSty;
  }

  static Matcher oM2f   = Pattern.compile("f").matcher("");
  static Matcher oM2FF  = Pattern.compile("FF[(]([^)]+)[)]").matcher("");
  static Matcher oM2b   = Pattern.compile("b").matcher("");
  static Matcher oM2i   = Pattern.compile("i").matcher("");
  static Matcher oM2s   = Pattern.compile("s").matcher("");
  static Matcher oM2l   = Pattern.compile("l").matcher("");
  static Matcher oM2pt  = Pattern.compile("[0-9]+([.][0-9]+)?").matcher("");
  static Matcher oM2tot = Pattern.compile("=").matcher("");
  static Matcher oM2sub = Pattern.compile("-").matcher("");
  static Matcher oM2fin = Pattern.compile("~").matcher("");
  static Matcher oM2FG  = Pattern.compile("FG[(](([a-zA-Z][a-zA-Z0-9-_]+)|([0-9]+,[0-9]+,[0-9]+))[)]").matcher("");
  static Matcher oM2BG  = Pattern.compile("BG[(](([a-zA-Z][a-zA-Z0-9-_]+)|([0-9]+,[0-9]+,[0-9]+))[)]").matcher("");
  static Matcher oM2C   = Pattern.compile("C").matcher("");
  static Matcher oM2L   = Pattern.compile("L").matcher("");
  static Matcher oM2R   = Pattern.compile("R").matcher("");
  static Matcher oM2T   = Pattern.compile("T").matcher("");
  static Matcher oM2B   = Pattern.compile("B").matcher("");
  static Matcher oM2M   = Pattern.compile("M").matcher("");
  static Matcher oM3rgb = Pattern.compile("([0-9]+),([0-9]+),([0-9]+)").matcher("");
  static Matcher oM2CE  = Pattern.compile("CE[(]([a-zA-Z0-9-_]+)[)]").matcher("");

  /**
    Internal method to parse style attributes into (@link StyleAttrs}.
  */
  private StyleAttrs parseStyleAttrs(StyleDef oSD) throws Exception {
    StyleAttrs oSA = new StyleAttrs();
    String[] sStyStr = new String[]{oSD.sStyStr};
    if (sStyStr[0].startsWith(":")) sStyStr[0] = sStyStr[0].substring(1);
    String s = null;
    // done early to avoid match ambiguity
    if ((s = applyMatcher(sStyStr,oM2FG )) != null) {oSA.bNeedFont = true; setColorRGB(true,oSA,s);}
    if ((s = applyMatcher(sStyStr,oM2BG )) != null) {setColorRGB(false,oSA,s);}
    if ((s = applyMatcher(sStyStr,oM2CE )) != null) {oSA.sCustExit = s.substring(3,s.length()-1);}

    if ((s = applyMatcher(sStyStr,oM2FF )) != null) {oSA.bNeedFont = true; oSA.sFontFamily = s;}
    if ((s = applyMatcher(sStyStr,oM2f  )) != null) {oSA.bNeedFont = true; oSA.sFontFamily = sFntFix;}
    if ((s = applyMatcher(sStyStr,oM2b  )) != null) {oSA.bNeedFont = true; oSA.bBold = true;}
    if ((s = applyMatcher(sStyStr,oM2i  )) != null) {oSA.bNeedFont = true; oSA.bItalic = true;}
    if ((s = applyMatcher(sStyStr,oM2s  )) != null) {oSA.bNeedFont = true; oSA.bStrike = true;}
    if ((s = applyMatcher(sStyStr,oM2l  )) != null) {oSA.bNeedFont = true; oSA.bLink = true;}
    if ((s = applyMatcher(sStyStr,oM2pt )) != null) {oSA.bNeedFont = true; oSA.dPoints = Double.parseDouble(s);}

    if ((s = applyMatcher(sStyStr,oM2L  )) != null)  {oSA.cHorzAlign = s.charAt(0);}
    if ((s = applyMatcher(sStyStr,oM2C  )) != null)  {oSA.cHorzAlign = s.charAt(0);}
    if ((s = applyMatcher(sStyStr,oM2R  )) != null)  {oSA.cHorzAlign = s.charAt(0);}
    if ((s = applyMatcher(sStyStr,oM2T  )) != null)  {oSA.cVertAlign = s.charAt(0);}
    if ((s = applyMatcher(sStyStr,oM2M  )) != null)  {oSA.cVertAlign = s.charAt(0);}
    if ((s = applyMatcher(sStyStr,oM2B  )) != null)  {oSA.cVertAlign = s.charAt(0);}

    if ((s = applyMatcher(sStyStr,oM2tot)) != null)  {oSA.bNeedFont = true; oSA.cTot = s.charAt(0);}
    if ((s = applyMatcher(sStyStr,oM2sub)) != null)  {oSA.bNeedFont = true; oSA.cTot = s.charAt(0);}
    if ((s = applyMatcher(sStyStr,oM2fin)) != null)  {oSA.bNeedFont = true; oSA.cTot = s.charAt(0);}


    if (sStyStr[0].trim().length() != 0) {
      log("Style Attrs "+oSD.sStyStr+" remnants "+sStyStr[0]+" had no effect");
    }
    if (!oSA.bNeedFont && (oSA.cVertAlign == 0) && (oSA.cHorzAlign == 0) && (oSA.oForeRGB == null) && (oSA.oBackRGB != null)) {
      oSA.bBackOnly = true;
    } else {
      if (oSA.cVertAlign == 0) oSA.cVertAlign ='M';
    }
    //if (oSA.cHorzAlign == 0) oSA.cHorzAlign = (oSD.sFmtStr != null)?'R':'L';
    return oSA;
  }

  private void setColorRGB(boolean bFG,StyleAttrs oSA,String sRGB) throws Exception {
    byte[] oRGB = null;
    short nIndex = -1;
    if (oM3rgb.reset(sRGB).find()) {
      oRGB = rgb(Integer.parseInt(oM3rgb.group(1)),Integer.parseInt(oM3rgb.group(2)),Integer.parseInt(oM3rgb.group(3)));
    } else {
      String sStr = sRGB.toUpperCase().replace('-','_');
      IndexedColors oIC = IndexedColors.valueOf(sStr);
      if (oIC != null) {
        nIndex = oIC.getIndex();
        IndexedColorMap oICM = ((XSSFWorkbook)oWB).getStylesSource().getIndexedColors();
        oRGB = oICM.getRGB(oIC.getIndex());
      } else {
        log("No match for color "+sStr);
      }
    }
    if (oRGB != null) {
      if (bFG) {
        oSA.oForeRGB = oRGB;
        oSA.nForeIX = nIndex;
      } else {
        oSA.oBackRGB = oRGB;
        //oSA.nBackIX = nIndex;
      }
    }
  }

  private String applyMatcher(String[] sStyStr,Matcher oM) throws Exception {
    if (oM.reset(sStyStr[0]).find()) {
      String sRet = oM.group();
      if (oM.groupCount() > 1) sRet = oM.group(1);
      sStyStr[0] = oM.reset(sStyStr[0]).replaceFirst("");
      return sRet;
    }
    return null;
  }

  private SpecFmt parseCellData(String sCellData) throws Exception {
    SpecFmt oSF = new SpecFmt();
    oSF.sData = sCellData;
    if ((sCellData != null) && (oFmt.reset(sCellData).find())) {
      oSF.sData = oFmt.group("rest");
      if (oFmt.group("mrg") != null) {
        if (oFmt.group("mrg").startsWith("0")) oSF.bPlain = true;
        if (!oSF.bPlain) { // Leading 0 kills merge, keeps col-count
          oSF.nMerge = Integer.parseInt(oFmt.group("mrg"));
        }
      }
      if ((oFmt.group("fmt") != null) && (oFmt.group("fmt").length() > 0)) {
        oSF.sName = oFmt.group("fmt");
        if (oSF.sName.startsWith(":")) {
          oSF.sAnonFmt = oSF.sName.substring(1);
          if (!oStyDefs.containsKey(oSF.sName)) {
            insertStyleDef(oSF.sName,oSF.sAnonFmt);
          }
        }
      }
    }
    return oSF;
  }

  private SpecFmt setCellContent(String sColFmt,Row oRow,int col,String sData) throws Exception {
    Cell oC = oRow.createCell(col);
    SpecFmt oSF = parseCellData(sData);
    for(DataFmt oDF:oDataFmts) {
      if (oDF.oM == null) {// catch all
        oC.setCellValue(oSF.sData);
        oSF.oDF = oDF;
        break;
      } else {
        if ((oSF.sData != null) && (oDF.oM.reset(oSF.sData).find())) {
          String sPureStr = oSF.sData.replaceAll("[^0-9.-]","");
          //log("Insert "+oSF.sData+" "+sData+" "+sPureStr+" as "+oDF.sDataFmt+" "+oDF.oM);
          if (oDF.bInteger) {
            long m = Long.decode(sPureStr);
            oC.setCellValue(m);
          } else {
            double d = Double.parseDouble(sPureStr);
            oC.setCellValue(d);
          }
          oSF.oDF = oDF;
          break;
        }
      }
    }
    if (oSF.oDF == null) throw e("cannot happen "+sData);

    CellStyle oSty = chooseStyle(oSF,sColFmt);
    if (oSty != null) oC.setCellStyle(oSty);


    return oSF;
  }

  private DataFmt getDefaultFmt() {
    for(DataFmt oDF:oDataFmts) {
      if (oDF.oM == null) {// catch all
        return oDF;
      }
    }
    return null;
  }

  /*protected void copyCell(String sSheet,int nSrcRow,int nSrcCol,int nTrgRow,int nTrgCol) throws Exception {
    Sheet oS = oWB.getSheet(sSheet);
    if (oS == null) return;
    Row oRow = oS.getRow(nSrcRow);
    if (oRow== null) return;
    Cell oC = oRow.getCell(nSrcCol);
    if (oC == null) return;
    oRow = oS.getRow(nTrgRow);
    if (oRow== null) return;
    copyCell(oRow,nTrgCol,oC);
  }*/

  /*private Font cloneFont(Font oF) throws Exception {
    Font oN = oWB.createFont();
    oN.setBold      (oF.getBold      ());
    oN.setCharSet   (oF.getCharSet   ());
    oN.setColor     (oF.getColor     ());
    oN.setFontHeight(oF.getFontHeight());
    oN.setFontName  (oF.getFontName  ());
    oN.setItalic    (oF.getItalic    ());
    oN.setStrikeout (oF.getStrikeout ());
    oN.setTypeOffset(oF.getTypeOffset());
    oN.setUnderline (oF.getUnderline ());
    return oN;
  }*/

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
        log("Formula "+s);
        return s;
      default: throw e("Bad Cell type:"+oType+" Col="+oC.getColumnIndex());
    }
  }

  private String dequote(String sD) {
    sD = sD.trim();
    while(sD.startsWith("\"")) sD = sD.substring(1);
    while(sD.endsWith("\"")) sD = sD.substring(0,sD.length()-1);
    sD = sD.trim();
    if (sD.length() == 0) return null;
    return sD;
  }

  private CellStyle useLinkStyle(String sData,String sLinkSty) throws Exception {
    if (sLinkSty == null) sLinkSty = "#lkc";
    return getLinkStyle(sData,sLinkSty);
  }



  private void clonePrintSetup(Sheet oNewS,Sheet oOldS,String sPrintArea) throws Exception {
    PrintSetup oFromPS = oOldS.getPrintSetup();
    PrintSetup oToPS   = oNewS.getPrintSetup();
    oToPS.setCopies(       oFromPS.getCopies());
    oToPS.setDraft(        oFromPS.getDraft());
    oToPS.setFitHeight(    oFromPS.getFitHeight());
    oToPS.setFitWidth(     oFromPS.getFitWidth());
    oToPS.setFooterMargin( oFromPS.getFooterMargin());
    oToPS.setHeaderMargin( oFromPS.getHeaderMargin());
    oToPS.setHResolution(  oFromPS.getHResolution());
    oToPS.setLandscape(    oFromPS.getLandscape());
    oToPS.setLeftToRight(  oFromPS.getLeftToRight());
    oToPS.setNoColor(      oFromPS.getNoColor());
    oToPS.setNoOrientation(oFromPS.getNoOrientation());
    oToPS.setPageStart(    oFromPS.getPageStart());
    oToPS.setPaperSize(    oFromPS.getPaperSize());
    oToPS.setScale(        oFromPS.getScale());
    oToPS.setUsePage(      oFromPS.getUsePage());
    oToPS.setValidSettings(oFromPS.getValidSettings());
    oToPS.setVResolution(  oFromPS.getVResolution());
    if (sPrintArea != null) {
      if (sPrintArea.contains("!")) {
         sPrintArea = sPrintArea.substring(sPrintArea.indexOf("!") + 1);
      }
      oWB.setPrintArea(oWB.getSheetIndex(oNewS),sPrintArea);
    }
    oNewS.setRepeatingColumns( oOldS.getRepeatingColumns());
    oNewS.setRepeatingRows(    oOldS.getRepeatingRows());

  }

  private void cloneChartObject(Sheet oNewS,Sheet oOldS) throws Exception {
    // turn out to be difficult.  Work around was to copy entire file, remove unwanted part and add other sheets.
  }


  private int copyRow(Sheet oS,int nRow,Row oRow,int nMaxCol) throws Exception {
    Row oNewR = oS.createRow(nRow);
    oNewR.setHeight(oRow.getHeight());
    for(int i=oRow.getFirstCellNum(),iMax = oRow.getLastCellNum(); i <= iMax; i++) {
      if (i < 0) break;
      if (i > nMaxCol) nMaxCol = i;
      Cell oCell = oRow.getCell(i);
      if (oCell != null) {
        copyCell(oNewR,i,oCell);
      }
    }
    return nMaxCol;
  }

  private Cell copyCell(Row oRow,int nCol,Cell oCell) throws Exception {
    Cell oNewCell = oRow.createCell(nCol);
    CellStyle oSty = oCell.getCellStyle();
    if (oSty != null) {
      int nHash = oSty.hashCode();
      CellStyle oNewSty = oStyMap.get(""+nHash);
      if (oNewSty == null) {
        oNewSty = oWB.createCellStyle();
        oNewSty.cloneStyleFrom(oSty);
        oStyMap.put(""+nHash,oNewSty);
      }
      oNewCell.setCellStyle(oNewSty);
    }
    return copyCellValue(oNewCell,oCell);
  }

  private Cell copyCellValue(Cell oNewCell,Cell oC) throws Exception {
    CellType oType = oC.getCellType();
    switch (oType) {
    case BLANK:
        oNewCell.setCellType(CellType.BLANK);
        break;
      case STRING:
        oNewCell.setCellValue(oC.getStringCellValue());
        break;
      case NUMERIC:
        oNewCell.setCellValue(oC.getNumericCellValue());
        break;
      case FORMULA:
        oNewCell.setCellFormula(oC.getCellFormula());
        break;
      case BOOLEAN:
        oNewCell.setCellValue(oC.getBooleanCellValue());
        break;
      case _NONE:
        break;
      case ERROR:
        oNewCell.setCellErrorValue(oC.getErrorCellValue());
        break;
    }
    return oNewCell;
  }

  private byte[] rgb(int r,int g,int b) {
    return new byte[]{(byte)r,(byte)g,(byte)b};
  }

  private void addCellComment(Row oRow,int nCol,String sText,Font oFont) throws Exception {
    Sheet oS = oRow.getSheet();
    Cell oC = oRow.getCell(nCol);
    if (oC == null) oC = oRow.createCell(nCol);
    Drawing<?> oD = oS.createDrawingPatriarch();
    CreationHelper oCH = oS.getWorkbook().getCreationHelper();
    ClientAnchor oA = oCH.createClientAnchor();
    oA.setCol1(oC.getColumnIndex());
    oA.setCol2(oC.getColumnIndex()+6);
    oA.setRow1(oRow.getRowNum());
    oA.setRow2(oRow.getRowNum()+6);
    Comment oComm = oD.createCellComment(oA);
    RichTextString oStr = oCH.createRichTextString(sText);
    if (oFont != null) oStr.applyFont(oFont);
    oComm.setString(oStr);
    oC.setCellComment(oComm);
  }

  /*private*/ void addCellComment(String sSheet,int nRow,int nCol,Cell oCell,Comment oNote) throws Exception {
    Sheet oS = oWB.getSheet(sSheet);
    if (oS == null) return;
    Row oRow = oS.getRow(nRow);
    if (oRow == null) oS.createRow(nRow);
    String sStr = getCellAsStr(oCell);
    Cell oC = null;
    if (sStr.startsWith("*")) { // retain existing
      oC = oRow.getCell(nCol);
      if (oC == null) oC = oRow.createCell(nCol);
    } else {
      oC = copyCell(oRow,nCol,oCell);
    }
    if (oNote != null) {
      Drawing<?> oD = oS.createDrawingPatriarch();
      CreationHelper oCH = oS.getWorkbook().getCreationHelper();
      ClientAnchor oA = oCH.createClientAnchor();
      oA.setCol1(oC.getColumnIndex());
      oA.setCol2(oC.getColumnIndex()+6);
      oA.setRow1(oRow.getRowNum());
      oA.setRow2(oRow.getRowNum()+6);
      Comment oComm = oD.createCellComment(oA);
      oComm.setString(oNote.getString());
      oComm.setAuthor(oNote.getAuthor());
      oC.setCellComment(oComm);
    }
  }


//  ----------- Area processing

  private HdrCol[] parseHeader(Area oA,String sColStr,String sHdrFmt) throws Exception {
    String[] sCols = sColStr.split("/");
    HdrCol[] oHdrs = new HdrCol[sCols.length];
    for(int i=0,iMax=sCols.length; i<iMax; i++) {
      String sCol = sCols[i];
      HdrCol oHC = oHdrs[i] = new HdrCol();
      oHC.nHdrIX = i;
      oHC.sText = sCol;
      //oHC.sRawText = sCol;
      oHC.sHdrFmt = sHdrFmt;
      if (oFmt.reset(sCol).find()) {
        oHC.sText = oFmt.group("rest");
        if (oFmt.group("mrg") != null) {
          if (oFmt.group("mrg").startsWith("0")) oHC.bPlain = true;
          if (!oHC.bPlain) { // Leading 0 kills merge, keeps col-count
            oHC.nMerge = Integer.parseInt(oFmt.group("mrg"));
          }
        }
        if ((oFmt.group("fmt") != null) && (oFmt.group("fmt").length() > 0)) {
          oHC.sHdrFmt = oFmt.group("fmt");
        }
      }
    }
    return oHdrs;
  }

  private void calcDimensions(Area oA) {
    if (oA.oHdrs.size() == 0) return;
    HdrCol[] oHCs = oA.oHdrs.get(oA.oHdrs.size() - 1); // last takes and is assumed to have most cols
    for(int col=0,colMax=oHCs.length; col<colMax; col++) {
      HdrCol oHC = oHCs[col];
      oHC.nMaxStr = Math.min(10,oHC.sText.length());
      oHC.nWidthMult = 320;
      for(int row=0,rowMax=oA.oRows.size(); row<rowMax; row++) {
        String[] sRow = oA.oRows.get(row);
        if (col < sRow.length) {
          if ((sRow[col] != null) && (sRow[col].length() > 0)) {
            String sStr = sRow[col];
            if (sStr.startsWith("{")) sStr = sStr.substring(sStr.lastIndexOf("}")+1);
            if (sStr.length() > oHC.nMaxStr) {
              oHC.nWidthMult = 280;
              oHC.nMaxStr = sRow[col].length();
            }
          }
        }
      }
    }
  }

  private Area writeArea(Area oA,String sSheet) throws Exception {
    calcDimensions(oA);
    Sheet oS = oWB.getSheet(sSheet);
    if (oS == null) oS = oWB.createSheet(sSheet);
    int nRow = oA.nBaseRow;
    HdrCol[] oHCs = null;
    for(int i=0,iMax=oA.oHdrs.size(); i<iMax; i++) {
      oHCs = oA.oHdrs.get(i); // last takes and is assumed to have most cols
      Row oHdr = oS.getRow(nRow);
      if (oHdr == null) oHdr = oS.createRow(nRow);
      nRow += 1;
      boolean bMerge = false;
      boolean bLast = (i == oA.oHdrs.size() - 1);
      int nBias = 0;
      for(HdrCol oHC:oHCs) {
        int nCol = oHC.nHdrIX+oA.nBaseCol+nBias;
        setCellContent(oHC.sHdrFmt,oHdr,nCol,oHC.sText);
        if (oHC.nMerge != 0) {
          bMerge = true;
          nBias += oHC.nMerge - 1;
        }
      }
      nBias = 0;
      for(HdrCol oHC:oHCs) {
        int nCol = oHC.nHdrIX+oA.nBaseCol;
        if (bMerge) {
          if (oHC.nMerge > 0) {
            oS.addMergedRegion(new CellRangeAddress(oHdr.getRowNum(),oHdr.getRowNum(),nCol+nBias,nCol+nBias+oHC.nMerge-1));
            nBias += oHC.nMerge - 1;
          }
        } else {
          if (bLast) oS.setColumnWidth(nCol,oHC.nMaxStr*oHC.nWidthMult);
        }
      }
      if (oHdr.getLastCellNum() > oA.nMaxCol) oA.nMaxCol = oHdr.getLastCellNum();
    }
    if (oA.oHdrs.size() > 0) {
      oS.createFreezePane(0,oA.oHdrs.size()+oA.getBaseRow());
    }

    int nMaxRows = oA.oRows.size();

    for(int row=0,rowMax=nMaxRows; row<rowMax; row++) {
      Row oRow = oS.getRow(nRow);
      if (oRow == null) oRow = oS.createRow(nRow);
      nRow += 1;
      int nBias = 0;
      String[] sRows = oA.oRows.get(row);
      //int nStripe = ((oA.nStripes != null) && (row < oA.nStripes.size()))?oA.nStripes.get(row).intValue():0;
      String sColFmt = ((oA.sColFmts != null) && (row < oA.sColFmts.size()))?oA.sColFmts.get(row):null;
      ArrayList<String> sMerges = new ArrayList<>();
      for(int col=0,colMax=oHCs.length; col<colMax; col++) {
        if (col >= sRows.length) break;
        HdrCol oHC = oHCs[col];
        String sData = sRows[col];
        if (sData == null) sData = "";
        int nCol = oA.nBaseCol+oHC.nHdrIX+nBias;
        SpecFmt oSF = setCellContent(sColFmt,oRow,nCol,sData);
        if (oSF.nMerge > 0) {
          int nMerge = oSF.nMerge;
          for(int j=1,jMax=nMerge; j<jMax; j++) {
            setCellContent(null,oRow,nCol+j,"");
          }
          sMerges.add(0,""+(nCol)+","+nMerge);
          //log("Merge store "+""+(col+nBias)+","+sMerge);
          nBias += nMerge - 1;
          //oS.addMergedRegion(new CellRangeAddress(oRow.getRowNum(),oRow.getRowNum(),col,col+nMerge));
        }
        if (oRow.getLastCellNum() > oA.nMaxCol) oA.nMaxCol = oRow.getLastCellNum();
      }

      if (sMerges.size() > 0) {
        for(String sMerge:sMerges) {
          String[] sParts = sMerge.split(",");
          int nCol = Integer.parseInt(sParts[0]);
          int nCols = Integer.parseInt(sParts[1]);
          //log("Merge "+nCol+" "+nCols);
          oS.addMergedRegion(new CellRangeAddress(oRow.getRowNum(),oRow.getRowNum(),nCol,nCol+nCols-1));
        }
      }
    }
    return oA;
  }

  private String cellAsString(Cell oC) throws Exception {
    StringBuilder oSB = new StringBuilder();
    oSB.append("Cell(");
    oSB.append(oC.getAddress());
    oSB.append(") ");
    oSB.append(getCellType(oC));
    CellStyle oCS = oC.getCellStyle();
    if (oCS != null) {
      int nFix = oCS.getFontIndexAsInt();
      if (nFix > 0) {
        String sFixKey = String.format("%08d",nFix);
        if (!oFontCache.containsKey(sFixKey)) {
          Font oFont = oWB.getFontAt(nFix);
          oFontCache.put(sFixKey,oFont);
        }
      }
      String sDataFmt = oCS.getDataFormatString();
      if (sDataFmt == null) sDataFmt = "";
      int nCix = oCS.getIndex();
      if (nCix > 0) {
        String sCixKey = String.format("%08d",nCix);
        if (!oStyleCache.containsKey(sCixKey)) {
          oStyleCache.put(sCixKey,oCS);
        }
      }
      oSB.append(String.format("%-48s",String.format("Sty(cix=%d,fix=%d,df=%s)",nCix,nFix,sDataFmt)));
    } else {
      oSB.append(String.format("%-48s","--no style--"));
    }
    oSB.append(" "+getCellAsStr(oC));
    return ""+oSB;
  }

  private String getCellType(Cell oC) {
    if (oC == null) return null;
    CellType oType = oC.getCellType(); // javadoc is wrong here
    //if (bDebug) log("Cell Type "+nType+" row "+/*oRow.getRowNum()* /"?"+" col "+nCol);
    switch (oType) {
      case BLANK:   return "blnk";
      case STRING:  return "str";
      case NUMERIC: return "num";
      case FORMULA: return "expr";
      default:      return "?";
    }
  }


}