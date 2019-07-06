// TestArgsPowerShell - Stub to report powershell arguments

/* This is also used to demonstrate the WriteExcel capabilities.
 *
  History:
    EC9510 - original
*/

/*
 @license
 Copyright (c) 2019 by Steve Pritchard of Rexcel Systems Inc.
 This file is made available under the terms of the Creative Commons Attribution-ShareAlike 3.0 license
 http://creativecommons.org/licenses/by-sa/3.0/.
 Contact: public.pritchard@gmail.com
*/

package com.psec.run;
import  com.psec.util.Parm;
import  com.psec.excel.WriteExcel;
import  com.psec.excel.ReadExcelFile;
import  com.psec.excel.ReadExcelRecord;
import  com.psec.util.OptionalFlag;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.regex.Pattern;
import java.util.regex.Matcher;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.BorderStyle;

/**
  Class used to test and demonstrate the
  {@link com.psec.excel.ReadExcelFile} and
  {@link com.psec.excel.WriteExcel} capabilities.
  <p>

  The <code>-what</code> parameter causes DemoExcel to run a particular test in its suite of commands.  The tests are:
  <ul>
  <li><code>basic</code> - basic test</li>
  <li><code>hello</code> - run HelloExcel example</li>
  <li><code>cloner</code> - demonstrate how to clone a Worksheet</li>
  <li><code>reader</code> - demonstrate the clone functions of WriteExcel</li>
  <li><code>chart</code> - demonstrate how to create a Workbook with a chart</li>
  <li><code>regress</code>- run a full regression test</li>
  </ul>
  <p>
  While this class is designed to also be used by the PSEC Powershell utility, it can also be run with command line interface commands provided in
  this package as batch files of the same name as the <code>-what</code> parameter.

*/

public class DemoExcel implements Parm.MainRun {


  private static void log(String sMsg) {System.out.println(sMsg);}
  //private static Exception e(String s) {return new Exception(s); }

  private static class Brief extends Parm.Brief {
    String  sWhat;
    @OptionalFlag
    String[] sMsgs;
  }

  // ----------------------- Globals ----------------------
  static DemoExcel oSelf;
  Opt              oOpt;
  Brief            oBrief = new Brief();
  WriteExcelReport oWER;
  ReadExcelFile    oREF;

  // ---------------------- Mainline ----------------------
  public void run(Parm.Opt oBaseOpt) throws Exception {
    oSelf = this;
    this.oOpt = (Opt)oBaseOpt;
    log("DemoExcel v1.0 starting,"
      +"\r\n   what= "+oOpt.sWhat
      +"\r\n   out="+oOpt.sOut);
    this.oOpt.setBrief(oBrief);
    if (oOpt.sStr1 == null);
    if (oOpt.sInp == null);
    if (oBrief.sWhat == null);

    boolean bGood = false;

    switch(oOpt.sWhat) {
      case "basic":  writeBasicSheets();  bGood = true; break;
      case "hello":  writeHello();        bGood = true; break;
      case "cloner": writeClonerOutput(); bGood = true; break;
      case "chart":  writeClonedChart();  bGood = true; break;
      case "reader": demoReader();        bGood = true; break;
      case "regress":runRegressTest();    bGood = true; break;
      //default: throw e("What request"+oOpt.sWhat+" not implemented");
    }

    if (bGood) {
      oBrief.sWhat = oOpt.sWhat;
      oBrief.sCondCode = "GOOD";
    } else {
      oBrief.sReason = "What request '"+oOpt.sWhat+"' not implemented";
      oBrief.sCondCode = "FAIL";
    }
  }


  private void writeHello() throws Exception {
    HelloExcel.main(new String[]{oOpt.sOut});
  }

  private void runRegressTest() throws Exception {
    oWER = new WriteExcelReport();
    oWER.begin(oOpt.sOut);
    oWER.addDataFormat("numx","^@@[-]?[0-9]+[.][0-9]+$","00.000;[Blue]-00.000");

    oWER.bookSheet("index");

    writeFormatsSheet();
    writeSalesSheet();
    writeLinksSheet();
    writeMiscSheet();

    oREF = new ReadExcelFile();
    oREF.openFile(oOpt.sInp);
    oWER.addExternalSheet("cloned",oREF.getSheet("sample-sales"),null);

    writeIndexSheet(); // must be done last as links are bi-directional

    oWER.end();

    readBackAndLogValues();
  }


  /** This exit is used in creating a chart from a template.
     <p>
     It exercises some functions not covered elsewhere and is thus
     part of the regression test.
  */

  public static class ReadCounts extends ReadExcelRecord {
    public String[] getColMap() {
      return "1=sMon;2=nMtg1;3=nMtg2".split("/");
    }
    public String sMon;
    public int nMtg1;
    public int nMtg2;
    @Override
    public boolean canAccept() throws Exception {
      if (this.sMon == null) return false;
      if (this.sMon.length() == 3) return true;
      return false;
    }
  }

  /* This method demonstrates cloning a chart
    and changing the data values.
    The refresh and refreshAll calls cause the formulas to be updated.

    Note that the getStrValue also triggers the cells to be updated in some cases.
  */

  private void writeClonedChart() throws Exception {
    oWER = new WriteExcelReport();
    oWER.chartCopy(oOpt.sOut,oOpt.sInp);
    Workbook oWB = oWER.getWorkbook();
    oWB.removeSheetAt(oWB.getSheetIndex("formats"));
    oWB.removeSheetAt(oWB.getSheetIndex("sample-sales"));
    oWB.removeSheetAt(oWB.getSheetIndex("sample-bal-template"));

    String sSheet = "index";
    WriteExcel.Area oA = oWER.createArea(sSheet,1,1)
      .header("/{#title}Workbook Index")
      .header("")
      .header("Link/Description of Sheet","#hdrBlue");

    ReadExcelFile oREF = new ReadExcelFile();
    oREF.attachWorkbook(oWB);
    ReadExcelRecord[] oRows = oREF.readExcelSmart("chart",ReadCounts.class,1,true);
    for(ReadExcelRecord oRER:oRows) {
      ReadCounts oRC = (ReadCounts)oRER;
      log(String.format("%2d %s %5d %5d",oRC.oRow.getRowNum(),oRC.sMon,oRC.nMtg1,oRC.nMtg2));
      oWER.zapCell("chart",oRC.oRow.getRowNum(),2,String.format("%d",(int)(oRC.nMtg1 * 0.90)));
      oWER.zapCell("chart",oRC.oRow.getRowNum(),3,String.format("%d",(int)(oRC.nMtg2 * 1.15)));
    }
    oREF.closeFile();
    log("Before refresh (4,4)="+getChartValue(4,4)+" (5,4)="+getChartValue(5,4));
    oWER.refreshCell("chart",4,4);
    log("after refresh (4,4)="+getChartValue(4,4)+" (5,4)="+getChartValue(5,4));
    oWER.refreshCells();
    log("after refreshAll (4,4)="+getChartValue(4,4)+" (5,4)="+getChartValue(5,4));


    oA.addRow(new String[]{"chart","Results for exercise of chart routines"});

    oA.writeArea().colWidth(-1,3).addDataFilterLine();

    oWER.makeIndexLink(null,"chart","chart",4,1,null,1,0);

    oWER.end();
  }

  private String getChartValue(int nRow,int nCol) throws Exception {
    return oWER.getStrValue("chart",nRow,nCol);
  }


  // ----------------- Writer routines --------------------
  private static class WriteExcelReport extends WriteExcel {
    WriteExcel oEW;
    public void begin(String sFileName) throws Exception {
      oEW = WriteExcel.create(this,sFileName);
      if ((oSelf.oOpt.bRed || oSelf.oOpt.sNegFmt != null)) {
        String sFmt = null;
        if ("paren".equals(oSelf.oOpt.sNegFmt))  sFmt = "()";
        if ("sign".equals(oSelf.oOpt.sNegFmt))   sFmt = "-";
        if ("color".equals(oSelf.oOpt.sNegFmt))  sFmt = "";
        oEW.setNegativeFormat(oSelf.oOpt.bRed,sFmt);
      }
    }
    public void chartCopy(String sFileName,String sTemplate) throws Exception {
      oEW = WriteExcel.create(this,sFileName,sTemplate);
    }
    public void end() throws Exception {
      oEW.close();
      oEW = null;
    }

    @Override
    public void customExit(CellStyle oCS,String sStr) {
      log("customExit "+sStr+" "+oCS.getIndex());
      switch(sStr) {
        case "s1":
          oCS.setBorderTop(BorderStyle.MEDIUM_DASH_DOT_DOT);
          break;
        case "str2":
          oCS.setBorderBottom(BorderStyle.THICK);
          oCS.setRotation((short)90);
          break;
      }
    }
  }

  private void writeLinksSheet() throws Exception {

    String sSheet = "links";
    WriteExcel.Area oA = oWER.createArea(sSheet,1,1)
      .header("Index/{2.#title}Sample Links")
      .header("")
      .header("Link Type/Link/Description of Usage","#hdrBlue");

    oA.addRow(new String[]{"Specific","specific-link","links to a specific sheet location"});
    oA.addRow(new String[]{"File","file-link","links to an external file"});
    oA.addRow(new String[]{"URL","url-link","links to an external URL"});
    oA.addRow(new String[]{"STD","std-link","links from std method"});
    oA.addRow(new String[]{"STD","std-link","links from std method - commutative"});
    oA.addRow(new String[]{"UNI","uni-link","links from uni method"});
    oA.addRow(new String[]{"UNI","uni-link","links from uni method multiple lines"});

    oA.writeArea().colWidth(-1,3).colWidth(0,20).addDataFilterLine();

    sSheet = "targets";
    WriteExcel.Area oT = oWER.createArea(sSheet,1,1)
      .header("Index/{2.#title}Target Links")
      .header("")
      .header("Link Type/Link/Description of Usage","#hdrBlue");

    oT.addRow(new String[]{"Specific","specific-link","links to a specific sheet location"});
    oT.addRow(new String[]{"STD","std-link","links from std method"});
    oT.addRow(new String[]{"STD","std-link","links from std method - commutative"});
    oT.addRow(new String[]{"UNI","uni-link","links from links.uni"});
    oT.addRow(new String[]{"UNI1","uni-link","links from links.uni"});
    oT.addRow(new String[]{"UNI2","uni-link","links from links.uni"});

    oT.writeArea().colWidth(-1,3).addDataFilterLine();

    String sFileName = oOpt.sInp.replace('\\','/');
    sFileName = sFileName.substring(0,sFileName.lastIndexOf("/")+1)+"sample-text-doc.txt";
    oWER.makeFileLink("#lnk","links",oA.getDataRow()+1,oA.getBaseCol(),"File Link",sFileName);
    oWER.makeUrlLink("#lnk","links",oA.getDataRow()+2,oA.getBaseCol(),"POI Case studies","https://poi.apache.org/casestudies.html");
    oWER.makeStdLink(null,"links",oA.getDataRow()+3,oA.getBaseCol(),null,"targets",oT.getDataRow()+1,oT.getBaseCol());
    oWER.makeStdLink(null,"targets",oT.getDataRow()+2,oT.getBaseCol(),null,"links",oA.getDataRow()+4,oA.getBaseCol());
    oWER.makeUniLink("#lnk","targets",oT.getDataRow()+3,oT.getBaseCol(),"links",oA.getDataRow()+5,oA.getBaseCol(),"uni-link");
    oWER.makeUniLink(null,"targets",oT.getDataRow()+4,oT.getBaseCol(),"links",oA.getDataRow()+6,oA.getBaseCol(),"uni-multi",2);
  }

  private void writeMiscSheet() throws Exception {

    String sSheet = "misc";
    WriteExcel.Area oA = oWER.createArea(sSheet,1,1)
      .header("Index/{#title}Miscellaneous Cells")
      .header("")
      .header("Cell/Description of Test","#hdrBlue");

    oA.addRow(new String[]{"{:i}Zap-Cell","Cell text is zapped with text,format lost"});
    oA.addRow(new String[]{"{:i}Zap-Cell","Cell text is zapped with a number, format lost"});
    oA.addRow(new String[]{"{:i}Zap-Cell","Cell text is zapped, format retained"});
    oA.addRow(new String[]{"{:i}Zap-Cell","Cell text is zapped, format changed"});
    oA.addRow(new String[]{"comment","This cell has a fixed font comment"});
    oA.addRow(new String[]{"comment","This cell has a default font comment"});

    oA.writeArea().colWidth(-1,3).colWidth(0,10).addDataFilterLine();

    oWER.zapCell("misc",oA.getDataRow()+0,oA.getBaseCol(),"text-zap");
    oWER.zapCell("misc",oA.getDataRow()+1,oA.getBaseCol(),"11.1");
    oWER.zapCell("misc",oA.getDataRow()+2,oA.getBaseCol(),"new-text",true);
    oWER.zapCell("misc",oA.getDataRow()+3,oA.getBaseCol(),"{:FG(red)}10.12");

    oWER.addCellComment("misc",oA.getDataRow()+4,oA.getBaseCol(),"Fixed font cell comment\r\nwith 2 lines");
    oWER.addCellComment("misc",oA.getDataRow()+5,oA.getBaseCol(),"Standard font cell comment\r\nwith 3 lines\r\n3rd row",false);

    // Create log entries for read values
    log("---- regression read values -----");
    log(String.format("method=%-12s val=%d","getRowCount",oWER.getRowCount("misc")));
    log(String.format("method=%-12s val=%d","getAbsRow",oA.getAbsRow()));
    log(String.format("method=%-12s val=%d","getHdrCount",oA.getHdrCount()));
    log(String.format("method=%-12s val=%d","getRows",oA.getRows().size()));

  }

  private void writeIndexSheet() throws Exception {

    String sSheet = "index";
    WriteExcel.Area oA = oWER.createArea(sSheet,1,1)
      .header("/{#title}Workbook Index")
      .header("")
      .header("Link/Description of Sheet","#hdrBlue");

    oA.addRow(new String[]{"formats","Results for exercise of formats routines"});
    oA.addRow(new String[]{"sample-sales","Results for exercise of sample-sales creation"});
    oA.addRow(new String[]{"links","Results for exercise of links creation routines"});
    oA.addRow(new String[]{"targets","Results for exercise of links(targets) creation routines"});
    oA.addRow(new String[]{"cloned","Results for exercise of clone routines"});
    oA.addRow(new String[]{"misc","Results for exercise of miscellaneous methods"});

    oA.writeArea().colWidth(-1,3).addDataFilterLine();

    oWER.makeIndexLink(null,"formats","formats",4,1,null);
    oWER.makeIndexLink("#lkc","sample-sales","sales",5,1,null);
    oWER.makeIndexLink("#lnk","links","links",6,1,"#lkc");
    oWER.makeIndexLink(null,"targets","targets",7,1,"#lnk");
    oWER.makeIndexLink("#lnk","cloned","cloned",8,1,null);
    oWER.makeIndexLink(null,"misc","misc",9,1,null);
  }

  private void writeBasicSheets() throws Exception {
    oWER = new WriteExcelReport();
    oWER.begin(oOpt.sOut);

    oWER.addDataFormat("numx","^@@[-]?[0-9]+[.][0-9]+$","00.000;[Blue]-00.000");

    writeFormatsSheet();
    writeSalesSheet();

    oWER.end();

  }

  private void writeFormatsSheet() throws Exception {
    String sSheet = "formats";
    WriteExcel.Area oA = oWER.createArea(sSheet,1,1)
      .header("Index/{8.#title}Formatting Results")
      .header("")
      .header("{2}Built-ins","#hdrBlue")
      .header("Format/Result","#hdrBlue");

    WriteExcel.Area oB = oWER.createArea(sSheet,1,4)
      .header("")
      .header("")
      .header("{2}Built-ins Supplemented","#hdrBlue")
      .header("Format/Result","#hdrBlue");

    WriteExcel.Area oC = oWER.createArea(sSheet,1,7)
      .header("")
      .header("")
      .header("{3}Specific Built-ins","#hdrBlue")
      .header("Name/Format/Result","#hdrBlue");

    String[] sFmts = "str:positive nums/nm1:11.1/num:10.10/nm3:0.056/nm4:12.1234/int:456/:/str:negative nums/nm1:-2.1/num:-3.01/nm3:-0.750/nm4:-4.0000/int:-77".split("/");
    oA.addRow(new String[0]);
    oWER.oEW.addStyleDefn("hdrSect","bC BG(lavender)");

    oA.addRow(new String[]{"{2.hdrSect} No Formatting"});
    for(String sFmt:sFmts) {
      String[] sParts = sFmt.split(":",-1);
      ArrayList<String> oRow = new ArrayList<>();
      oRow.add(sParts[0]);
      oRow.add(sParts[1]);
      oA.addRow(oRow);
    }

    oA.addRow(new String[0]);
    oA.addRow(new String[]{"{2.hdrSect} No Format with Anon BG Color"});
    for(String sFmt:sFmts) {
      String[] sParts = sFmt.split(":",-1);
      ArrayList<String> oRow = new ArrayList<>();
      oRow.add(sParts[0]);
      oRow.add("{:BG(0,248,0)}"+sParts[1]);
      oA.addRow(oRow);
    }

    oA.addRow(new String[0]);
    oA.addRow(new String[]{"{2.hdrSect} No Format with BG Color"});
    for(String sFmt:sFmts) {
      String[] sParts = sFmt.split(":",-1);
      ArrayList<String> oRow = new ArrayList<>();
      oRow.add(sParts[0]);
      oRow.add("{#TOT}"+sParts[1]);
      oA.addRow(oRow);
    }

    oA.addRow(new String[0]);
    oA.addRow(new String[]{"{2.hdrSect} No Fmt Special Numbers"});
    for(String sFmt:"pos:@@10.00/neg:@@-45.10/neg:@@-50.00/pos:@@1234.1234".split("/")) {
      String[] sParts = sFmt.split(":",-1);
      ArrayList<String> oRow = new ArrayList<>();
      oRow.add(sParts[0]);
      oRow.add(sParts[1]);
      oA.addRow(oRow);
    }

    oA.addRow(new String[0]);
    oA.addRow(new String[]{"{2.hdrSect} Row Fmt"});
    for(String sFmt:sFmts) {
      String[] sParts = sFmt.split(":",-1);
      ArrayList<String> oRow = new ArrayList<>();
      oRow.add(sParts[0]);
      oRow.add(sParts[1]);
      oA.addRow(oRow,"#TOT");
    }

    for(String sStyStr:oWER.oEW.getStdExtras()) {
      String[] sParts = sStyStr.split(":",-1);
      String sStyName = sParts[0];
      String sStyAttr = sParts[1];
      oA.addRow(new String[0]);
      oA.addRow(new String[]{String.format("{2.hdrSect}%s : %s",sStyName,sStyAttr)});
      for(String sFmt:sFmts) {
        sParts = sFmt.split(":",-1);
        ArrayList<String> oRow = new ArrayList<>();
        oRow.add(sParts[0]);
        oRow.add("{"+sStyName+"}"+sParts[1]);
        oA.addRow(oRow);
      }
    }

    addSupplementFormats(oB,sFmts);
    addSpecificFormats(oC,sFmts);
    oB.writeArea().colWidth(-1,3);
    oC.writeArea().colWidth(-1,3);
    oA.writeArea().colWidth(-1,3).addDataFilterLine();
  }

  private void addSupplementFormats(WriteExcel.Area oB,String[] sFmts) throws Exception {
    String[] sStys = ("bold:b/italic:i/strike:s/fixed:f/tiny:5T/small:7.2M/green:FG(green)/gray:FG(192,192,192)/grey:FG(grey-25-percent)"
                     +"/mixed:biC/tot:~/sub:-/fin:=/lnk:l/lkc:lC"
                     +"/cust1:b CE(s1)/cust2:CE(str2)/sub:-/fin:=/lnk:l/lkc:lC"
                     ).split("/");
    for(String sStyStr:sStys) { // create styles
      String[] sParts = sStyStr.split(":",-1);
      oWER.oEW.addStyleDefn(sParts[0],sParts[1]);
    }

    for(String sStyStr:sStys) { // use styles
      String[] sParts = sStyStr.split(":",-1);
      String sStyName = sParts[0];
      String sStyAttr = sParts[1];
      oB.addRow(new String[0]);
      oB.addRow(new String[]{String.format("{2.hdrSect}%s : %s",sStyName,sStyAttr)});

      for(String sFmt:sFmts) {
        sParts = sFmt.split(":",-1);
        ArrayList<String> oRow = new ArrayList<>();
        oRow.add(sParts[0]);
        oRow.add("{"+sStyName+"}"+sParts[1]);
        oB.addRow(oRow);
      }
    }
  }

  private void addSpecificFormats(WriteExcel.Area oC,String[] sFmts) throws Exception {
    for(String sStyStr:oWER.oEW.getHdrExtras()) {
      String[] sParts = sStyStr.split(":",-1);
      String sStyName = sParts[0];
      String sStyFmt  = sParts[1];
      ArrayList<String> oRow = new ArrayList<>();
      oRow.add(sStyName);
      oRow.add(sStyFmt);
      oRow.add("{"+sStyName+"}formatted string");
      oC.addRow(oRow);
    }
  }


  private void writeSalesSheet() throws Exception {

    oWER.oEW.addStyleDefn("#TOTx","BG(0,192,0)");

    String sSheet = "sample-sales";
    WriteExcel.Area oA = oWER.createArea(sSheet,1,1)
      .header("Index/{3.#title}Sample Sales Report")
      .header("")
      .header("Month/Unit Sales/Avg. Price/Revenue","#hdrBlue");

    String[] sMonths = "January/February/March/April/May/June/July/August/September/October/November/December".split("/");
    double[] dPrices = new double[]{10.01,11.02,15.03 ,9.04,10.05,17.06 ,22.07,23.08,14.09 ,12.10,13.11,18.12};
    int[]    nSales  = new int[]{15,61,88 ,23,-3,54 ,67,53,21 ,13,23,33};
    int[]    nQtr    = new int[]{0,0,1 ,0,0,2 ,0,0,3 ,0,0,4};

    int    nQtrSales = 0;
    double dQtrRev   = 0.0;
    int    nTotSales = 0;
    double dTotRev   = 0.0;
    for(int i=0,iMax=12; i<iMax; i++) {
      oA.addRow(String.format("{:R}%s/%d/%.2f/%.2f",sMonths[i],nSales[i],dPrices[i],dPrices[i] * nSales[i]).split("/"),i+1);
      nQtrSales += nSales[i];
      dQtrRev   += dPrices[i] * nSales[i];
      nTotSales += nSales[i];
      dTotRev   += dPrices[i] * nSales[i];
      if (nQtr[i] != 0) {
        oA.addRow(String.format("{:Rb}Q%d/%d/%.2f/%.2f",nQtr[i],nQtrSales,dQtrRev / nQtrSales,dQtrRev).split("/"),"#qtr");
        nQtrSales    = 0;
        dQtrRev      = 0.0;
      }
    }
    oA.addRow(new String[0]);
    oA.addRow(String.format("{:Rb}TOTAL/%d/%.2f/%.2f",nTotSales,dTotRev / nTotSales,dTotRev).split("/"),"#TOTx");
    oA.writeArea().colWidth(-1,3).addDataFilterLine();
  }

  private void writeClonerOutput() throws Exception {
    oREF = new ReadExcelFile();
    oREF.openFile(oOpt.sInp);
    oWER = new WriteExcelReport();
    oWER.begin(oOpt.sOut);
    oWER.addExternalSheet("cloned",oREF.getSheet("sample-sales"),null);

    //oWER.addDataFormat("numx","^@@[-]?[0-9]+[.][0-9]+$","00.000;[Blue]-00.000");

    //writeFormatsSheet();
    //writeSalesSheet();

    oWER.end();

  }

  /**
    Used to process selected records. Use of Reflection requires this be made public.
  */
  public static class OpenBalRec extends ReadExcelRecord {
    static Matcher oM = Pattern.compile("^[ALS][0-9]{3}$").matcher("");
    public String[] getColMap() {
      return "3=sLab;2=sName;4=dVal".split("/");
    }
    public String sLab;
    public String sName;
    public double dVal;

    @Override
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

  private void readBackAndLogValues() throws Exception {
    oREF = new ReadExcelFile();
    oREF.openFile(oOpt.sOut);
    Workbook oWB =  oREF.getWorkbook();
    WriteExcel oWE = WriteExcel.create(oWB);

    for(Iterator<Sheet> oI = oWB.sheetIterator(); oI.hasNext();) {
      Sheet oS = oI.next();
      log(String.format("----- sheet:%-14s index:%2d NumRows:%3d FirstRow:%2d LastRow:%3d",
        oS.getSheetName(),oWB.getSheetIndex(oS.getSheetName()),oS.getPhysicalNumberOfRows(),oS.getFirstRowNum(),oS.getLastRowNum()));
      for(int i=oS.getFirstRowNum(),iMax=oS.getLastRowNum(); i<=iMax; i++) {
        Row oRow = oS.getRow(i);
        if (oRow != null) {
          log(String.format("  row:%3d NumCell:%3d FirstCell:%2d LastCell:%3d",
            i,oRow.getPhysicalNumberOfCells(),oRow.getFirstCellNum(),oRow.getLastCellNum()));
          if (oRow.getPhysicalNumberOfCells() > 0) {
            for(int j=oRow.getFirstCellNum(),jMax=oRow.getLastCellNum(); j<=jMax; j++) {
              Cell oC = oRow.getCell(j);
              if (oC != null) {
                log("    "+oWE.cellSummary(oC));
              }
            }
          }
        }
      }
    }
    log("----------- CellStyle Cache --------------");
    for(String s:oWE.dumpCellStyleCache()) {
      log("  "+s);
    }
    log("----------- Font Cache --------------");
    for(String s:oWE.dumpCellFontCache()) {
      log("  "+s);
    }
  }


/*************************************************************************
**************************************************************************
**********                START-UP ROUTINES                     **********
**************************************************************************
**************************************************************************/

  private static class Opt extends Parm.Opt {
    public String  sStr1;
    public String  sOut;
    public String  sInp;
    public String  sWhat;
    public String  sNegFmt;
    public boolean bRed;
  }

  static Opt oOPT = new Opt();

  static Parm[] oParms = {
     new Parm(new Parm.PS(){public void set(String sN) {oOPT.bHelp=true; }},     ".","-h",      null,      "Generate this Help Information that you are now reading")
    ,new Parm(new Parm.PS(){public void set(String sN) {oOPT.sBrief=sN;}},       "m","-brief",  "brief",   "path for Brief output")
    ,new Parm(new Parm.PS(){public void set(String sN) {oOPT.sOut=sN;}},         "m","-out",    "file",    "Target output name")
    ,new Parm(new Parm.PS(){public void set(String sN) {oOPT.sInp=sN;}},         ".","-inp",    "file",    "input template file")
    ,new Parm(new Parm.PS(){public void set(String sN) {oOPT.sWhat=sN;}},        "m","-what",   "string",  "What we are to run")
    ,new Parm(new Parm.PS(){public void set(String sN) {oOPT.bRed=true; }},      ".","-red",      null,    "Show negative numbers in red")
    ,new Parm(new Parm.PS(){public void set(String sN) {oOPT.sNegFmt=sN; }},     ".","-negfmt",  "string", "Negative format")
    ,new Parm(new Parm.PS(){public void set(String sN) {oOPT.addProp(sN);}},     ".","-def",    "key=val", "Defined Property string")
  };

  public static void main(String[] sArgs) {
    oOPT.mainStart(oParms,sArgs,new DemoExcel());
  }
}

