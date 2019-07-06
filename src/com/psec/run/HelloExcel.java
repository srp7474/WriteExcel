// HelloExcel - Most Basic example of using WriteExcel

/* This is also used to demonstrate the WriteExcel capabilities.
 *
  History:
    EC9519 - original
*/

/*
 @license
 Copyright (c) 2019 by Steve Pritchard of Rexcel Systems Inc.
 This file is made available under the terms of the Creative Commons Attribution-ShareAlike 3.0 license
 http://creativecommons.org/licenses/by-sa/3.0/.
 Contact: public.pritchard@gmail.com
*/
package com.psec.run;
import  com.psec.excel.WriteExcel;

/**
Demonstrate a very basic example of WriteExcel. The code uses Hungarian notation and is copied into the
{@link <a href={@docRoot}overview-summary.html>Javadoc overview</a>}.

*/

public class HelloExcel extends WriteExcel {


  public static void log(String sMsg) {System.out.println(sMsg);}
  public static Exception e(String s) {return new Exception(s); }

  // ----------------------- Globals ----------------------
  WriteExcel oWE;

  // ---------------------- Mainline ----------------------
  public void run(String[] sArgs) throws Exception {
    log("HelloExcel writing "+sArgs[0]);
    oWE = WriteExcel.create(this,sArgs[0]);
    oWE.setNegativeFormat(true,"");

    writeSalesSheet();

    oWE.close();
    oWE = null;
    log("HelloExcel completed writing "+sArgs[0]);
  }

  private void writeSalesSheet() throws Exception {

    String sSheet = "sample-sales";
    Area oA = oWE.createArea(sSheet,1,1)
      .header("{4.#title}Sample Sales Report")
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
    oA.addRow(String.format("{:Rb}TOTAL/%d/%.2f/%.2f",nTotSales,dTotRev / nTotSales,dTotRev).split("/"),"#TOT");
    oA.writeArea().colWidth(-1,3).addDataFilterLine();
  }

  public static void main(String[] sArgs) {
    try {
      HelloExcel oHE = new HelloExcel();
      oHE.run(sArgs);
    } catch (Exception e) {
      log("HelloExcel Croaked:"+e);
    }
  }
}

