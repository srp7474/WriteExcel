// HelloExcelNHN - Most Basic example of using WriteExcel (No Hungarian notation)

/* This is also used to demonstrate the WriteExcel capabilities.
 *
  History:
    EC9530 - original
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
Demonstrate a very basic example of WriteExcel. The code does not use Hungarian notation and is copied into the
{@link <a href={@docRoot}overview-summary.html>Javadoc overview</a>}.
*/
public class HelloExcelNHN extends WriteExcel {


  public static void log(String msg) {System.out.println(msg);}
  public static Exception e(String s) {return new Exception(s); }

  // ----------------------- Globals ----------------------
  WriteExcel wrtExcel;

  // ---------------------- Mainline ----------------------
  public void run(String[] args) throws Exception {
    log("HelloExcel writing "+args[0]);
    wrtExcel = WriteExcel.create(this,args[0]);
    wrtExcel.setNegativeFormat(true,"");

    writeSalesSheet();

    wrtExcel.close();
    wrtExcel = null;
    log("HelloExcel completed writing "+args[0]);
  }

  private void writeSalesSheet() throws Exception {

    String sheet = "sample-sales";
    Area area = wrtExcel.createArea(sheet,1,1)
      .header("{4.#title}Sample Sales Report")
      .header("")
      .header("Month/Unit Sales/Avg. Price/Revenue","#hdrBlue");

    String[] months = "January/February/March/April/May/June/July/August/September/October/November/December".split("/");
    double[] prices = new double[]{10.01,11.02,15.03 ,9.04,10.05,17.06 ,22.07,23.08,14.09 ,12.10,13.11,18.12};
    int[]    sales  = new int[]{15,61,88 ,23,-3,54 ,67,53,21 ,13,23,33};
    int[]    qtr    = new int[]{0,0,1 ,0,0,2 ,0,0,3 ,0,0,4};

    int    qtrSales = 0;
    double qtrRev   = 0.0;
    int    totSales = 0;
    double totRev   = 0.0;
    for(int i=0,iMax=12; i<iMax; i++) {
      area.addRow(String.format("{:R}%s/%d/%.2f/%.2f",months[i],sales[i],prices[i],prices[i] * sales[i]).split("/"),i+1);
      qtrSales += sales[i];
      qtrRev   += prices[i] * sales[i];
      totSales += sales[i];
      totRev   += prices[i] * sales[i];
      if (qtr[i] != 0) {
        area.addRow(String.format("{:Rb}Q%d/%d/%.2f/%.2f",qtr[i],qtrSales,qtrRev / qtrSales,qtrRev).split("/"),"#qtr");
        qtrSales    = 0;
        qtrRev      = 0.0;
      }
    }
    area.addRow(new String[0]);
    area.addRow(String.format("{:Rb}TOTAL/%d/%.2f/%.2f",totSales,totRev / totSales,totRev).split("/"),"#TOT");
    area.writeArea().colWidth(-1,3).addDataFilterLine();
  }

  public static void main(String[] args) {
    try {
      HelloExcelNHN helloExcel = new HelloExcelNHN();
      helloExcel.run(args);
    } catch (Exception e) {
      log("HelloExcel Croaked:"+e);
    }
  }
}

