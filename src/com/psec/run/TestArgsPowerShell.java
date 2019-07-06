// TestArgsPowerShell - Stub to report powershell arguments

/* This is also used to test the Java infrastructure as to passing back
 * completion status and completion-status-json files
 *
  History:
    EC9203 - improved for PSEC

    EC8715 - added completion status

    EC8715 - original
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
import  com.psec.util.OptionalFlag;
import java.util.ArrayList;

/**
Used by the PSEC utility to exercise the parameter passing functions.
*/
public class TestArgsPowerShell implements Parm.MainRun {


  private static void log(String sMsg) {System.out.println(sMsg);}
  private static Exception e(String s) {return new Exception(s); }

  private static class Brief extends Parm.Brief {
    String  sStr1;
    String  sGen;
    String  sGenStr;
    String  sCombo;
    boolean bBool1;
    @OptionalFlag
    String[] sGenArr;
    String[] sDef;
    boolean bTrap;
  }

  Opt   oOpt;
  Brief oBrief = new Brief();
  // -------- Mainline
  public void run(Parm.Opt oBaseOpt) throws Exception {
    this.oOpt = (Opt)oBaseOpt;
    log("TestArgsPowerShell starting,"
      +"\r\n   sStr1= "+oOpt.sStr1
      +"\r\n    sGen= "+oOpt.sGen
      +((oOpt.sGenStr != null)?"\r\n sGenStr= "+oOpt.sGenStr:"")
      +((oOpt.sTrap != null)?"\r\n   trap= "+oOpt.sTrap:"")
      +"\r\n   bool1="+oOpt.bBool1);
    this.oOpt.setBrief(oBrief);
    if (oBrief.sStr1 == null);
    if (oBrief.sGen == null);
    if (oBrief.sGenStr == null);
    if (oBrief.sCombo == null);
    if (oBrief.bBool1);
    if (oBrief.bTrap);

    for(String sProp:oOpt.getProps()) {
      log(" property "+sProp+"="+oOpt.getProp(sProp));
    }
    oBrief.sStr1   = oOpt.sStr1;
    oBrief.sGen    = oOpt.sGen;
    oBrief.sGenStr = oOpt.sGenStr;
    oBrief.sCombo  = oOpt.sCombo;
    oBrief.bBool1  = oOpt.bBool1;
    String[] sProps = oOpt.getProps();
    if (sProps.length > 0) {
      oBrief.sDef = new String[sProps.length];
      for(int i=0,iMax=sProps.length; i<iMax; i++) {
        String sKey=sProps[i];
        String sVal= oOpt.getProp(sKey);
        oBrief.sDef[i] = sKey+"="+sVal;
      }
    }
    oBrief.sCondCode = "GOOD";
    if (oOpt.sGen.matches("^[0-9]+$")) {
      int nTimes = Integer.parseInt(oOpt.sGen);
      for(int i=1,iMax=nTimes; i<=iMax; i++) {
        log(oOpt.sGenStr+" "+i);
      }
    }
    if (oOpt.sGenBrf.matches("^[0-9]+$")) {
      int nTimes = Integer.parseInt(oOpt.sGenBrf);
      ArrayList<String> oLst = new ArrayList<>();
      for(int i=1,iMax=nTimes; i<=iMax; i++) {
        oLst.add(oOpt.sGenStr+" "+i);
      }
      if (nTimes > 0) oBrief.sGenArr = (String[])oLst.toArray(new String[oLst.size()]);
    }
    if (oOpt.sTrap != null) {
      if (!oOpt.sTrap.equals("trap")) {
        oBrief.sCondCode = "FAIL";
        oBrief.sReason = "Fail Code "+oOpt.sTrap;
      }
      if (oOpt.sTrap.equals("trap")) {
        oBrief.bTrap = true;
        throw e("TRAP:"+oOpt.sTrap);
      }
    }
  }

/*************************************************************************
**************************************************************************
**********                START-UP ROUTINES                     **********
**************************************************************************
**************************************************************************/

  private static class Opt extends Parm.Opt {
    public String  sStr1;
    public String  sGen    = "1";
    public String  sGenBrf = "0";
    public String  sGenStr = "gen default string ";
    public String  sCombo;
    public String  sTrap;
    public boolean bBool1;
  }

  static Opt oOPT = new Opt();

  static Parm[] oParms = {
     new Parm(new Parm.PS(){public void set(String sN) {oOPT.bHelp=true; }},     ".","-h",      null,      "Generate this Help Information that you are now reading")
    ,new Parm(new Parm.PS(){public void set(String sN) {oOPT.sBrief=sN;}},       "m","-brief",  "brief",   "path for Brief output")
    ,new Parm(new Parm.PS(){public void set(String sN) {oOPT.sStr1=sN;}},        "m","-str",    "string",  "mandatory string test")
    ,new Parm(new Parm.PS(){public void set(String sN) {oOPT.sGenStr=sN;}},      ".","-genstr", "string",  "output str for '-gen,-genbrf' options")
    ,new Parm(new Parm.PS(){public void set(String sN) {oOPT.sGen=sN;}},         ".","-gen",    "integer", "generate n lines of output")
    ,new Parm(new Parm.PS(){public void set(String sN) {oOPT.sGenBrf=sN;}},      ".","-genbrf", "integer", "generate n lines of brief output")
    ,new Parm(new Parm.PS(){public void set(String sN) {oOPT.sTrap=sN;}},        ".","-trap",   "string",  "cause Exception return when supplied")
    ,new Parm(new Parm.PS(){public void set(String sN) {oOPT.sCombo=sN;}},       ".","-combo",  "string",  "combo flag setting")
    ,new Parm(new Parm.PS(){public void set(String sN) {oOPT.bBool1=true;}},     ".","-bool1",   null,     "optional boolean flag")
    ,new Parm(new Parm.PS(){public void set(String sN) {oOPT.addProp(sN);}},     ".","-def",    "key=val", "Defined Property string")
  };

  public static void main(String[] sArgs) {
    oOPT.mainStart(oParms,sArgs,new TestArgsPowerShell());
  }
}

