// Parm - Standard Parm Processor

// Copyright (c) 2000-2002 AltCode Systems Inc, All Rights Reserved.
// Copyright (c) 2002-2015 Rexcel Systems Inc, All Rights Reserved.

/*
  History:
    EC9101 - Add Brief notion
    EC8803 - Add Return Status, Run-Status

    EC0101 - Original from AltCode
*/
/*
 @license
 Copyright (c) 2019 by Steve Pritchard of Rexcel Systems Inc.
 This file is made available under the terms of the Creative Commons Attribution-ShareAlike 3.0 license
 http://creativecommons.org/licenses/by-sa/3.0/.
 Contact: public.pritchard@gmail.com
*/
package com.psec.util;
import java.util.HashMap;
import java.util.Map;
import java.util.Iterator;
import java.io.FileWriter;
import java.io.File;
import com.google.gson.Gson;
import com.google.gson.GsonBuilder;


/**
 <h4>Table based Parms Handler</h4>
 The <code>Parm</code> class and its nested classes and interfaces provide a mechanism that allows for convenient and consistent
 table based program arguments handling and processing.
 <p>
 The <code>Parm[]</code> table is matched against the invokation arguments. Errors are reported (such as the absence
 of a mandatory field).
 <p>
 If the arguments parse is successful, the class run method is invoked.

 <h4>The Parm Table</h4>
 The <code>Parm</code> table consists of an array of <code>Parm</code> objects, each constructed using an anonymous instance of the {@link Parm.PS} interface.
 <p>
 The result of the parse is the customized {@link Parm.Opt} object where each field is populated according to the <code>String[] args</code> parameter provided to the class <code>main</code> method.
 <p>
 To improving the readability of the table, the symbolic names are short so that each option can be conveniently specified on one text line.
 <h4>Input Conventions</h4>
 <ul>
 <li>Keywords are lower case, specified in any order</li>
 <li>The last quietly overides the former in the case of duplication on the command line</li>
 <li>Values are input <code>-key space(s) value</code></li>
 <li>Booleans are input <code>-key</code> (no trailing value)</li>
 <li>Properies are input <code>-def<i>name=value</i></code> (no intermediate space, -def used by convention but it could be some other such as <code>-D</code></li>

 </ul>

 <p>
 <b>Example for below</b>
 <p>
 <code>-brief c:\temp\brief.json.txt -out d:\1\out.xlsx -what basic -red -negfmt "()"</code>

 <h4>Working Example</h4>
 Here is an example of a table used in {@link com.psec.run.DemoExcel DemoExcel}.
 <p>
 <pre style="font-size:80%;">
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
 </pre>


 */

public class Parm {

  /**
    Base class for option set.
    <p>
    It is intended that this base class be extended with custom fields (and even methods if needed) according to the program requirements.
    <p>
    This also serves to supply the initialization methods to populate the options and
    afterwards invoke the custom run time routines that process based on these options.
  */
  public static class Opt { // Add common property support
    /** Trigger dump of options when true. When true the Parm parser will trigger a dump of the possible options and terminate. */
    public boolean bHelp = false;
    /** Fully qualified name of Brief file. Populated by {@link Parm.Opt#setBrief setBrief} method.*/
    public String  sBrief = null;
    private Brief  oBrief = null; // populated if to produce Brief
    HashMap<String,String> oProps = new HashMap<String,String>();

    /**
      Parse and add property string.
      After removing the prefix, by convention <code>-def</code>,
      the <code>sProp</Code> string is split on the first equal sign (=), the left being the name and the right being the value.
      <p>
      Multiple values are permitted.  The last silently replaces the former.
      <p>
      Inproperly formatted strings are silently ignored.
      @param sProp the property string.
    */
    public void addProp(String sProp) {
      int nIX = sProp.indexOf("=");
      if (nIX > 0) {
        String sKey = sProp.substring(0,nIX);
        String sValue = sProp.substring(nIX+1);
        oProps.put(sKey,sValue);
        //log("property "+sKey+"="+sValue);
      }
    }

    /**
      Print to System.out standard introductory line. The <code>self</code> parameter is used to
      determine the class.
      @param self The class to which this message applies.
    */
    public void selfIntro(Object self) {
      System.out.println(""+self.getClass().getSimpleName()+" started. Parms:");
    }

    /**
      Return property value for <code>sKey</code> or null;
      @param sKey The case sensitive property name.
    */
    public String getProp(String sKey) {
      return getProp(sKey,null);
    }

    /**
      Return property value for <code>sKey</code> or sDef;
      @param sKey The case sensitive property name.
      @param sDef The default value to return if <code>sKey</code> does not exist.
    */
    public String getProp(String sKey,String sDef) {
      String sProp = oProps.get(sKey);
      if (sProp == null) return sDef;
      return sProp;
    }

    /**
      Return an array of property value names;
      @return the String array of names.
    */
    public String[] getProps() {
      String[] sKeys = new String[oProps.size()];
      int i = 0;
      for(Iterator<Map.Entry<String,String>> oI = oProps.entrySet().iterator(); oI.hasNext();) {
        Map.Entry<String,String> oE = oI.next();
        sKeys[i++] = oE.getKey();
      }
      return sKeys;
    }

    /**
      Return whether a property value represents <code>true</code>. A true value starts with "t", "T", "y" or "Y".
      @param sProp The case sensitive property name.
      @return the result.
    */
    public boolean isTrue(String sProp) {
      String sStr = getProp(sProp,"f").toUpperCase();
      if (sStr.startsWith("T")) return true;
      if (sStr.startsWith("Y")) return true;
      return false;
    }

    // EC8803 - refactor in standard processing
    /**
      The <code>mainStart</code> method handles the standard startup and termination procedures.
      <p>
      First the parameters are passed using {@link Parm#processArguments processArguments}.
      <p>
      If the parse fails or the {@link Parm.Opt#bHelp} flag is true, the {@link Parm#printHelp printHelp} method is called and the run terminates.
      <p>
      Otherwise the {@link Parm.MainRun#run run} method is called.
      <p>
      When the {@link Parm.MainRun#run run} method returns and the internal <code>oBrief</code> object populated by the {@link Parm.Opt#setBrief setBrief}
      method call is populated, it is transformed to a JSON object using GsonBuilder and written to the file specified by {@link Parm.Opt#sBrief sBrief}.
      <p>
      Finally, the run terminates.
      @param oParms The option table to parse.
      @param sArgs  The arguments to match.
      @param oRun   The <code>MainRun</code> mainline instance on which to invoke the {@link MainRun#run run} method.
    */
    public void mainStart(Parm[] oParms,String[] sArgs,MainRun oRun) {
      String sName = oRun.getClass().getSimpleName();
      try {
        Parm.processArguments(sArgs,oParms);
        if (this.bHelp) {
          Parm.printHelp(oParms,sName);
          System.exit(901);
        }
        Parm.checkParms(oParms);
        oRun.run(this);
        if (oBrief != null) {  // Populated if we are to emit Brief json file
          this.emitResultsBrief(oBrief);
        }
      } catch (Exception e) {
        System.out.println("ERROR:"+e);
        e.printStackTrace();
        System.exit(900);
      }
    }

    /**
      Set the internal <code>oBrief</code> value to paramater <code>oBrief</code> value.
      <p>
      The internal {@link Parm.Opt#sBrief sBrief} value must be populated or an Exception occurs.
      <p>
      If the physical files exists it is deleted.  This means that at the end of the run the external routines
      can know that the non-existance of the Brief file implies a run failure.
      <p>
      For this protcol to work reliably, this method call
      should be issued at the beginning of the run invokation.
      @param oBrief The {@link Parm.Brief Brief} extended instance.
    */
    public void setBrief(Brief oBrief) throws Exception {
      this.oBrief = oBrief;
      if (sBrief == null) throw e("Required Brief FileName not specified (-brief)");
      File oFile = new File(sBrief);
      if (oFile.exists()) oFile.delete();
    }

    private void emitResultsBrief(Brief oBrief) throws Exception {
      Gson oG = new GsonBuilder().setPrettyPrinting().disableHtmlEscaping().create();
      String sJson = oG.toJson(oBrief);
      FileWriter oFW = new FileWriter(sBrief);
      oFW.write(""+sJson);
      oFW.close();
    }
  }



  // Used to process each possible option
  /**
    Interface used to process each option. A static reference to the <code>Parm.Opt</code> object instance being constructed allows the set method
    to set a specific value in that <code>Parm.Opt</code> object.
    <p>
  */
   // There is an implicit connection between the input data and the target variable based on the
   // position of the anonymous Parm.PS object instance created in the table. This connection is used to
  public static interface PS {
    /**
    The set method is used to move (and possibly transform) the <code>sNext</code> value to the static <code>Parm.Opt</code> object instance being constructed.
    The following conventions are used in handling the data.  Others could be adopted.
    <p>
    <ul>
    <li>Boolean targets are set to <code>true</code> if the set is invoked.</li>
    <li>Property targets call the {@link Parm.Opt#addProp addProp} to process the <code>sNext</code> value.
    <li>Otherwise the string is moved as is to the target variable it represents. If the target is an integer or double value, it could be converted
    to the proper format at this point.
    </li>
    </ul>
    @param sNext The non-null String to process.  The key prefix will have been removed.
    @see <a
    */
    public void set(String sNext);
  }

  /**
    The interface class to be implemented by the class that is to be called by the {@link Parm.Opt#mainStart mainStart} method.
  */
  public static interface MainRun {
    /**
     <code>run</code> method called by {@link Parm.Opt#mainStart mainStart}. This should be the beginning of the program processing.
     @param oOpt The {@link Parm.Opt extended options} object.
   */
    public void run(Parm.Opt oOpt) throws Exception;
  }

  /**
    Brief base class used by PSEC routines.
    <p>
    It is intended that this be extended with custom fields according to the processing requirements.
    <p>
    The object is converted to a JSON object and writen to the {@link Parm.Opt#sBrief sBrief} file  by the {@link Parm.Opt#mainStart mainStart} processing routines.
    This file can then be inspected  by the external routines that initiated the process.
  */
  public static class Brief {
    /** Final condition. Sould be set to GOOD or FAIL. */
    public String sCondCode = "FAIL"; // Final condition GOOD,FAIL
    /** Reason for failure. Meaningful when sCondCode is set to FAIL. */
    public String sReason;            // Reported failure reason
  }

  char cFlag;
  String sParm;
  String sDesc;
  String sExtra;
  PS     oSet;
  boolean bHave = false;

  /**
  Parm instance constructor.
  The table definition employs the following conventions to create a consistent and predictable interface with the user. Each Parm entry follows
  the convention.
  <p>
   The {@link Parm constructor} has 4 input parameters.
   <ol>
     <li>The the anonymous function constructed to move (and possibly transform) the input string to the target object as
     detailed in the {@link Parm.PS#set set} method.</li>
     <li>The </li>
   </ol>
     @param oSet  The anonymous <code><Parm.PS</code> used to transform the input data.
     @param sFlag The mandatory flag.
     @param sParm The external key string used to recognize the parameter.
     @param sExtra A string to indicate the value type.
     @param sDesc A brief description that helps the user know the meaning and format of the data.
  */

  public Parm(PS oSet,String sFlag,String sParm,String sExtra,String sDesc) {
    this.cFlag  = sFlag.charAt(0);
    this.sParm  = sParm;
    this.sExtra = sExtra;
    this.sDesc  = sDesc;
    this.oSet   = oSet;
    if (cFlag >= '0' && cFlag <= '9') cFlag &= 0x0F;
  }

  /**
    List the options.
    The options are listed in a manner suitable for a user to know how to
    supply a particular parameter.
    @param oParms the <code>Parm[]</code> table;
    @param sTitle The list title.
  */
  public static void printHelp(Parm[] oParms,String sTitle) {
    System.out.println(sTitle+" Parameters:");
    for(int i=0,iMax=oParms.length; i<iMax; i++) {
      Parm oParm = oParms[i];
      String sStr = oParm.sParm+" ";
      if (oParm.sExtra != null) sStr += oParm.sExtra;
      sStr = (sStr+"                             ").substring(0,15)+" ";
      sStr += oParm.sDesc;
      if (oParm.cFlag == 'm') sStr += " (mandatory)";
      System.out.println(sStr);
    }
  }

  private static void checkParms(Parm[] oParms) throws Exception {
    boolean bError = false;
    for(int i=0,iMax=oParms.length; i<iMax; i++) {
      Parm oParm = oParms[i];
      if ((oParm.cFlag == 'm') && (!oParm.bHave)) {
        bError = true;
        System.out.println(oParm.sParm+" parameter not specified and is required");
      }
      if (bError) throw e("Terminated due to parm error. Use -h to see parm requirements");
    }
  }

  private static void processArguments(String[] oArgs,Parm[] oParms) throws Exception{
    int nParm = 0;
    while(true) {
      if (nParm >= oArgs.length) break;
      boolean bFound = false;
      for(int i=0,iMax=oParms.length; i<iMax; i++) {
        Parm oP = oParms[i];
        String sArg = oArgs[nParm];
        if ((oP.cFlag > 0) && (oP.cFlag <= 9)) {
          if (sArg.startsWith(oP.sParm.substring(0,oP.cFlag))) {
            bFound = true;
            if (sArg.length() == oP.cFlag) {
              throw e(oP.sParm.substring(0,oP.cFlag)+" requires data (no space after). Use -h to see format");
            }
            oP.oSet.set(sArg.substring(oP.cFlag));
            oP.bHave = true;
            break;
          }
        } else {
          if (oP.sParm.equals(sArg)) {
            String sNext = null;
            if (oP.sExtra != null) {
              nParm++;
              if (nParm >= oArgs.length) throw e("Expect "+oP.sExtra+" parm after "+oP.sParm);
              sNext = oArgs[nParm];
            }
            oP.oSet.set(sNext);
            oP.bHave = true;
            bFound = true;
            break;
          }
        }
      }
      if (!bFound) throw e("Parm "+oArgs[nParm]+" not known, use -h to see parm requirements");
      nParm++;
    }
  }

  private static Exception e(String sMsg) {return new Exception(sMsg);}

}
