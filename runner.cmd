@echo off

if "%1" EQU "" (
  echo requires parameter
  goto :EOF
)

@echo runner invoked %1

set lib=run\lib

set str=%lib%\psec-java.jar
set str=%str%;%lib%\gson-2.8.1.jar

set str=%str%;%lib%\poi-4.0.0.jar
set str=%str%;%lib%\poi-ooxml-4.0.0.jar
set str=%str%;%lib%\poi-excelant-4.0.0.jar
set str=%str%;%lib%\poi-scratchpad-4.0.0.jar
set str=%str%;%lib%\poi-ooxml-schemas-4.0.0.jar
set str=%str%;%lib%\commons-collections4-4.2.jar
set str=%str%;%lib%\xmlbeans-3.0.1.jar
set str=%str%;%lib%\commons-compress-1.18.jar
set str=%str%;%lib%\commons-math3-3.6.1.jar

rem CP %str%

set javaexe=%java_home%\jre\bin\java.exe

%javaexe% -cp %str% com.psec.run.DemoExcel -brief run\out\brief.%1.json.txt -what %1 -out run\out\out.%1.xlsx -inp run\inp\demo-excel-inp.xlsx
rem java -cp $this.vbls.cp $this.vbls.main $($parms) 2>&1