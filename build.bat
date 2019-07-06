@echo off


@echo builder invoked

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


set javac=%java_home%\bin\javac.exe
set jar=%java_home%\bin\jar.exe

%javac% -cp %str% -sourcepath src -d cls src\com\psec\excel\*.java src\com\psec\util\*.java src\com\psec\run\*.java

%jar% cf run/lib/psec-java.jar -C cls .


