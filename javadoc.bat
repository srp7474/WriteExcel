@echo off


@echo javadoc creator invoked

set javadoc=%java_home%\bin\javadoc.exe

%javadoc% -public -sourcepath src -subpackages com -d javadocs -overview src/html/overview.htm


Echo NOTE: 62 errors is correct