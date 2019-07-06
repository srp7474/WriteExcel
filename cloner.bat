@echo off
set what=%0
echo %what% starting
call runner %what% >run\out\log.%what%.txt
echo %what completed