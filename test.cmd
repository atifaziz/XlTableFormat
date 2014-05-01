@echo off
setlocal
call build && call :test Debug %* && call :test Release %*
goto :EOF

:test
setlocal
cd bin\%1
 call test %2 %3 %4 %5 %6 %7 %8 %9 > test.log ^
 && findstr "/c:Errors and Failures" test.log > nul && (type test.log & exit /b 1)
type test.log
goto :EOF
