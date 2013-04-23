@ECHO OFF
cls
ECHO *************************************
ECHO *************************************
ECHO **                                 **
ECHO **      Creating Shortcut          **
ECHO **                                 **
ECHO *************************************
ECHO -------------------------------------
copy "Open Order Report.lnk" "%userprofile%\Desktop"
ECHO -------------------------------------
ECHO.
ECHO.
ECHO.
ECHO Press any key to quit..
pause > nul