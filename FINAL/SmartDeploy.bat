
REM TURN SCREEN DISPLAY OFF
@echo OFF

REM "---  Create and unHide directory and copy latest file to biztech.  ---"

CLS

:: FAST DEPLOYMENT SOLUTIONS BY EROS NIKO ALVAREZ 
:: CREATED ON 01/10/2017
:: CHANGE <envFolder> to DEv or MOD or OPS.
:: BY DEFAULT IT WILL GO TO DEV ENVIRONMENT IF YOU DECLARE <envFolder> INCORRECTLY.
:: 
:: WARNING : USE THIS BATCH FILE IF YOU IMPLEMENT FASTDEPLOY MODULE
::
:: DEV -- DEVELOPER'S ENVIRONMENT 
:: MOD -- UAT ENVIRONMENT
:: OPS -- PRODUCTION
set "envFolder=MOD"

set "strPath=C:\Users\alverer\Documents\Proto-DynamicDeployment\Proto-DynamicDeployment.accdb"
set "strNormal=True"
set "strFolder=C:\$BIZTECH"
set "strFile=%strFolder%\%EnvFolder%\Proto-DynamicDeployment.accdb"

::CREATE strFolder if not exists in local C:
IF NOT EXIST "%strFolder%" (
  cd C:
  md "%strFolder%"
)

::GOTO strFolder if exists in local C:
IF EXIST "%strFolder%" (
  cd "%strFolder%"
  attrib +H "%strFolder%"
)

::CREATE envFolder if not exists in strFolder
IF NOT EXIST "%envFolder%" (
  md "%envFolder%"
)

::GOTO envFolder if exists in strFolder
IF EXIST "%envFolder%" (
  cd "%envFolder%"
  attrib +H "%strFolder%\%envFolder%"
)

IF EXIST "%strFile%" (
  del "%strFile%"
  echo Deleted File: "%strFile%"
)

COPY "%strPath%" "%strFile%"

REM "---  Execute Normal MS Office.  ---"


set "strPath=C:\Program Files\Microsoft Office XP\Office14\MSACCESS.EXE"
IF EXIST "%strPath%" (
  GoTo EndNow
)

set "strPath=C:\Program Files\Microsoft Office 2010\Office14\MSACCESS.EXE"
IF EXIST "%strPath%" (
  GoTo EndNow
)

set "strPath=C:\Program Files\Microsoft Office\Office14\MSACCESS.EXE"
IF EXIST "%strPath%" (
  GoTo EndNow
)

set "strPath=C:\Program Files (x86)\Microsoft Office\Office14\MSACCESS.EXE"
IF EXIST "%strPath%" (
  GoTo EndNow
)

REM "---  Execute Virtualized MS Office.  ---"

  REM "---  TESTING:  Open MS Access Virtualized Version in client computer.  ---"
  REM   ok (launches MS office)   - "C:\Program Files\Microsoft Application Virtualization Client\sfttray.exe" /launch "Microsoft Office Access 2003 SP3 11.0.8166.0"
  REM   ok (launches the program) - "C:\Program Files\Microsoft Application Virtualization Client\sfttray.exe" /launch "Microsoft Office Access 2003 SP3 11.8321.8341" "C:\$BIZTECH\SpecialServicesReconUI.mdb"
  REM   ok (launches the program) - START  "NBDASH" "%strPath%" "%strFile%"
  REM   ok (launches the program) - "%strPath%" "%strFile%"
  REM   ok (launches the program) - "%strPath%" /launch "%strFile%"
  REM "---  END:  Successful for 'Microsoft Office Access 2003 SP3 (V)' version.  ---"

set "strPath=C:\Program Files\Microsoft Application Virtualization Client\sfttray.exe"
IF EXIST "%strPath%" (
  set "strNormal=False"
  GoTo EndNow
)

set "strPath=C:\Program Files (x86)\Microsoft Application Virtualization Client\sfttray.exe"
IF EXIST "%strPath%" (
  set "strNormal=False"
  GoTo EndNow
)

set "strPath=%LOCALAPPDATA%\Microsoft\AppV\Client\Integration\B5E07289-280C-4C22-8919-63FA6B4364EA\Root\VFS\ProgramFilesX86\Microsoft Office\Office14\MSACCESS.EXE"
IF EXIST "%strPath%" (
  GoTo EndNow


:EndNow

echo "%strNormal%" "%strPath%"

IF "%strNormal%"=="True" (
  START  "NBDASH" "%strPath%" "%strFile%" 
) ELSE IF "%strNormal%"=="False" (
  "%strPath%" /launch "%strFile%"
)

REM "---  Hide the folder.  ---"

attrib +H "%strFolder%"

REM TURN SCREEN DISPLAY ON
REM SETLOCAL DisableDelayedExpansion
@echo ON
REM echo &PAUSE&GOTO:EOF
echo &GOTO:EOF