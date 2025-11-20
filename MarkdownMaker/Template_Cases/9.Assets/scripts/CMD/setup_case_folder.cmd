@echo off
REM Ask for the case number
set /p caseNumber=Enter case number: 

REM Ask for the case name
set /p caseName=Enter case name: 

REM Create the case directory structure under "Cases" and "CaseWF"

REM Check and create the "Cases" directory if it doesn't exist
if not exist "Cases" mkdir "Cases"

REM Check and create the "CaseWF" directory if it doesn't exist
if not exist "CaseWF" mkdir "CaseWF"

REM Check and create the "<caseNumber>_<caseName>" directory under "Cases"
if not exist "Cases\%caseNumber%_%caseName%" mkdir "Cases\%caseNumber%_%caseName%"

REM Check and create the "%caseNumber%_%caseName%_CaseNotes" directory under "Cases\<caseNumber>_<caseName>"
if not exist "Cases\%caseNumber%_%caseName%\%caseNumber%_%caseName%_CaseNotes" mkdir "Cases\%caseNumber%_%caseName%\%caseNumber%_%caseName%_CaseNotes"

REM Check and create the "DigitalEvidence" directory under "Cases\<caseNumber>_<caseName>"
if not exist "Cases\%caseNumber%_%caseName%\DigitalEvidence" mkdir "Cases\%caseNumber%_%caseName%\DigitalEvidence"

REM Check and create the "Ex" directory under "Cases\<caseNumber>_<caseName>"
if not exist "Cases\%caseNumber%_%caseName%\Ex_" mkdir "Cases\%caseNumber%_%caseName%\Ex_"

REM Check and create the "Photos" directory under "Cases\<caseNumber>_<caseName>\DigitalEvidence"
if not exist "Cases\%caseNumber%_%caseName%\DigitalEvidence\Photos" mkdir "Cases\%caseNumber%_%caseName%\DigitalEvidence\Photos"

REM Check and create the "Image" directory under "Cases\<caseNumber>_<caseName>\Ex_"
if not exist "Cases\%caseNumber%_%caseName%\Ex_\Image" mkdir "Cases\%caseNumber%_%caseName%\Ex_\Image"

REM Check and create the "<caseNumber>_<caseName>" directory under "CaseWF"
if not exist "CaseWF\%caseNumber%_%caseName%" mkdir "CaseWF\%caseNumber%_%caseName%"


REM Check and create the "Ex_" directory under "CaseWF\<caseNumber>_<caseName>"
if not exist "CaseWF\%caseNumber%_%caseName%\Ex_" mkdir "CaseWF\%caseNumber%_%caseName%\Ex_"

REM Check and create the "Warrant" directory under "CaseWF\<caseNumber>_<caseName>"
if not exist "CaseWF\%caseNumber%_%caseName%\Warrant" mkdir "CaseWF\%caseNumber%_%caseName%\Warrant"


REM Check and create the "Case" directory under "CaseWF\<caseNumber>_<caseName>\Ex_"
if not exist "CaseWF\%caseNumber%_%caseName%\Ex_\Case" mkdir "CaseWF\%caseNumber%_%caseName%\Ex_\Case"


REM Check and create the "Exports" directory under "CaseWF\<caseNumber>_<caseName>\Ex_"
if not exist "CaseWF\%caseNumber%_%caseName%\Ex_\Exports" mkdir "CaseWF\%caseNumber%_%caseName%\Ex_\Exports"

REM Check and create the "imageCOPY" directory under "CaseWF\<caseNumber>_<caseName>\imageCOPY"
if not exist "CaseWF\%caseNumber%_%caseName%\Ex_\imageCOPY" mkdir "CaseWF\%caseNumber%_%caseName%\Ex_\imageCOPY"

echo Directory structure for case %caseNumber%_%caseName% created successfully.


