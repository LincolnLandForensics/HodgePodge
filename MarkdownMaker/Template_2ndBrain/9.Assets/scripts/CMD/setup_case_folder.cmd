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

REM Check and create the "Notes_%caseNumber%_%caseName%" directory under "Cases\<caseNumber>_<caseName>"
if not exist "Cases\%caseNumber%_%caseName%\Notes_%caseNumber%_%caseName%" mkdir "Cases\%caseNumber%_%caseName%\Notes_%caseNumber%_%caseName%"

REM Check and create the "DigitalEvidence" directory under "Cases\<caseNumber>_<caseName>"
if not exist "Cases\%caseNumber%_%caseName%\DigitalEvidence" mkdir "Cases\%caseNumber%_%caseName%\DigitalEvidence"

REM Check and create the "Ex" directory under "Cases\<caseNumber>_<caseName>"
if not exist "Cases\%caseNumber%_%caseName%\Ex" mkdir "Cases\%caseNumber%_%caseName%\Ex"

REM Check and create the "Photos" directory under "Cases\<caseNumber>_<caseName>\DigitalEvidence"
if not exist "Cases\%caseNumber%_%caseName%\DigitalEvidence\Photos" mkdir "Cases\%caseNumber%_%caseName%\DigitalEvidence\Photos"

REM Check and create the "Image" directory under "Cases\<caseNumber>_<caseName>\Ex"
if not exist "Cases\%caseNumber%_%caseName%\Ex\Image" mkdir "Cases\%caseNumber%_%caseName%\Ex\Image"

REM Check and create the "<caseNumber>_<caseName>" directory under "CaseWF"
if not exist "CaseWF\%caseNumber%_%caseName%" mkdir "CaseWF\%caseNumber%_%caseName%"


REM Check and create the "Ex" directory under "CaseWF\<caseNumber>_<caseName>"
if not exist "CaseWF\%caseNumber%_%caseName%\Ex" mkdir "CaseWF\%caseNumber%_%caseName%\Ex"


REM Check and create the "Case" directory under "CaseWF\<caseNumber>_<caseName>\Ex"
if not exist "CaseWF\%caseNumber%_%caseName%\Ex\Case" mkdir "CaseWF\%caseNumber%_%caseName%\Ex\Case"


REM Check and create the "Exports" directory under "CaseWF\<caseNumber>_<caseName>\Ex"
if not exist "CaseWF\%caseNumber%_%caseName%\Ex\Exports" mkdir "CaseWF\%caseNumber%_%caseName%\Ex\Exports"

REM Check and create the "imageCOPY" directory under "CaseWF\<caseNumber>_<caseName>\imageCOPY"
if not exist "CaseWF\%caseNumber%_%caseName%\Ex\imageCOPY" mkdir "CaseWF\%caseNumber%_%caseName%\Ex\imageCOPY"

echo Directory structure for case %caseNumber%_%caseName% created successfully.


