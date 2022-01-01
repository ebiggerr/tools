
@echo off
echo "Running..."
echo "This script is for update the username and password in SAP logon vbs script..."
echo "Getting the old credentials and new credentials..."

timeout 2 > NUL

set /p oUsername="Enter the username that is being used: "
echo "Old username :"
echo %oUsername%
set /p oPassword="Enter the password that is being used: "

set /p nUsername="Enter the new username: "
echo "New username :"
echo %nUsername%
set /p nPassword="Enter the new password: "

SET oldUsername=%oUsername%
SET newUsername=%nUsername%

SET oldPassword=%oPassword%
SET newPassword=%nPassword%

echo "Running the replace.vbs"

timeout 5 > NUL

cscript replace.vbs "<pathOfTheSAPScript>" %oldUsername% %newUsername%
cscript replace.vbs "<pathOfTheSAPScript>" %oldPassword% %oldPassword%

echo "Completed"

cmd /k