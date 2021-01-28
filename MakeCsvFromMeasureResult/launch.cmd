@echo off
echo "Make CSV From Measurement result txt data!!"
pushd %~dp0
pwsh -Noprofile -ExecutionPolicy RemoteSigned -File ./main.ps1
pause > nul
exit