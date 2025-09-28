#!/bin/bash
# This hook executes the Windows batch script hook
echo "Executing Shell pre-commit script"

FILE_TO_CHECK="BinanceActivity.xlsm"

if git diff --cached --name-only | grep -q "$FILE_TO_CHECK"; then
  echo "The file '$FILE_TO_CHECK' IS marked for commit - Clean Data Check WILL BE executed"
else
  echo "The file '$FILE_TO_CHECK' IS NOT marked for commit - Clean Data Check WILL NOT BE executed"
  exit 0
fi

cmd.exe //c pre-commit.cmd

if [ $? -eq 0 ]; then
	echo "Windows 'pre-commit.cmd' file execution succeeded - continue commit"
	exit 0
else
	echo "Windows 'pre-commit.cmd' file execution failed - stop commit"
	exit 1
fi
