#!/bin/sh
echo "Executing Shell pre-commit script for Windows"

cmd.exe //c pre-commit.cmd

if [ $? -eq 0 ]; then
    echo "Windows 'pre-commit.cmd' file execution succeeded - continue commit"
    exit 0
else
    echo "Windows 'pre-commit.cmd' file execution failed - stop commit"
    exit 1
fi
