@echo Pre-Commit workbooks clean-up is executed!!!
@cscript CleanupWorkbook.vbs BinanceActivity.xlsm
@IF %ERRORLEVEL% EQU 0 (
    @echo Cleaning-up BinanceActivity.xlsm failed
    @exit 1
) ELSE (
    @echo Cleaning-up BinanceActivity.xlsm succeeded
    @exit 0
)
