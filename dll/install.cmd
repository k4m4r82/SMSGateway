cls
echo. install dhRichClient
pause
copy dhRichClient3.dll %systemroot%\system32
copy sqlite36_engine.dll %systemroot%\system32
copy ASmsCtrl.dll %systemroot%\system32
regsvr32 /s %systemroot%\system32\dhRichClient3.dll
regsvr32 /s %systemroot%\system32\ASmsCtrl.dll