@ECHO OFF

SET VFP_EXE_PATH=C:\Program Files\Microsoft Visual FoxPro 9\vfp9.exe

IF "%1"=="" GOTO MAKE_HELP
IF "%1"=="run" GOTO MAKE_RUN
IF "%1"=="clean" GOTO MAKE_CLEAN
GOTO MAKE_HELP

:MAKE_RUN
IF "%2"=="" GOTO MAKE_HELP
CALL:MAKE_CLEAN
IF EXIST "%2" "%VFP_EXE_PATH%" "%2"
GOTO MAKE_END

:MAKE_CLEAN
FOR %%P IN (*.fxp) DO DEL "%%P"
FOR %%P IN (*.bak) DO DEL "%%P"
GOTO MAKE_END

:MAKE_HELP
ECHO syntax: make ^<command^>
ECHO command:
ECHO     run ^<file^>  run prg from command line
ECHO     clean         delete unnecessary files
GOTO MAKE_END

:MAKE_END
