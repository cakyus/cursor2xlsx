SET DEFAULT TO JUSTPATH(SYS(16))
DO p_set.prg
DO p_clear.prg

SET PROCEDURE TO 'class_unit_test_case.prg' ADDITIVE

ErrorLogPath = 'TEST.ERR'

IF FILE(ErrorLogPath) THEN
	DELETE FILE &ErrorLogPath
ENDIF

&& <-- MAIN
TRY
	oExcelCell_TestCase = CREATEOBJECT('ExcelWrite_TestCase')
	oExcelCell_TestCase.start()
CATCH TO oErr1
	cError = '[EXCEPTION]' + ;
		CHR(10) + CHR(13) + [  Error: ] + STR(oErr1.ErrorNo) + ;
		CHR(10) + CHR(13) + [  LineNo: ] + STR(oErr1.LineNo)  + ;
		CHR(10) + CHR(13) + [  Message: ] + oErr1.Message  + ;
		CHR(10) + CHR(13) + [  Procedure: ] + oErr1.Procedure  + ;
		CHR(10) + CHR(13) + [  Details: ] + oErr1.Details  + ;
		CHR(10) + CHR(13) + [  StackLevel: ] + STR(oErr1.StackLevel)  + ;
		CHR(10) + CHR(13) + [  LineContents: ] + oErr1.LineContents 
	STRTOFILE(cError, ErrorLogPath, 0)
ENDTRY

IF FILE(ErrorLogPath) THEN
	RUN /N7 EXPLORER &ErrorLogPath
ENDIF

&& QUIT

&& MAIN -->

DEFINE CLASS ExcelWrite_TestCase As UnitTestCase

	FUNCTION Test_IllegalCharacter
	
		LOCAL oExcellWriter, cFileOutput
		
		oExcellWriter = NEWOBJECT('Excell_Writer','class_excell_writer.prg')
		CREATE CURSOR EMPLOYEES ( USER_NAME CHR(254))
		
		&& FILE TEMPORER
		cFileOutput = SYS(2023) + '\FOX' + SYS(3) + '_' + 'Test_IllegalCharacter.XLSX'
		
		SELECT EMPLOYEES
		APPEND BLANK
		REPLACE USER_NAME WITH '&'
		GO TOP
		
		oExcellWriter.SetCursor('EMPLOYEES')
		oExcellWriter.SetFileOutputPath(cFileOutput)
		oExcellWriter.Convert()
		
		RUN /N7 EXPLORER &cFileOutput
		
	ENDFUNC
	
	FUNCTION Test_UnsupportedCharacter
	
		LOCAL oExcellWriter, cFileOutput
		
		oExcellWriter = NEWOBJECT('Excell_Writer','class_excell_writer.prg')
		CREATE CURSOR EMPLOYEES ( USER_NAME CHR(254))
		
		&& FILE TEMPORER
		cFileOutput = SYS(2023) + '\FOX' + SYS(3) + '_' + 'Test_UnsupportedCharacter.XLSX'
		
		SELECT EMPLOYEES
		APPEND BLANK
		REPLACE USER_NAME WITH CHR(2)
		APPEND BLANK
		REPLACE USER_NAME WITH CHR(31)
		APPEND BLANK
		REPLACE USER_NAME WITH CHR(14)
		GO TOP
		
		oExcellWriter.SetCursor('EMPLOYEES')
		oExcellWriter.SetFileOutputPath(cFileOutput)
		oExcellWriter.Convert()
		
		RUN /N7 EXPLORER &cFileOutput
		
	ENDFUNC
	
	FUNCTION Test_ExtendedAsciiCharacter
	
		LOCAL oExcellWriter, cFileOutput
		
		oExcellWriter = NEWOBJECT('Excell_Writer','class_excell_writer.prg')
		CREATE CURSOR EMPLOYEES ( USER_NAME CHR(254))
		
		&& FILE TEMPORER
		cFileOutput = SYS(2023) + '\FOX' + SYS(3) + '_' + 'Test_ExtendedAsciiCharacter.XLSX'
		
		SELECT EMPLOYEES
		APPEND BLANK
		REPLACE USER_NAME WITH CHR(128)
		APPEND BLANK
		REPLACE USER_NAME WITH CHR(255)
		GO TOP
		
		oExcellWriter.SetCursor('EMPLOYEES')
		oExcellWriter.SetFileOutputPath(cFileOutput)
		oExcellWriter.Convert()
		
		RUN /N7 EXPLORER &cFileOutput
		
	ENDFUNC
	
	FUNCTION Test_ShowProgress
	
		LOCAL oExcellWriter
		LOCAL cFileOutput
		
		oExcellWriter = NEWOBJECT('Excell_Writer','class_excell_writer.prg')	
		
		&& APPEND DATA
		CREATE CURSOR EMPLOYEES ( USER_NAME CHAR(4) )
		FOR I = 1 TO 400
			APPEND BLANK
			REPLACE USER_NAME WITH TRIM(TRANSFORM(I))
		ENDFOR	
		GO TOP
		
		&& FILE TEMPORER
		cFileOutput = SYS(2023) + '\FOX' + SYS(3) + '_' + 'Test_ShowProgress.XLSX'
		
		&& CREATE XLSX
		oExcellWriter.SetCursor('EMPLOYEES')
		oExcellWriter.SetFileOutputPath(cFileOutput)
		oExcellWriter.ShowProgress = .T.
		oExcellWriter.Convert()
		
		&& RUN OUTPUT FILE
		IF FILE(cFileOutput)
			RUN /N7 EXPLORER &cFileOutput
		ENDIF
	ENDFUNC
ENDDEFINE
