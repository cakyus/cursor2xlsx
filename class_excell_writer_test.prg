SET DEFAULT TO JUSTPATH(SYS(16))
DO p_set.prg
DO p_clear.prg

SET PROCEDURE TO 'class_unit_test_case.prg' ADDITIVE

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
	STRTOFILE(cError, 'TEST.ERR', 0)
ENDTRY

QUIT

&& MAIN -->

DEFINE CLASS ExcelWrite_TestCase As UnitTestCase

	FUNCTION Test_ShowProgress
	
		LOCAL oExcellWriter, sFileOutput
		
		oExcellWriter = NEWOBJECT('Excell_Writer','class_excell_writer.prg')
		CREATE CURSOR employee (Specialty CHR(254))
		
		sFileOutput = "C:\Book1.xlsx"
		
		SELECT employee
		APPEND BLANK
		GO TOP
		mSpecialty = '&'
		REPLACE Specialty WITH mSpecialty
		
		oExcellWriter.SetCursor('employee')
		oExcellWriter.SetFileOutputPath('C:\Book1.xlsx')
		oExcellWriter.Convert()
		
		RUN /N7 EXPLORER &sFileOutput
		
	ENDFUNC
ENDDEFINE
