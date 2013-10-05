SET DEFAULT TO JUSTPATH(SYS(16))
DO p_set.prg
DO p_clear.prg

SET PROCEDURE TO 'class_unit_test_case.prg' ADDITIVE

&& -- MAIN --
TRY
	oExcelCell_TestCase = CREATEOBJECT('ExcelCell_TestCase')
	oExcelCell_TestCase.start()
CATCH TO oErr1
	? [EXCEPTION]
	?[  Error: ] + STR(oErr1.ErrorNo)
	?[  LineNo: ] + STR(oErr1.LineNo) 
	?[  Message: ] + oErr1.Message 
	?[  Procedure: ] + oErr1.Procedure 
	?[  Details: ] + oErr1.Details 
	?[  StackLevel: ] + STR(oErr1.StackLevel) 
	?[  LineContents: ] + oErr1.LineContents 
ENDTRY


DEFINE CLASS ExcelCell_TestCase As UnitTestCase

	FUNCTION Test_HTMLEncode_PerMillion
	
		oExcellCell= NEWOBJECT('Excell_Cell','class_excell_cell.prg')
		
		TEXT TO sTextInput NOSHOW PRETEXT 2
			Denda 1 ‰ untuk setiap hari keterlambatan dari nilai pekerjaan yang belum diselesaikan
		ENDTEXT
		
		TEXT TO sTextOutput NOSHOW PRETEXT 2
			Denda 1 &#0137; untuk setiap hari keterlambatan dari nilai pekerjaan yang belum diselesaikan
		ENDTEXT
		
		sTextResult = oExcellCell.HTMLEncode(sTextInput)
		THIS.assertEqual(sTextResult, sTextOutput)
	ENDFUNC
	
	FUNCTION Test_HTMLEncode_XML_Character_References
	
		oExcellCell= NEWOBJECT('Excell_Cell','class_excell_cell.prg')
		
		TEXT TO sTextInput NOSHOW PRETEXT 2
			& < > " '
		ENDTEXT
		
		TEXT TO sTextOutput NOSHOW PRETEXT 2
			&amp; &lt; &gt; &quot; &apos;
		ENDTEXT
		
		sTextResult = oExcellCell.HTMLEncode(sTextInput)
		THIS.assertEqual(sTextResult, sTextOutput)
	ENDFUNC
	
	&& Single-stroke type-able keyboard characters
	
	FUNCTION Test_HTMLEncode_Keyboard_Characters
	
		oExcellCell= NEWOBJECT('Excell_Cell','class_excell_cell.prg')
		
		TEXT TO sTextInput NOSHOW PRETEXT 2
			1234567890  !@#$%^&*() abcdefghijklmnopqrstuvwxyz ABCDEFGHIJKLMNOPQRSTUVWXYZ ~_+{}|:"<>? `-=[]\;',./
		ENDTEXT
		
		TEXT TO sTextOutput NOSHOW PRETEXT 2
			1234567890  !@#$%^&amp;*() abcdefghijklmnopqrstuvwxyz ABCDEFGHIJKLMNOPQRSTUVWXYZ ~_+{}|:&quot;&lt;&gt;? `-=[]\;&apos;,./
		ENDTEXT
		
		sTextResult = oExcellCell.HTMLEncode(sTextInput)
		THIS.assertEqual(sTextResult, sTextOutput)
	ENDFUNC	
	
	FUNCTION Test_HTMLEncode_AltCodes
		LOCAL oExcellCell, sTextInput, sTextOutput 
		LOCAL sCRC32 
	
		oExcellCell= NEWOBJECT('Excell_Cell','class_excell_cell.prg')
		
		sTextInput = FILETOSTR('alt-codes-utf-8.txt')
		sTextResult = oExcellCell.HTMLEncode(sTextInput)
		sTextResult = THIS.getCRC32(sTextResult)
		sTextOutput = 'A3108302'
		
		THIS.assertEqual(sTextResult, sTextOutput)
	ENDFUNC
	
	FUNCTION Test_Null_Character
		LOCAL oExcellCell, cText
		oExcellCell= NEWOBJECT('Excell_Cell','class_excell_cell.prg')
		oExcellCell.Value = .NULL.
		oExcellCell.FoxproFieldType = 'C'
		cText = oExcellCell.GetFieldValueString()
		THIS.assertEqual(cText, '')
	ENDFUNC
	
	FUNCTION Test_Null_Numeric
		LOCAL oExcellCell, cText
		oExcellCell= NEWOBJECT('Excell_Cell','class_excell_cell.prg')
		oExcellCell.Value = .NULL.
		oExcellCell.FoxproFieldType = 'N'
		cText = oExcellCell.GetFieldValueString()
		THIS.assertEqual(cText, '0')
	ENDFUNC
	
	FUNCTION Test_Null_Date
	
		LOCAL oExcellCell, cText
		oExcellCell= NEWOBJECT('Excell_Cell','class_excell_cell.prg')
		oExcellCell.Value = .NULL.
		oExcellCell.FoxproFieldType = 'D'
		cText = oExcellCell.GetFieldValueString()
		THIS.assertEqual(cText, '0')
	ENDFUNC
	
	FUNCTION Test_Null_Date_ZeroDateString
	
		LOCAL oExcellCell, cText
		oExcellCell= NEWOBJECT('Excell_Cell','class_excell_cell.prg')
		oExcellCell.Value = .NULL.		
		oExcellCell.FoxproFieldType = 'D'
		oExcellCell.ZeroDateString = .T.
		cText = oExcellCell.GetFieldValueString()
		THIS.assertEqual(cText, '  -  -')
	ENDFUNC
ENDDEF