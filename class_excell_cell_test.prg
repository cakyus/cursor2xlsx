SET DEFAULT TO JUSTPATH(SYS(16))
SET SAFETY OFF
CLEAR

TRY
	TestHTMLEncode()
	TestNull()
CATCH TO oErr
	? [Error: ] + LTRIM(STR(oErr.ErrorNo)) ;
		+ [ Line: ] + LTRIM(STR(oErr.LineNo)) ;
		+ [ Procedure: ] + LTRIM(oErr.Procedure);
		+ [ Message: ] + oErr.Message ;
		+ [ UserValue: ] + oErr.UserValue
ENDTRY

FUNCTION TestHTMLEncode
	
	LOCAL oExcellCell
	LOCAL sText

	oExcellCell= NEWOBJECT('Excell_Cell','class_excell_cell.prg')
	
	sText = oExcellCell.HTMLEncode(FILETOSTR('class_excell_cell_test_001.txt'))
	STRTOFILE('<html><body>' + sText + '</body></html>', 'class_excell_cell_test_001_out.html.tmp')
	? sText
	IF sText == 'Denda 1 &#0137; untuk setiap hari keterlambatan dari nilai pekerjaan yang belum diselesaikan'
		&& DO NOTHING
	ELSE
		THROW 'Fail'
	ENDIF
	
	sText = oExcellCell.HTMLEncode('& < > "')
	? sText
	IF sText == '&amp; &lt; &gt; &quote;'
		&& DO NOTHING
	ELSE
		THROW 'Fail'
	ENDIF
	
	sText = oExcellCell.HTMLEncode("'")
	? sText
	IF sText == '&apos;'
		&& DO NOTHING
	ELSE
		THROW 'Fail'
	ENDIF
ENDFUNC

FUNCTION TestNull
	
	LOCAL oExcellCell
	LOCAL sText

	oExcellCell= NEWOBJECT('Excell_Cell','class_excell_cell.prg')
	
	oExcellCell.Value = .NULL.
	
	oExcellCell.FoxproFieldType = 'C'
	sText = oExcellCell.GetFieldValueString()
	IF sText == ''
		&& DO NOTHING
	ELSE
		THROW 'Fail'
	ENDIF
	
	oExcellCell.FoxproFieldType = 'N'
	sText = oExcellCell.GetFieldValueString()
	IF sText == '0'
		&& DO NOTHING
	ELSE
		THROW 'Fail'
	ENDIF
	
	oExcellCell.FoxproFieldType = 'D'
	sText = oExcellCell.GetFieldValueString()
	IF sText == '0'
		&& DO NOTHING
	ELSE
		THROW 'Fail'
	ENDIF
	
	oExcellCell.FoxproFieldType = 'D'
	oExcellCell.ZeroDateString = .T.
	sText = oExcellCell.GetFieldValueString()
	IF sText == '  -  -'
		&& DO NOTHING
	ELSE
		THROW 'Fail'
	ENDIF
ENDFUNC
