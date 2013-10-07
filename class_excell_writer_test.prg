SET DEFAULT TO JUSTPATH(SYS(16))
CLOSE TABLES ALL

DO p_set.prg

*!*	&& d_spmind : table dengan berbagai field type
*!*			
*!*		USE C:\D_SPMIND.DBF IN 0

*!*		LOCAL oExcellWriter

*!*		oExcellWriter = NEWOBJECT('Excell_Writer','class_excell_writer.prg')
*!*		
*!*		oExcellWriter.ZeroDateString = .T.
*!*		oExcellWriter.SetCursor('D_SPMIND')
*!*		oExcellWriter.SetFileOutputPath('C:\Book1.xlsx')
*!*		oExcellWriter.Convert()

*!*	&& d_spmind : table dengan berbagai field type

*!*		USE C:\D_SPMIND.DBF IN 0

*!*		LOCAL oExcellWriter

*!*		oExcellWriter = NEWOBJECT('Excell_Writer','class_excell_writer.prg')
*!*		
*!*		&& dengan custom header
*!*		oExcellWriter.Headers.Add('thang')
*!*		oExcellWriter.Headers.Add('kdsatker')
*!*		oExcellWriter.Headers.Add('nospm')
*!*		
*!*		&& menggunakan "  -  -" untuk tanggal kosong
*!*		oExcellWriter.ZeroDateString = .T.
*!*		
*!*		oExcellWriter.SetCursor('D_SPMIND')
*!*		oExcellWriter.SetFileOutputPath('C:\Book1.xlsx')
*!*		oExcellWriter.Convert()

&& d_spmind : tanpa nama field

	&& CREATE CURSOR employee ( ;
		&& EmpID N(5), Name Character(20), Address C(30), City C(30), ;
		&& PostalCode C(10), OfficeNo C(8) NULL, Specialty Memo;
		&& )
	
	LOCAL oExcellWriter
	oExcellWriter = NEWOBJECT('Excell_Writer','class_excell_writer.prg')
	
	CREATE CURSOR employee (Specialty CHR(254))
	APPEND BLANK
	mSpecialty = FILETOSTR('alt-codes-utf-8.txt')
	REPLACE Specialty WITH mSpecialty
	
	oExcellWriter.SetCursor('employee')
	oExcellWriter.SetFileOutputPath('C:\Book1.xlsx')
	oExcellWriter.Convert()
	
	&& LOCAL oExcellWriter

	&& oExcellWriter = NEWOBJECT('Excell_Writer','class_excell_writer.prg')
	
	&& oExcellWriter.ExportFieldNames = .F.
	
	&& oExcellWriter.ZeroDateString = .T.
	
	&& oExcellWriter.SetCursor('D_SPMIND')
	&& oExcellWriter.SetFileOutputPath('C:\Book1.xlsx')
	&& oExcellWriter.Convert()

