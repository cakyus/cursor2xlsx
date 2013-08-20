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

	USE C:\D_SPMIND.DBF IN 0

	LOCAL oExcellWriter

	oExcellWriter = NEWOBJECT('Excell_Writer','class_excell_writer.prg')
	
	oExcellWriter.ExportFieldNames = .F.
	
	&& menggunakan "  -  -" untuk tanggal kosong
	oExcellWriter.ZeroDateString = .T.
	
	oExcellWriter.SetCursor('D_SPMIND')
	oExcellWriter.SetFileOutputPath('C:\Book1.xlsx')
	oExcellWriter.Convert()

