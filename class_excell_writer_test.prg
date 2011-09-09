SET DEFAULT TO JUSTPATH(SYS(16))
DO p_set.prg

&& t_satker : table dengan field type character saja
		
	&& USE E:\T_SATKER.DBF IN 0

	&& LOCAL oExcellWriter

	&& oExcellWriter = NEWOBJECT('Excell_Writer','class_excell_writer.prg')

	&& oExcellWriter.SetCursor('T_SATKER')
	&& oExcellWriter.SetFileOutputPath('E:\Book1.xlsx')
	&& oExcellWriter.Convert()

&& d_spmind : table dengan berbagai field type
		
	USE E:\D_SPMIND.DBF IN 0

	LOCAL oExcellWriter

	oExcellWriter = NEWOBJECT('Excell_Writer','class_excell_writer.prg')

	oExcellWriter.SetCursor('D_SPMIND')
	oExcellWriter.SetFileOutputPath('E:\Book1.xlsx')
	oExcellWriter.Convert()

