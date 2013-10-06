&& @purpose Generate Microsoft Excell 2007 from Cursor
&& @todo check tabel dengan banyak kolom (akurasi)
&& @todo check tabel dengan banyak baris (performance)
&& @todo check mapping field type dari cursor / dbf ke cell data type di excell
&&  setelah dipetakan pada d_spmind tinggal field dengan tipe data date dan datetime
&& @todo pengaturan file template (nama dan lokasi file)
&& @todo check ada / tidak file template (nama dan lokasi file)
&& @depend book1.zip template file excell, tanpa sheet1.xml
&& @depend Scripting_File, Compression

DEFINE CLASS Excell_Writer AS Custom

	TemplateFilePath = 'Book1.zip'
	
	
	FileOutputPath = ''
	FilePath = ''
	RowNumber = ''
	RowCount = 0
	FieldCount = 0
	CursorName = ''
	
	Headers = Null
	
	&& menggunakan "  -  -" untuk tanggal kosong
	&& output yang sama dengan visual foxpro export to xls
	&& opsi ini akan di teruskan ke Excell_Cell.ZeroDateString
	ZeroDateString = .F.
	
	&& Baris pertama adalah Nama Field
	ExportFieldNames = .T.
	
	&& Menampilkan Progress WINDOW
	ShowProgress = .F.
	
	&& PROCEDURE Error
		&& LPARAMETERS nError, cMethod, nLine
		&& MESSAGEBOX cMethod
	&& ENDPROC
	
	PROCEDURE Init
		This.Headers = NEWOBJECT('Collection', 'class_collection.prg')
	ENDPROC

	FUNCTION SetCursor
		LPARAMETERS lcCursorName
		This.CursorName = lcCursorName
	ENDFUNC

	FUNCTION SetFilePath
		LPARAMETERS lcFilePath
		This.FilePath = lcFilePath
	ENDFUNC
	
	FUNCTION SetFileOutputPath
		LPARAMETERS lcFilePath
		This.FileOutputPath = lcFilePath
	ENDFUNC	

	FUNCTION Convert
		LOCAL oFolderTemp, oFile, oFileSheet1
		
		oFolderTemp = NEWOBJECT('Folder_Temp','class_folder_temp.prg')
		oFile = NEWOBJECT('Scripting_File','class_scripting_file.prg')
		oFileSheet1 = NEWOBJECT('Scripting_File','class_scripting_file.prg')
	
		oFileSheet1.Open(oFolderTemp.GetPath() + '\sheet1.xml')
		This.SetFilePath(oFileSheet1.GetPath())
		
		oFile.Open('Book1.zip')
		oFile.Copy(This.FileOutputPath)
		
		This.CreateSheet1()
		
		&& nge zip
		&& @todo return nya ZipOpen, ZipClose, ZipFileRelative harus .t. semua
		SET LIBRARY TO LOCFILE('vfpcompression.fll')
		ZipOpen(This.FileOutputPath, JUSTPATH(This.FileOutputPath) + '\', .T.) 
		ZipFileRelative(oFileSheet1.GetPath(), 'xl/worksheets/')
		ZipClose() 
		SET LIBRARY TO

		oFileSheet1.Delete()
		
	ENDFUNC
	
	FUNCTION CreateSheet1
	
		LOCAL lcCursorName, lnFieldNumber, lcCellNumber, lcXML, lcFieldName
		LOCAL lcCellValue
		LOCAL loExcellCell
		
		&& Selisih Nomor Baris, berbeda sesuai dengan nilai ExportFieldNames 
		LOCAL liRowNumberDelta
		
		loExcellCell = NEWOBJECT('Excell_Cell', 'class_excell_cell.prg')
		loExcellCell.ZeroDateString = This.ZeroDateString
		
		lcCursorName = This.CursorName

		SELECT &lcCursorName
		This.RowCount = RECCOUNT()
		This.FieldCount = FCOUNT()

		This.FileCreate()
		This.FileWrite(This.GetXmlTableBegin())

		This.RowNumber = LTRIM(STR(RECNO()))
		IF This.ExportFieldNames = .T.
			This.FileWrite(This.GetRowFieldNames())
			liRowNumberDelta = 1
		ELSE
			liRowNumberDelta = 0
		ENDIF
		
		DO WHILE NOT EOF()
		
			This.RowNumber = LTRIM(STR(RECNO()+liRowNumberDelta))

			IF THIS.ShowProgress
				WAIT WINDOW 'Processing '+LTRIM(STR(RECNO()))+' ('+LTRIM(STR(RECNO()*100/RECCOUNT()))+'%)' NOWAIT
			ENDIF
			
			lcXML = This.GetXmlRowBegin()
			
			FOR lnFieldNumber = 1 TO This.FieldCount
			
				lcFieldName = FIELD(lnFieldNumber)
				lcFoxproFieldType = TYPE(lcFieldName)
				
				loExcellCell.ColumnNumber = lnFieldNumber
				loExcellCell.RowNumber = This.RowNumber
				loExcellCell.FoxproFieldType = lcFoxproFieldType
				loExcellCell.Value = &lcFieldName
				
				lcXML = lcXML + loExcellCell.ToString()
				
			ENDFOR
			lcXML = lcXML + This.GetXmlRowEnd()
			This.FileWrite(lcXML)
			SKIP
		ENDDO
		This.FileWrite(This.GetXmlTableEnd())
	ENDFUNC

	FUNCTION GetRowFieldNames
		LOCAL lcXML, lnFieldNumber
		LOCAL loExcellCell
		
		loExcellCell = NEWOBJECT('Excell_Cell', 'class_excell_cell.prg')
		loExcellCell.ZeroDateString = This.ZeroDateString
		
		lcXML = This.GetXmlRowBegin()
		
		IF This.Headers.Count = 0
			FOR lnFieldNumber = 1 TO This.FieldCount

				lcFieldName = FIELD(lnFieldNumber)
				lcFoxproFieldType = TYPE(lcFieldName)

				loExcellCell.ColumnNumber = lnFieldNumber
				loExcellCell.RowNumber = This.RowNumber
				loExcellCell.FoxproFieldType = 'C'
				loExcellCell.Value = lcFieldName
				
				lcXML = lcXML + loExcellCell.ToString()
			ENDFOR
		ELSE
			FOR lnFieldNumber = 1 TO This.Headers.Count

				lcFieldName = This.Headers.Item(lnFieldNumber)

				loExcellCell.ColumnNumber = lnFieldNumber
				loExcellCell.RowNumber = This.RowNumber
				loExcellCell.FoxproFieldType = 'C'
				loExcellCell.Value = lcFieldName
				
				lcXML = lcXML + loExcellCell.ToString()
			ENDFOR
		ENDIF
		
		lcXML = lcXML + This.GetXmlRowEnd()
		
		RETURN lcXML
	ENDFUNC

	FUNCTION FileCreate
		STRTOFILE('',This.FilePath,4)
	ENDFUNC

	FUNCTION FileWrite
		LPARAMETERS lcContent
		STRTOFILE(lcContent,This.FilePath,1)
	ENDFUNC

	FUNCTION WriteCell
		LOCAL lnFieldNumber,lcXML
		FOR lnFieldNumber = 1 TO This.FieldCount
			cellNumber = This.GetColumnNameFromNumber(lnFieldNumber)
			lcXML = '<c r="' + cellNumber + This.RowNumber ;
				+ '" s="0" t="s"><v>0</v></c>'
			STRTOFILE(lcXML,This.FilePath,1)
		ENDFOR
	ENDFUNC

	FUNCTION GetXmlRowBegin
		RETURN '<row collapsed="false" customFormat="false"' ;
					+ ' customHeight="false" hidden="false"' ;
					+ ' ht="12.8" outlineLevel="0" r="' ;
					+ This.RowNumber + '">'
	ENDFUNC

	FUNCTION GetXmlRowEnd
		RETURN '</row>'
	ENDFUNC

	FUNCTION GetXmlTableBegin
		RETURN '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' ;
				+ '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' ;
					+ ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' ;
				+ '<sheetPr filterMode="false">' ;
					+ '<pageSetUpPr fitToPage="false"/>' ;
				+ '</sheetPr>' ;
				+ '<dimension ref="A1"/>' ;
				+ '<sheetViews>' ;
					+ '<sheetView colorId="64" defaultGridColor="true" rightToLeft="false"' ;
						+ ' showFormulas="false" showGridLines="true" showOutlineSymbols="true"' ;
						+ ' showRowColHeaders="true" showZeros="true" tabSelected="true" topLeftCell="A1"' ;
						+ ' view="normal" windowProtection="false" workbookViewId="0"' ;
						+ ' zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100">' ;
						+ '<selection activeCell="A1" activeCellId="0" pane="topLeft" sqref="A1"/>' ;
					+ '</sheetView>' ;
				+ '</sheetViews>' ;
				+ '<cols>' ;
					+ '<col collapsed="false" hidden="false" max="257" min="1" style="0" width="11.6235294117647"/>' ;
				+ '</cols>' ;
				+ '<sheetData>'
	ENDFUNC

	FUNCTION GetXmlTableEnd
		RETURN '</sheetData>' ;
			+ '<printOptions headings="false" gridLines="false" gridLinesSet="true"' ;
				+ ' horizontalCentered="false" verticalCentered="false"/>' ;
			+ '<pageMargins left="0.7875" right="0.7875" top="1.05277777777778"' ;
				+ ' bottom="1.05277777777778" header="0.7875" footer="0.7875"/>' ;
			+ '<pageSetup blackAndWhite="false" cellComments="none" copies="1"' ;
				+ ' draft="false" firstPageNumber="1" fitToHeight="1" fitToWidth="1"' ;
				+ ' horizontalDpi="300" orientation="portrait" pageOrder="downThenOver"' ;
				+ ' paperSize="1" scale="100" useFirstPageNumber="true" usePrinterDefaults="false" verticalDpi="300"/>' ;
			+ '<headerFooter differentFirst="false" differentOddEven="false">' ;
				+ '<oddHeader>&amp;C&amp;&quot;Times New Roman,Normal&quot;&amp;12&amp;A</oddHeader>' ;
				+ '<oddFooter>&amp;C&amp;&quot;Times New Roman,Normal&quot;&amp;12Page &amp;P</oddFooter>' ;
			+ '</headerFooter>' ;
		+ '</worksheet>'
	ENDFUNC
ENDDEFINE
