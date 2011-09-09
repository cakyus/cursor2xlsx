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
	FieldCount = 0
	CursorName = ''
	
	&& PROCEDURE Error
		&& LPARAMETERS nError, cMethod, nLine
		&& MESSAGEBOX cMethod
	&& ENDPROC

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
		
		This.ConvertOriginal()
		
		&& nge zip
		&& @todo return nya ZipOpen, ZipClose, ZipFileRelative harus .t. semua
		SET LIBRARY TO LOCFILE('vfpcompression.fll')
		ZipOpen(This.FileOutputPath, JUSTDRIVE(This.FileOutputPath) + '\', .T.) 
		ZipFileRelative(oFileSheet1.GetPath(), 'xl/worksheets/')
		ZipClose() 
		SET LIBRARY TO

		oFileSheet1.Delete()
		
	ENDFUNC
	
	FUNCTION ConvertOriginal
		LOCAL lcCursorName, lnFieldNumber, lcCellNumber, lcXML, lcFieldName
		LOCAL lcCellValue
		LOCAL loExcellCell
		
		loExcellCell = NEWOBJECT('Excell_Cell', 'class_excell_cell.prg')
		
		lcCursorName = This.CursorName

		SELECT &lcCursorName
		This.FieldCount = FCOUNT()

		This.FileCreate()
		This.FileWrite(This.GetXmlTableBegin())

		This.RowNumber = LTRIM(STR(RECNO()))
		This.FileWrite(This.GetRowFieldNames())
		DO WHILE NOT EOF()
			This.RowNumber = LTRIM(STR(RECNO()+1))
			lcXML = This.GetXmlRowBegin()
			FOR lnFieldNumber = 1 TO This.FieldCount
			
				lcFieldName = FIELD(lnFieldNumber)
				lcFoxproFieldType = TYPE(lcFieldName)
				&& lcExcellFieldType = This.GetExcellFieldType(lcFoxproFieldType)
				&& lcCellNumber = This.GetColumnNameFromNumber(lnFieldNumber)
				&& lcCellValue = This.GetFieldValueString(lcFoxproFieldType, &lcFieldName, lcFieldName)
				
				loExcellCell.ColumnNumber = lnFieldNumber
				loExcellCell.RowNumber = This.RowNumber
				loExcellCell.FoxproFieldType = lcFoxproFieldType
				loExcellCell.Value = &lcFieldName
				
				lcXML = lcXML + loExcellCell.ToString()
				
				&& lcXML = lcXML + '<c r="' + lcCellNumber + This.RowNumber ;
					&& + '" t="' + lcExcellFieldType + '">' ;
					&& + '<v>' + lcCellValue + '</v>' ;
					&& + '</c>'
			ENDFOR
			lcXML = lcXML + This.GetXmlRowEnd()
			This.FileWrite(lcXML)
			SKIP
		ENDDO
		This.FileWrite(This.GetXmlTableEnd())
	ENDFUNC
	
	FUNCTION GetFieldValueString
		LPARAMETERS pcFieldType, paFieldValue, pcFieldName
		DO CASE
			CASE pcFieldType = 'C'
				RETURN RTRIM(paFieldValue)
			CASE pcFieldType = 'N'
				RETURN LTRIM(STR(paFieldValue))
			CASE pcFieldType = 'D'
				RETURN LTRIM(STR(paFieldValue - DATE(1900,1,1)))
			CASE pcFieldType = 'T'
			OTHERWISE
				ERROR 'Undefined foxpro field type ' + pcFieldType
		ENDCASE
		RETURN ''
	ENDFUNC

	FUNCTION GetRowFieldNames
		LOCAL lcXML, lnFieldNumber
		LOCAL loExcellCell
		
		loExcellCell = NEWOBJECT('Excell_Cell', 'class_excell_cell.prg')
		
		lcXML = This.GetXmlRowBegin()
		
		FOR lnFieldNumber = 1 TO This.FieldCount

			lcFieldName = FIELD(lnFieldNumber)
			lcFoxproFieldType = TYPE(lcFieldName)

			loExcellCell.ColumnNumber = lnFieldNumber
			loExcellCell.RowNumber = This.RowNumber
			loExcellCell.FoxproFieldType = 'C'
			loExcellCell.Value = lcFieldName
			
			lcXML = lcXML + loExcellCell.ToString()
		ENDFOR
		
		lcXML = lcXML + This.GetXmlRowEnd()
		
		RETURN lcXML
	ENDFUNC

	FUNCTION GetExcellFieldType
		LPARAMETERS lcFoxproFieldType
		LOCAL lcExcellCellDataType

&& lcFoxproFieldType
&& A Array (only returned when the optional 1 parameter is included)
&& C Character, Varchar, Varchar (Binary)
&& D Date
&& G General
&& L Logical
&& M Memo
&& N Numeric, Float, Double, or Integer
&& O Object
&& Q Blob, Varbinary
&& S Screen
&& T DateTime
&& U Undefined type of expression or cannot evaluate expression.
&& Y Currency

&& lcExcellCellDataType
&& const TYPE_STRING2		= 'str';
&& const TYPE_STRING		= 's';
&& const TYPE_FORMULA		= 'f';
&& const TYPE_NUMERIC		= 'n';
&& const TYPE_BOOL			= 'b';
&& const TYPE_NULL			= 's';
&& const TYPE_INLINE		= 'inlineStr';
&& const TYPE_ERROR		= 'e';

		DO CASE
			CASE lcFoxproFieldType = 'C'
				RETURN 'str'
			CASE lcFoxproFieldType = 'N'
				RETURN 'n'
			CASE lcFoxproFieldType = 'D'
			CASE lcFoxproFieldType = 'T'
			OTHERWISE
				ERROR 'Undefined foxpro field type ' + lcFoxproFieldType
		ENDCASE
		RETURN 'str'
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

	FUNCTION GetColumnNameFromNumber
		LPARAMETERS lnNumber
		LOCAL lnReminder, lcResult, lnLength
		
		lcResult = ''
		lnLength = 1
		
		DO WHILE .T.
		
			lnReminder = MOD(lnNumber , 26)
			lnResult = FLOOR(lnNumber / 26)

			IF lnReminder = 0
				lcResult = 'Z' + lcResult
			ELSE
				lcResult = CHR(64 + lnReminder) + lcResult
			ENDIF
			
			lnLength = lnLength + 1
			lnNumber = lnResult
			
			IF lnResult =< 26
				IF lnResult > 0
					IF lnReminder > 0
						lcResult = CHR(64 + lnResult) + lcResult
					ELSE
						IF lnResult > 1
							lcResult = CHR(64 + lnResult - 1) + lcResult
						ENDIF
					ENDIF
				ENDIF
				EXIT
			ENDIF
			
		ENDDO
		
		RETURN lcResult
	ENDFUNC
ENDDEFINE
