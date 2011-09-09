
DEFINE CLASS Scripting_File As Custom

	FilePath = ''
	
	FUNCTION Open
		LPARAMETERS lcFilePath
		This.FilePath = lcFilePath
	ENDFUNC
	
	FUNCTION Copy
		LPARAMETERS lcFileOutputPath
		LOCAL lcFilePath, loFile
		
		lcFilePath= This.FilePath
		COPY FILE &lcFilePath TO &lcFileOutputPath
	ENDFUNC
	
	FUNCTION Delete
		LOCAL lcFilePath
		lcFilePath = This.FilePath
		DELETE FILE &lcFilePath
	ENDFUNC
	
	FUNCTION GetPath
		RETURN This.FilePath
	ENDFUNC
ENDDEFINE
