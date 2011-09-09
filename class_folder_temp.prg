
DEFINE CLASS Folder_Temp As Custom

	PROTECTED Path

	PROCEDURE Init
		&& buat folder
		LOCAL lcFolderPath
		This.Path = SYS(5) + SYS(2003) + '\TEMP\' + SYS(2015)
		lcFolderPath = This.Path
		MKDIR &lcFolderPath
	ENDFUNC	
	
	FUNCTION GetPath
		RETURN This.Path
	ENDFUNC	
	
	FUNCTION Destroy
		&& hapus folder
		LOCAL lcFolderPath
		lcFolderPath = This.Path
		RMDIR &lcFolderPath
	ENDFUNC	
ENDDEFINE
