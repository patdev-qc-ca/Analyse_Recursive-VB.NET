Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class Form1
	Inherits System.Windows.Forms.Form
	Dim colInPaths As Collection
	Dim colOutpaths As Collection
	Dim sInputPath As String
	Dim sOutputPath As String
	Dim sInputPath2 As String
	Dim sOutputPath2 As String
	Dim lTotalProcess As Integer
	Dim SearchPath, FindStr As String
	Dim FileSize As Integer
	Dim NumFiles, NumDirs As Short
	Dim AppStringName As String
	Dim cTempCollection As Collection
	
	'UPGRADE_WARNING: Structure WIN32_FIND_DATA may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function FindFirstFile Lib "Kernel32"  Alias "FindFirstFileA"(ByVal lpFileName As String, ByRef lpFindFileData As WIN32_FIND_DATA) As Integer
	'UPGRADE_WARNING: Structure WIN32_FIND_DATA may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function FindNextFile Lib "Kernel32"  Alias "FindNextFileA"(ByVal hFindFile As Integer, ByRef lpFindFileData As WIN32_FIND_DATA) As Integer
	Private Declare Function GetFileAttributes Lib "Kernel32"  Alias "GetFileAttributesA"(ByVal lpFileName As String) As Integer
	Private Declare Function FindClose Lib "Kernel32" (ByVal hFindFile As Integer) As Integer
	Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Integer)
	Private Declare Function lstrcat Lib "Kernel32"  Alias "lstrcatA"(ByVal lpString1 As String, ByVal lpString2 As String) As Integer
	'UPGRADE_WARNING: Structure BrowseInfo may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function SHBrowseForFolder Lib "shell32" (ByRef lpbi As BrowseInfo) As Integer
	Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Integer, ByVal lpBuffer As String) As Integer
	
	Const MAX_PATH As Short = 260
	Const MAXDWORD As Integer = &HFFFF
	Const INVALID_HANDLE_VALUE As Short = -1
	Const FILE_ATTRIBUTE_ARCHIVE As Integer = &H20
	Const FILE_ATTRIBUTE_DIRECTORY As Integer = &H10
	Const FILE_ATTRIBUTE_HIDDEN As Integer = &H2
	Const FILE_ATTRIBUTE_NORMAL As Integer = &H80
	Const FILE_ATTRIBUTE_READONLY As Integer = &H1
	Const FILE_ATTRIBUTE_SYSTEM As Integer = &H4
	Const FILE_ATTRIBUTE_TEMPORARY As Integer = &H100
	Const BIF_RETURNONLYFSDIRS As Short = 1
	
	Private Structure BrowseInfo
		Dim lngHwnd As Integer
		Dim pIDLRoot As Integer
		Dim pszDisplayName As Integer
		Dim lpszTitle As Integer
		Dim ulFlags As Integer
		Dim lpfnCallback As Integer
		Dim lParam As Integer
		Dim iImage As Integer
	End Structure
	Private Structure FILETIME
		Dim dwLowDateTime As Integer
		Dim dwHighDateTime As Integer
	End Structure
	Enum ListPaths
		PathsAndFilenames = 1
		FilenamesOnly = 2
		PathsOnly = 3
	End Enum
	Dim ListSelected As ListPaths
	Private Structure WIN32_FIND_DATA
		Dim dwFileAttributes As Integer
		Dim ftCreationTime As FILETIME
		Dim ftLastAccessTime As FILETIME
		Dim ftLastWriteTime As FILETIME
		Dim nFileSizeHigh As Integer
		Dim nFileSizeLow As Integer
		Dim dwReserved0 As Integer
		Dim dwReserved1 As Integer
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(MAX_PATH),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=MAX_PATH)> Public cFileName() As Char
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(14),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=14)> Public cAlternate() As Char
	End Structure
	
	Dim dWork As String
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		'UPGRADE_ISSUE: CommandButton property Command1.Value was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        dWork = BrowseForFolder(0, CStr(Command1.Text))
		Dim cfList As New Collection
		' Set cfList =
		DoFileSystemSearch(dWork, "*.*", ListPaths.PathsAndFilenames)
	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		Dim x As Object
		FileOpen(1, dWork & "\dossiers.inf", OpenMode.Output)
		PrintLine(1, "[Folder]")
		For x = 0 To dossiers.Items.Count - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(1, x + 1 & "=" & VB.Right(VB6.GetItemString(dossiers, x), Len(VB6.GetItemString(dossiers, x)) - Len(dWork) - 1))
		Next 
		FileClose(1)
		FileOpen(1, dWork & "\fichiers.inf", OpenMode.Output)
		PrintLine(1, "[Files]")
		For x = 0 To List1.Items.Count - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PrintLine(1, x + 1 & "=" & VB.Right(VB6.GetItemString(List1, x), Len(VB6.GetItemString(List1, x)) - Len(dWork) - 1))
		Next 
		FileClose(1)
		FileOpen(1, dWork & "\RCDatas.cpp", OpenMode.Output)
		PrintLine(1, "CreateArboretum(char* Path){")
		For x = 0 To dossiers.Items.Count - 1
			PrintLine(1, "CreateDirectory(" & Chr(34) & "%s\" & VB.Right(VB6.GetItemString(dossiers, x), Len(VB6.GetItemString(dossiers, x)) - Len(dWork) - 1) & Chr(34) & ",Path);")
		Next 
		PrintLine(1, "}")
		FileClose(1)
		
	End Sub
	
	Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
		End
	End Sub
	
	Private Function BrowseForFolder(ByVal lngHwnd As Integer, ByVal strPrompt As String) As String
		
		On Error GoTo ehBrowseForFolder 'Trap for errors
		
		Dim intNull As Short
		Dim lngIDList, lngResult As Integer
		Dim strPath As String
		Dim udtBI As BrowseInfo
		
		'Set API properties (housed in a UDT)
		With udtBI
			.lngHwnd = lngHwnd
			.lpszTitle = lstrcat(strPrompt, "")
			.ulFlags = BIF_RETURNONLYFSDIRS
		End With
		
		'Display the browse folder...
		lngIDList = SHBrowseForFolder(udtBI)
		
		If lngIDList <> 0 Then
			'Create string of nulls so it will fill in with the path
			strPath = New String(Chr(0), MAX_PATH)
			
			'Retrieves the path selected, places in the null
			'character filled string
			lngResult = SHGetPathFromIDList(lngIDList, strPath)
			
			'Frees memory
			Call CoTaskMemFree(lngIDList)
			
			'Find the first instance of a null character,
			'so we can get just the path
			intNull = InStr(strPath, vbNullChar)
			'Greater than 0 means the path exists...
			If intNull > 0 Then
				'Set the value
				strPath = VB.Left(strPath, intNull - 1)
			End If
		End If
		
		'Return the path name
		BrowseForFolder = strPath
		Exit Function 'Abort
		
ehBrowseForFolder: 
		
		'Return no value
		BrowseForFolder = CStr(Nothing)
		
	End Function
	Function StripNulls(ByRef OriginalStr As String) As String
		If (InStr(OriginalStr, Chr(0)) > 0) Then
			OriginalStr = VB.Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
		End If
		StripNulls = OriginalStr
	End Function
	
	Function FindFilesAPI(ByVal Path As String, ByVal SearchStr As String, ByVal FileCount As Short, ByVal DirCount As Short) As Object
		
		
		Dim FileName As String ' Walking filename variable...
		Dim DirName As String ' SubDirectory Name
		Dim dirNames() As String ' Buffer for directory name entries
		Dim nDir As Short ' Number of directories in this path
		Dim i As Short ' For-loop counter...
		Dim hSearch As Integer ' Search Handle
		Dim WFD As WIN32_FIND_DATA
		Dim Cont As Short
		If VB.Right(Path, 1) <> "\" Then Path = Path & "\"
		' Search for subdirectories.
		nDir = 0
		ReDim dirNames(nDir)
		Cont = True
		hSearch = FindFirstFile(Path & "*", WFD)
		If hSearch <> INVALID_HANDLE_VALUE Then
			Do While Cont
				DirName = StripNulls(WFD.cFileName)
				' Ignore the current and encompassing directories.
				If (DirName <> ".") And (DirName <> "..") Then
					' Check for directory with bitwise comparison.
					If GetFileAttributes(Path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
						If InStr(1, Path & DirName, "Processed") = 0 Then
							dirNames(nDir) = DirName
							dossiers.Items.Add(Path & DirName)
							DirCount = DirCount + 1
							nDir = nDir + 1
							ReDim Preserve dirNames(nDir)
						End If
					End If
				End If
				Cont = FindNextFile(hSearch, WFD) 'Get next subdirectory.
			Loop 
			Cont = FindClose(hSearch)
		End If
		' Walk through this directory and sum file sizes.
		hSearch = FindFirstFile(Path & SearchStr, WFD)
		Cont = True
		If hSearch <> INVALID_HANDLE_VALUE Then
			While Cont
				FileName = StripNulls(WFD.cFileName)
				If (FileName <> ".") And (FileName <> "..") Then
					'UPGRADE_WARNING: Couldn't resolve default property of object FindFilesAPI. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
					FileCount = FileCount + 1
					
					If InStr(1, Path & FileName, "SYSTEM FILES") <> 0 Then
						
						'// SYSTEM FILES DIRECTORY
						
					Else
						
						'// OTHER DIRECTORIES
						
						cTempCollection.Add(Path & FileName)
						List1.Items.Add(Path & FileName)
						
					End If
					
				End If
				Cont = FindNextFile(hSearch, WFD) ' Get next file
			End While
			Cont = FindClose(hSearch)
		End If
		' If there are sub-directories...
		If nDir > 0 Then
			' Recursively walk into them...
			For i = 0 To nDir - 1
				'UPGRADE_WARNING: Couldn't resolve default property of object FindFilesAPI(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object FindFilesAPI. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FindFilesAPI = FindFilesAPI + FindFilesAPI(Path & dirNames(i) & "\", SearchStr, FileCount, DirCount)
			Next i
		End If
	End Function
	
	Private Function DoFileSystemSearch(ByVal sPath As String, ByVal sFilter As String, ByVal ListAction As ListPaths) As Collection
		
		ListSelected = ListAction
		
		cTempCollection = New Collection
		
		FindFilesAPI(sPath, sFilter, NumFiles, NumDirs)
		
		DoFileSystemSearch = cTempCollection
		MsgBox(cTempCollection.Count(), MsgBoxStyle.Information, "Collection")
		
	End Function
End Class