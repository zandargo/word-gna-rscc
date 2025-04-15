Option Explicit

' Get the local path of the OneDrive folder
Public Function GetLocalPath() As String
	Const HKCU = &H80000001
	Dim Path As String
	Path = ThisDocument.FullName
	Dim objReg As Object, rPath As String, subKeys(), subKey
	Dim urlNamespace As String, mountPoint As String, secPart As String
	Dim fileNamePos As Long
	Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	rPath = "Software\SyncEngines\Providers\OneDrive\"
	objReg.EnumKey HKCU, rPath, subKeys
	For Each subKey In subKeys
		objReg.GetStringValue HKCU, rPath & subKey, "UrlNamespace", urlNamespace
		If InStr(Path, urlNamespace) > 0 Then
			objReg.GetStringValue HKCU, rPath & subKey, "MountPoint", mountPoint
			secPart = Replace(Mid(Path, Len(urlNamespace)), "/", "\")
			Path = mountPoint & secPart
			Do Until Dir(Path, vbDirectory) <> "" Or InStr(2, secPart, "\") = 0
				secPart = Mid(secPart, InStr(2, secPart, "\"))
				Path = mountPoint & secPart
			Loop
			Exit For
		End If
	Next
	fileNamePos = InStrRev(Path, "\")
	If fileNamePos > 0 Then
		Path = Left(Path, fileNamePos - 1)
	End If
	Debug.Print "LocalPath: " & Path
	GetLocalPath = Path
End Function

' Utility function to find a content control by its tag
Public Function FindContentControlByTag(ByVal tagToFind As String) As ContentControl
	Dim cc As ContentControl

	For Each cc In ActiveDocument.ContentControls
		If cc.Tag = tagToFind Then
			Set FindContentControlByTag = cc
			Exit Function
		End If
	Next cc

	Set FindContentControlByTag = Nothing
End Function

' Utility function to check if a content control is checked
Public Function IsCheckboxChecked(ByVal checkboxCC As ContentControl) As Boolean
	If checkboxCC Is Nothing Then
		IsCheckboxChecked = False
		Exit Function
	End If
	
	IsCheckboxChecked = (checkboxCC.Checked = True)
End Function


' Utility function to check if a string exists in a dropdown content control's list items
Public Function IsStringInDropdown(ByVal searchString As String, ByVal dropdownCC As ContentControl) As Boolean
	Dim i As Long
	Dim cleanSearchString As String
	Dim cleanEntryText As String
	
	If dropdownCC Is Nothing Then
		IsStringInDropdown = False
		Exit Function
	End If
	
	If dropdownCC.Type <> wdContentControlDropdownList Then
		IsStringInDropdown = False
		Exit Function
	End If

	' Clean up the search string
	cleanSearchString = Trim(searchString)

	For i = 1 To dropdownCC.DropdownListEntries.Count
		' Clean up the entry text
		cleanEntryText = Trim(dropdownCC.DropdownListEntries.Item(i).Text)

		' Compare cleaned strings
		If StrComp(cleanSearchString, cleanEntryText, vbTextCompare) = 0 Then
			IsStringInDropdown = True
			Exit Function
		End If
	Next i
	
	IsStringInDropdown = False
End Function

' Utility function to check if a content control is showing placeholder text or is empty
Public Function IsPlaceholderOrEmpty(ByVal contentControl As ContentControl) As Boolean
	Dim contentText As String

	If contentControl Is Nothing Then
		IsPlaceholderOrEmpty = True
		Exit Function
	End If

	contentText = Trim(contentControl.Range.Text)

	' Check if the control is showing placeholder text or is empty
	IsPlaceholderOrEmpty = (contentControl.ShowingPlaceholderText Or Len(contentText) = 0)
End Function

' Utility function to handle toggle button press

' Function to get tasks from content controls in the document
' Returns a Collection of task names found in dropdown content controls with tag starting with "task_name"
Public Function GetTasksFromRSCCs() As Collection
	Dim cc As ContentControl
	Dim taskName As String
	Dim taskList As New Collection

	' Iterate through all content controls in the document
	For Each cc In ActiveDocument.ContentControls
		' Check if it is a dropdown and if the tag starts with "task_name"
		If (cc.Type = wdContentControlDropdownList Or cc.Type = wdContentControlComboBox) And Left(cc.Tag, 9) = "task_name" Then
			' Iterate through the dropdown items
			Dim i As Long
			For i = 1 To cc.DropdownListEntries.Count
				taskName = Trim(cc.DropdownListEntries.Item(i).Text)

				' Check if the task has already been added
				On Error Resume Next
				taskList.Add taskName, taskName
				On Error GoTo 0
			Next i
		End If
	Next cc

	Set GetTasksFromRSCCs = taskList
End Function