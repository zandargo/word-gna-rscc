Option Explicit

' Module-level variables
Dim isToggleSave1Pressed As Boolean
Dim toggleTaskSave1 As ContentControl
Dim toggleJobSave1 As ContentControl
Dim TaskSave1IsPressed As Boolean

Private Sub Document_ContentControlOnEnter(ByVal ContentControl As ContentControl)
	Debug.Print ""
	Debug.Print "Document_ContentControlOnEnter"
	Debug.Print "Tag: " & ContentControl.Tag
	Debug.Print "Type: " & ContentControl.Type

End Sub

Private Sub Document_ContentControlOnExit(ByVal ContentControl As ContentControl, Cancel As Boolean)
	Debug.Print ""
	Debug.Print "Document_ContentControlOnExit"
	Debug.Print "Tag:  " & ContentControl.Tag
	Debug.Print "ID :  " & ContentControl.ID
	Debug.Print "SPT:  " & ContentControl.ShowingPlaceholderText
	Debug.Print "Type: " & ContentControl.Type

	' Call test sub to print information about child content controls
	' TestPrintChildControls ContentControl

	Dim currentValue As String
	
	If ContentControl.Tag = "component_or_process" Or ContentControl.Tag = "job_objective" Then
		' HandleJobSaveChanges toggleJobSave1
	End If

End Sub

' Test sub to print information about child Rich Text Content Controls (RSCC)
Private Sub TestPrintChildControls(ByVal parentControl As ContentControl)
	Debug.Print ""
	Debug.Print "--- Child Controls Information ---"
	Debug.Print "Parent Control Tag: " & parentControl.Tag

	' Check if the control is a rich text content control that might contain other controls
	If parentControl.Type = wdContentControlRichText Then
		' Get the Range of the parent control
		Dim parentRange As Range
		Set parentRange = parentControl.Range

		' Check for content controls within the range
		Dim childCC As ContentControl
		On Error Resume Next
		For Each childCC In parentRange.ContentControls
			Debug.Print "Child Control Found:"
			Debug.Print "  - Tag: " & childCC.Tag
			Debug.Print "  - ID: " & childCC.ID
			Debug.Print "  - Title: " & childCC.Title
			Debug.Print "  - Type: " & childCC.Type
			Debug.Print "  - Text: " & Left(childCC.Range.Text, 30) & IIf(Len(childCC.Range.Text) > 30, "...", "")
		Next childCC

		' If no child controls were found
		If parentRange.ContentControls.Count = 0 Then
			Debug.Print "No child content controls found."
		Else
			Debug.Print "Total child controls: " & parentRange.ContentControls.Count
		End If
		On Error GoTo 0
	Else
		Debug.Print "This is not a rich text content control. Type: " & parentControl.Type
	End If

	Debug.Print "--- End of Child Controls Information ---"
End Sub

Private Sub Document_New()
End Sub

Private Sub Document_Open()
	Debug.Print "Document opened in " & Now

	' Initialize Word objects first
	InitializeWordObjects

	' Initialize the TaskSave1 toggle state
	If Not toggleTaskSave1 Is Nothing Then
		TaskSave1IsPressed = toggleTaskSave1.Checked
	End If
End Sub

Private Sub InitializeDocument()
End Sub

Private Sub InitializeWordObjects()
	' Initialize toggle buttons
	On Error Resume Next
	' Set toggleTaskSave1 = ActiveDocument.SelectContentControlsByTag("toggle_task_save1").Item(1)
	' Set toggleJobSave1 = ActiveDocument.SelectContentControlsByTag("toggle_job_save1").Item(1)
	On Error GoTo 0

	' Validate initialization
	' If toggleTaskSave1 Is Nothing Then
	' 	MsgBox "Aviso: Control 'toggle_task_save1' não encontrado no documento.", vbExclamation
	' End If

	' If toggleJobSave1 Is Nothing Then
	' 	MsgBox "Aviso: Control 'toggle_job_save1' não encontrado no documento.", vbExclamation
	' End If
End Sub

Private Sub toggleJobSave1_Click()
	HandleToggleButtonPress toggleJobSave1
End Sub

Private Sub btnTaskManager_Click()
	TaskForm1.Show
End Sub

Private Sub OpenForm()
	TaskForm1.Show
End Sub