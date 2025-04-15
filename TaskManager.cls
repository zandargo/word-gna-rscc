Option Explicit

Private Sub UserForm_Initialize()
	txtNewTask.Text = ""
End Sub

Private Sub UserForm_Activate()
	txtNewTask.SetFocus
	LoadTasksFromRSCCs
End Sub

Private Sub UserForm_Deactivate()

End Sub

Private Sub TaskManager_Close()
	' Code to execute when TaskManager is closed
End Sub

Private Sub btnAddTask_Click()
	Dim newTask As String
	newTask = Trim(txtNewTask.Text)
	If newTask <> "" Then
		Dim i As Integer
		Dim isDuplicate As Boolean
		isDuplicate = False
		For i = 0 To lstTasks.ListCount - 1
			If StrComp(lstTasks.List(i), newTask, vbTextCompare) = 0 Then
				isDuplicate = True
				Exit For
			End If
		Next i
		If isDuplicate Then
			MsgBox "Task already exists in the list.", vbExclamation
		Else
			lstTasks.AddItem newTask
		End If
	End If
	DoEvents
	txtNewTask.SetFocus
End Sub

Private Sub btnDelete_Click()
	If lstTasks.ListIndex <> - 1 Then
		lstTasks.RemoveItem lstTasks.ListIndex
	End If
End Sub

Private Sub btnMoveUp_Click()
	Dim index As Integer
	index = lstTasks.ListIndex
	If index > 0 Then
		Dim temp As String
		temp = lstTasks.List(index - 1)
		lstTasks.List(index - 1) = lstTasks.List(index)
		lstTasks.List(index) = temp
		lstTasks.ListIndex = index - 1
	End If
End Sub

Private Sub btnMoveDown_Click()
	Dim index As Integer
	index = lstTasks.ListIndex
	If index < lstTasks.ListCount - 1 Then
		Dim temp As String
		temp = lstTasks.List(index + 1)
		lstTasks.List(index + 1) = lstTasks.List(index)
		lstTasks.List(index) = temp
		lstTasks.ListIndex = index + 1
	End If
End Sub

Private Sub txtNewTask_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
	If KeyCode = vbKeyReturn Then
		btnAddTask_Click
		DoEvents
		txtNewTask.SetFocus
	End If
End Sub

Private Sub btnEdit_Click()
	If lstTasks.ListIndex <> - 1 Then
		Dim currentTask As String
		currentTask = lstTasks.List(lstTasks.ListIndex)
		txtNewTask.Text = currentTask
		lstTasks.RemoveItem lstTasks.ListIndex
		txtNewTask.SetFocus
	End If
End Sub

Private Sub btnSort_Click()
	Dim i As Long, j As Long
	Dim temp As String
	
	' Bubble sort implementation for the listbox items
	For i = 0 To lstTasks.ListCount - 2
		For j = i + 1 To lstTasks.ListCount - 1
			If StrComp(lstTasks.List(i), lstTasks.List(j), vbTextCompare) > 0 Then
				' Swap items
				temp = lstTasks.List(i)
				lstTasks.List(i) = lstTasks.List(j)
				lstTasks.List(j) = temp
			End If
		Next j
	Next i
End Sub

Public Sub LoadTasksFromRSCCs()
	Dim tasks As Collection
	Dim task As Variant

	' Clear the current list
	lstTasks.Clear

	' Get tasks collection from the reusable function
	Set tasks = GetTasksFromRSCCs()

	' Add all tasks to the listbox
	For Each task In tasks
		lstTasks.AddItem task
	Next task
End Sub

Private Sub btnUpdate_Click()
	Dim cc As ContentControl
	Dim i As Long
	
	' Iterate through all content controls in the document
	For Each cc In ActiveDocument.ContentControls
		' Check if it is a dropdown/combobox and if the tag starts with "task_name"
		If (cc.Type = wdContentControlDropdownList Or cc.Type = wdContentControlComboBox) And Left(cc.Tag, 9) = "task_name" Then
			' Clear existing items
			cc.DropdownListEntries.Clear
			
			' Add all items from lstTasks
			For i = 0 To lstTasks.ListCount - 1
				cc.DropdownListEntries.Add Text : = lstTasks.List(i)
			Next i
		End If
	Next cc
	
	MsgBox "All task dropdowns have been updated.", vbInformation
End Sub