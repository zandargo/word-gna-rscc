Option Explicit

Private Sub UserForm_Initialize()
	txtNewTaskSaved.Text = ""
End Sub

Private Sub UserForm_Activate()
	' Set multipage control to display Page1
	MultiPage1.Value = 0 ' Index 0 represents Page1

	txtNewTaskSaved.SetFocus
	LoadTasksFromRSCCs
	
	' Populate comboTaskBoxSaved with filtered items from lstTasksSaved
	PopulateTaskBox1
End Sub

Private Sub UserForm_Deactivate()
	
End Sub
Private Sub TaskManager_Close()
	' Code to execute when TaskManager is closed
End Sub

' *// ---------------------------------------------------------- *//
' *//                           PAGE 1                           *//
' *// ---------------------------------------------------------- *//

Private Sub btnAddTaskSaved_Click()
	Dim newTask As String
	newTask = Trim(txtNewTaskSaved.Text)
	If newTask <> "" Then
		Dim i As Integer
		Dim isDuplicate As Boolean
		isDuplicate = False
		For i = 0 To lstTasksSaved.ListCount - 1
			If StrComp(lstTasksSaved.List(i), newTask, vbTextCompare) = 0 Then
				isDuplicate = True
				Exit For
			End If
		Next i
		If isDuplicate Then
			MsgBox "Task already exists in the list.", vbExclamation
		Else
			lstTasksSaved.AddItem newTask
		End If
	End If
	DoEvents
	txtNewTaskSaved.SetFocus
End Sub

Private Sub btnDeleteSaved_Click()
	If lstTasksSaved.ListIndex <> - 1 Then
		lstTasksSaved.RemoveItem lstTasksSaved.ListIndex
	End If
End Sub

Private Sub btnMoveUpSaved_Click()
	Dim index As Integer
	index = lstTasksSaved.ListIndex
	If index > 0 Then
		Dim temp As String
		temp = lstTasksSaved.List(index - 1)
		lstTasksSaved.List(index - 1) = lstTasksSaved.List(index)
		lstTasksSaved.List(index) = temp
		lstTasksSaved.ListIndex = index - 1
	End If
End Sub

Private Sub btnMoveDownSaved_Click()
	Dim index As Integer
	index = lstTasksSaved.ListIndex
	If index < lstTasksSaved.ListCount - 1 Then
		Dim temp As String
		temp = lstTasksSaved.List(index + 1)
		lstTasksSaved.List(index + 1) = lstTasksSaved.List(index)
		lstTasksSaved.List(index) = temp
		lstTasksSaved.ListIndex = index + 1
	End If
End Sub

Private Sub txtNewTaskSaved_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
	If KeyCode = vbKeyReturn Then
		btnAddTaskSaved_Click
		DoEvents
		txtNewTaskSaved.SetFocus
	End If
End Sub

Private Sub btnEditSaved_Click()
	If lstTasksSaved.ListIndex <> - 1 Then
		Dim currentTask As String
		currentTask = lstTasksSaved.List(lstTasksSaved.ListIndex)
		txtNewTaskSaved.Text = currentTask
		lstTasksSaved.RemoveItem lstTasksSaved.ListIndex
		txtNewTaskSaved.SetFocus
	End If
End Sub

Private Sub btnSortSaved_Click()
	Dim i As Long, j As Long
	Dim temp As String

	' Bubble sort implementation for the listbox items
	For i = 0 To lstTasksSaved.ListCount - 2
		For j = i + 1 To lstTasksSaved.ListCount - 1
			If StrComp(lstTasksSaved.List(i), lstTasksSaved.List(j), vbTextCompare) > 0 Then
				' Swap items
				temp = lstTasksSaved.List(i)
				lstTasksSaved.List(i) = lstTasksSaved.List(j)
				lstTasksSaved.List(j) = temp
			End If
		Next j
	Next i
End Sub

Public Sub LoadTasksFromRSCCs()
	Dim tasks As Collection
	Dim task As Variant
	
	' Clear the current list
	lstTasksSaved.Clear
	
	' Get tasks collection from the reusable function
	Set tasks = GetTasksFromRSCCs()
	
	' Add all tasks to the listbox
	For Each task In tasks
		lstTasksSaved.AddItem task
	Next task
End Sub

Private Sub btnUpdateSaved_Click()
	Dim cc As ContentControl
	Dim i As Long

	' Iterate through all content controls in the document
	For Each cc In ActiveDocument.ContentControls
		' Check if it is a dropdown/combobox and if the tag starts with "task_name"
		If (cc.Type = wdContentControlDropdownList Or cc.Type = wdContentControlComboBox) And Left(cc.Tag, 9) = "task_name" Then
			' Clear existing items
			cc.DropdownListEntries.Clear

			' Add all items from lstTasksSaved
			For i = 0 To lstTasksSaved.ListCount - 1
				cc.DropdownListEntries.Add Text : = lstTasksSaved.List(i)
			Next i
		End If
	Next cc

	MsgBox "All task dropdowns have been updated.", vbInformation
End Sub





' *// ---------------------------------------------------------- *//
' *//                           PAGE 2                           *//
' *// ---------------------------------------------------------- *//

' Populates comboTaskBoxSaved with items from lstTasksSaved, excluding those starting with '(job)'
Private Sub PopulateTaskBox1()
	' Clear existing items in comboTaskBoxSaved
	comboTaskBoxSaved.Clear
	
	Dim i As Long
	Dim taskItem As String
	
	' Loop through all items in lstTasksSaved
	For i = 0 To lstTasksSaved.ListCount - 1
		taskItem = lstTasksSaved.List(i)
		
		' Only add items that don't start with "(job)"
		If Left(taskItem, 5) <> "(job)" Then
			comboTaskBoxSaved.AddItem taskItem
		End If
	Next i
End Sub

Private Sub comboTaskBoxSaved_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
	If KeyCode = vbKeyReturn Then
		btnAddTaskDoc_Click
		DoEvents
		comboTaskBoxSaved.SetFocus
	End If
End Sub

Private Sub btnAddTaskDoc_Click()
	Dim newTask As String
	newTask = Trim(comboTaskBoxSaved.Text)
	If newTask <> "" Then
		Dim i As Integer
		Dim isDuplicate As Boolean
		isDuplicate = False
		For i = 0 To listCurrentTasks1.ListCount - 1
			If StrComp(listCurrentTasks1.List(i), newTask, vbTextCompare) = 0 Then
				isDuplicate = True
				Exit For
			End If
		Next i
		If isDuplicate Then
			MsgBox "Task already exists in the list.", vbExclamation
		Else
			listCurrentTasks1.AddItem newTask
			
			' Remove the task from comboTaskBoxSaved
			Dim comboIndex As Integer
			For i = 0 To comboTaskBoxSaved.ListCount - 1
				If StrComp(comboTaskBoxSaved.List(i), newTask, vbTextCompare) = 0 Then
					comboTaskBoxSaved.RemoveItem i
					Exit For
				End If
			Next i
			
			' Clear the input value
			comboTaskBoxSaved.Text = ""
		End If
	End If
	DoEvents
	comboTaskBoxSaved.SetFocus
End Sub

Private Sub btnDeleteDoc_Click()
	If listCurrentTasks1.ListIndex <> - 1 Then
		' Store the task before removing it
		Dim taskToDelete As String
		taskToDelete = listCurrentTasks1.List(listCurrentTasks1.ListIndex)
		
		' Remove it from the list
		listCurrentTasks1.RemoveItem listCurrentTasks1.ListIndex
		
		' First check if the task exists in lstTasksSaved
		Dim i As Integer, j As Integer
		Dim existsInSaved As Boolean
		existsInSaved = False

		For i = 0 To lstTasksSaved.ListCount - 1
			If StrComp(lstTasksSaved.List(i), taskToDelete, vbTextCompare) = 0 Then
				existsInSaved = True
				Exit For
			End If
		Next i

		' Only proceed if the task exists in lstTasksSaved
		If existsInSaved Then
			' Check if the task already exists in comboTaskBoxSaved
			Dim existsInCombo As Boolean
			existsInCombo = False
			
			For i = 0 To comboTaskBoxSaved.ListCount - 1
				If StrComp(comboTaskBoxSaved.List(i), taskToDelete, vbTextCompare) = 0 Then
					existsInCombo = True
					Exit For
				End If
			Next i
			
			' Only add the task back to comboTaskBoxSaved if it doesn't already exist there
			If Not existsInCombo Then
				comboTaskBoxSaved.AddItem taskToDelete

				' Sort the comboTaskBoxSaved options
				Dim temp As String
				For i = 0 To comboTaskBoxSaved.ListCount - 2
					For j = i + 1 To comboTaskBoxSaved.ListCount - 1
						If StrComp(comboTaskBoxSaved.List(i), comboTaskBoxSaved.List(j), vbTextCompare) > 0 Then
							temp = comboTaskBoxSaved.List(i)
							comboTaskBoxSaved.List(i) = comboTaskBoxSaved.List(j)
							comboTaskBoxSaved.List(j) = temp
						End If
					Next j
				Next i
			End If
		End If
		' If task doesn't exist in lstTasksSaved, it's simply discarded
	End If
End Sub

Private Sub btnMoveUpDoc_Click()
	Dim index As Integer
	index = listCurrentTasks1.ListIndex
	If index > 0 Then
		Dim temp As String
		temp = listCurrentTasks1.List(index - 1)
		listCurrentTasks1.List(index - 1) = listCurrentTasks1.List(index)
		listCurrentTasks1.List(index) = temp
		listCurrentTasks1.ListIndex = index - 1
	End If
End Sub

Private Sub btnMoveDownDoc_Click()
	Dim index As Integer
	index = listCurrentTasks1.ListIndex
	If index < listCurrentTasks1.ListCount - 1 Then
		Dim temp As String
		temp = listCurrentTasks1.List(index + 1)
		listCurrentTasks1.List(index + 1) = listCurrentTasks1.List(index)
		listCurrentTasks1.List(index) = temp
		listCurrentTasks1.ListIndex = index + 1
	End If
End Sub

Private Sub btnEditDoc_Click()
	If listCurrentTasks1.ListIndex <> - 1 Then
		Dim currentTask As String
		currentTask = listCurrentTasks1.List(listCurrentTasks1.ListIndex)
		comboTaskBoxSaved.Text = currentTask
		listCurrentTasks1.RemoveItem listCurrentTasks1.ListIndex
		comboTaskBoxSaved.SetFocus
	End If
End Sub

Private Sub btnSortDoc_Click()
	Dim i As Long, j As Long
	Dim temp As String

	For i = 0 To listCurrentTasks1.ListCount - 2
		For j = i + 1 To listCurrentTasks1.ListCount - 1
			If StrComp(listCurrentTasks1.List(i), listCurrentTasks1.List(j), vbTextCompare) > 0 Then
				temp = listCurrentTasks1.List(i)
				listCurrentTasks1.List(i) = listCurrentTasks1.List(j)
				listCurrentTasks1.List(j) = temp
			End If
		Next j
	Next i
End Sub

Private Sub btnUpdateDoc_Click()
	' Empty as requested
	Debug.Print "btnUpdateDoc_Click called"
End Sub