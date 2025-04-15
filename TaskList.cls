Option Explicit

' Form-level variables
Private labelCount As Long
Private defaultLabelCaption As String
Private defaultLabelWidth As Long
Private defaultLabelHeight As Long
Private Const LABEL_MARGIN As Long = 10
Private Const CONTROL_VERTICAL_OFFSET As Long = 24 ' Vertical spacing between controls

Private Sub UserForm_Initialize()
	' Initialize form-level variables
	labelCount = 0
	defaultLabelCaption = "New Label"
	defaultLabelWidth = 100
	defaultLabelHeight = 20
	
	' Initialize any controls on the form
	InitializeFormControls
End Sub

Private Sub InitializeFormControls()
	' Get the task list from TaskForm1 and populate TaskBox, excluding items that start with '(job)'
	PopulateTaskBox
End Sub

Private Sub PopulateTaskBox()
	' Get a reference to the TaskBox combobox
	Dim taskBox As MSForms.ComboBox
	Set taskBox = Me.Controls("TaskBox")
	
	' Clear existing items in the combobox
	taskBox.Clear
	
	' Get a reference to TaskForm1 to access its lstTasks
	Dim taskForm As Object
	On Error Resume Next
	Set taskForm = UserForms("TaskForm1")
	
	If Not taskForm Is Nothing Then
		Dim i As Long
		Dim taskItem As String
		
		' Loop through all items in lstTasks
		For i = 0 To taskForm.lstTasks.ListCount - 1
			taskItem = taskForm.lstTasks.List(i)
			
			' Only add items that don't start with "(job)"
			If Left(taskItem, 5) <> "(job)" Then
				taskBox.AddItem taskItem
			End If
		Next i
	End If
	
	On Error GoTo 0
End Sub

Private Sub btnAddTask_Click()
	' Increment the label counter
	labelCount = labelCount + 1

	' Create a new set of controls (TaskBox and Delete button)
	CreateTaskControls labelCount
End Sub

Private Sub CreateTaskControls(ByVal controlIndex As Long)
	' Get the size and position of the original TaskBox combobox
	Dim origTaskBox As MSForms.ComboBox
	Set origTaskBox = Me.Controls("TaskBox")
	
	Dim taskBoxLeft As Single
	Dim taskBoxTop As Single
	Dim taskBoxWidth As Single
	Dim taskBoxHeight As Single
	
	taskBoxLeft = origTaskBox.Left
	taskBoxTop = origTaskBox.Top
	taskBoxWidth = origTaskBox.Width
	taskBoxHeight = origTaskBox.Height
	
	' Get the size and position of the original btnDelete button
	Dim origDeleteBtn As MSForms.CommandButton
	Set origDeleteBtn = Me.Controls("btnDelete")
	
	Dim deleteBtnLeft As Single
	Dim deleteBtnTop As Single
	Dim deleteBtnWidth As Single
	Dim deleteBtnHeight As Single
	
	deleteBtnLeft = origDeleteBtn.Left
	deleteBtnTop = origDeleteBtn.Top
	deleteBtnWidth = origDeleteBtn.Width
	deleteBtnHeight = origDeleteBtn.Height
	
	' Calculate the vertical positions for the new controls
	' Position is determined by:
	' - Original control's top position
	' - Plus vertical offset times the control index
	Dim newTaskBoxTop As Single
	Dim newDeleteBtnTop As Single
	
	newTaskBoxTop = taskBoxTop + (CONTROL_VERTICAL_OFFSET * controlIndex)
	newDeleteBtnTop = deleteBtnTop + (CONTROL_VERTICAL_OFFSET * controlIndex)
	
	' Create a new TaskBox combobox
	Dim newTaskBox As MSForms.ComboBox
	Set newTaskBox = Me.Controls.Add("Forms.ComboBox.1", "TaskBox" & controlIndex, True)
	
	' Set the properties of the new TaskBox
	With newTaskBox
		.Left = taskBoxLeft
		.Top = newTaskBoxTop
		.Width = taskBoxWidth
		.Height = taskBoxHeight
		
		' Copy list items from the original TaskBox
		Dim i As Long
		For i = 0 To origTaskBox.ListCount - 1
			.AddItem origTaskBox.List(i)
		Next i
	End With
	
	' Create a new Delete button
	Dim newDeleteBtn As MSForms.CommandButton
	Set newDeleteBtn = Me.Controls.Add("Forms.CommandButton.1", "btnDelete" & controlIndex, True)
	
	' Set the properties of the new Delete button
	With newDeleteBtn
		.Left = deleteBtnLeft
		.Top = newDeleteBtnTop
		.Width = deleteBtnWidth
		.Height = deleteBtnHeight
		.Caption = origDeleteBtn.Caption
		
		' Add event handler for the new delete button if needed
		' This would require additional code elsewhere
	End With
End Sub