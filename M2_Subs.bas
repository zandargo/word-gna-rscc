'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module: Module2
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub HandleToggleButtonPress(ByVal toggleButton As Object)
    If toggleButton Is Nothing Then
        MsgBox "Erro: O controle de alternância não foi inicializado corretamente.", vbExclamation
        Exit Sub
    End If

    ' Debug.Print "Toggle button pressed: " & toggleButton.Name
    ' Debug.Print "   Value: " & toggleButton.Value
    ' Debug.Print "   Color: " & toggleButton.ForeColor
    
    If toggleButton.Checked Then
        ' Add logic for pressed state here
        toggleButton.ForeColor = 49152 ' Active color
    Else
        ' Add logic for unpressed state here
        toggleButton.ForeColor = -2147483631 ' Inactive color
    End If

    If toggleButton.Name = "toggleJobSave1" Then
        Debug.Print "JobSave1IsPressed: " & toggleButton.Checked
        Debug.Print ""
        HandleJobSaveChanges toggleButton
    End If

    If toggleButton.Name = "toggleTaskSave1" Then
        Debug.Print "JobSave1IsPressed: " & toggleButton.Checked
        Debug.Print ""
    End If
    
End Sub

Public Sub HandleJobSaveChanges(ByVal toggleButton As Object)

    ' Only proceed if button is pressed (True)
    If toggleButton.Checked Then
        Dim componentCC As ContentControl
        Dim objectiveCC As ContentControl
        Dim taskCC As ContentControl
        Dim combinedValue As String

        ' Get the content controls
        Set componentCC = FindContentControlByTag("component_or_process")
        Set objectiveCC = FindContentControlByTag("job_objective")

        ' Check if we have valid controls and they're not empty
        If Not componentCC Is Nothing And Not objectiveCC Is Nothing Then
            If Not IsPlaceholderOrEmpty(componentCC) And Not IsPlaceholderOrEmpty(objectiveCC) Then
                ' Create combined value in required format
                combinedValue = "(job) " & Trim(componentCC.Range.Text) & " : " & Trim(objectiveCC.Range.Text)

                ' Iterate through all content controls
                For Each taskCC In ActiveDocument.ContentControls
                    ' Check if it's a dropdown/combobox with task_name tag
                    If (taskCC.Type = wdContentControlDropdownList Or taskCC.Type = wdContentControlComboBox) And _
                            Left(taskCC.Tag, 9) = "task_name" Then

                        ' Check if value already exists
                        If Not IsStringInDropdown(combinedValue, taskCC) Then
                            ' Add the new value if it doesn't exist
                            taskCC.DropdownListEntries.Add Text:=combinedValue
                        End If
                    End If
                Next taskCC
            End If
        End If
    End If
End Sub

Public Sub ListAllContentControlsInfo()
    Dim cc As ContentControl
    Dim pageNum As Long
    Dim fileNum As Integer
    Dim outputPath As String
    
    ' Create the output file path in the same directory as the document
    outputPath = GetLocalPath & "\ContentControlsList.txt"
    
    ' Get the next available file number
    fileNum = FreeFile
    
    ' Open the file for output
    Open outputPath For Output As fileNum
    
    Print #fileNum, "Content Controls Information:"
    Print #fileNum, "-----------------------------"
    
    For Each cc In ActiveDocument.ContentControls
        ' Get the page number where the content control is located
        pageNum = cc.Range.Information(wdActiveEndPageNumber)
        
        ' Print the information to file
        Print #fileNum, "Tag: " & IIf(cc.Tag = "", "[No Tag]", cc.Tag)
        Print #fileNum, "Type: " & GetContentControlTypeName(cc.Type)
        Print #fileNum, "Page: " & pageNum
        Print #fileNum, "Text: " & Left(cc.Range.Text, 50) & IIf(Len(cc.Range.Text) > 50, "...", "")
        Print #fileNum, "-----------------------------"
    Next cc
    
    ' Close the file
    Close #fileNum
    
    ' Inform user
    MsgBox "Content Controls list has been saved to:" & vbCrLf & outputPath, vbInformation
End Sub

Private Function GetContentControlTypeName(ByVal ccType As WdContentControlType) As String
    Select Case ccType
        Case wdContentControlRichText
            GetContentControlTypeName = "Rich Text"
        Case wdContentControlText
            GetContentControlTypeName = "Plain Text"
        Case wdContentControlComboBox
            GetContentControlTypeName = "Combo Box"
        Case wdContentControlDropdownList
            GetContentControlTypeName = "Dropdown List"
        Case wdContentControlDate
            GetContentControlTypeName = "Date Picker"
        Case wdContentControlGroup
            GetContentControlTypeName = "Group"
        Case wdContentControlCheckBox
            GetContentControlTypeName = "Check Box"
        Case Else
            GetContentControlTypeName = "Other (" & ccType & ")"
    End Select
End Function


Sub TestDuplicateContentControlByTag()
    AddRepeatingSection "capability_section"
End Sub


Sub AddRepeatingSection(ByVal controlTag As String)
    Dim cc As Word.ContentControl
    Dim repCC As Word.RepeatingSectionItem
    
    Set cc = FindContentControlByTag(controlTag)
    
    If Not cc Is Nothing Then
        Set repCC = cc.RepeatingSectionItems.Item(cc.RepeatingSectionItems.Count)
        repCC.InsertItemAfter
    Else
        MsgBox "Content control with tag '" & controlTag & "' not found.", vbExclamation
    End If
End Sub