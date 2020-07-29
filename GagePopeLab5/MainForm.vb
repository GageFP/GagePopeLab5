Option Strict On
' NETD2202 - Lab 5
' Gage Pope
' July 29, 2020
' Description: A text editor windows form application in Visual Basic.
Imports System.IO


Public Class frmMain
#Region "Variables, Constants, and Arrays"
    Dim title As String = "Text Editor - "
    Dim fileName As String = String.Empty
    Dim filePath As String = String.Empty
    Dim fileEdited As Boolean = False
#End Region

#Region "Methods and Subroutines"
    ''' <summary>
    ''' subbroutine to save the file being worked on
    ''' </summary>
    ''' <param name="file"></param>
    Sub SaveFile(ByVal file As String)
        ' create a streamwriter variable
        Dim outputStream As StreamWriter

        ' set streamwriter variable as the file path passed to the subroutine
        outputStream = New StreamWriter(file)
        ' save text into the file
        outputStream.Write(txtBox.Text)
        ' change fileName variable to the newly saved file name without the path
        fileName = Path.GetFileName(file)
        ' close the streamwriter after use
        outputStream.Close()

        ' change the form text to reflect the filename and set the fileEdited variable to false
        Me.Text = title & fileName
        fileEdited = False
    End Sub

    ''' <summary>
    ''' subroutine for cutting and copying selected text
    ''' </summary>
    ''' <param name="selectedText"></param>
    Sub CopyCut(ByVal selectedText As String)
        ' if selected text isnt empty then copy to clipboard
        If Not selectedText = String.Empty Then
            My.Computer.Clipboard.SetText(selectedText)
        End If
    End Sub

    ''' <summary>
    ''' subroutine to open a new file
    ''' </summary>
    Sub OpenFile()
        ' create a streamreader variable
        Dim inputStream As StreamReader

        ' if the openDialog result is the ok button
        If openDialog.ShowDialog() = DialogResult.OK Then
            ' set streamreader variable to selected file
            inputStream = New StreamReader(openDialog.FileName)
            ' set fileName to file name without path and filePath to file name with path
            fileName = openDialog.SafeFileName
            filePath = openDialog.FileName
            ' put the contents of the opened file into the textbox
            txtBox.Text = inputStream.ReadToEnd()
            ' close the streamreader after use
            inputStream.Close()

            ' set form text to reflect opened file name and set fileEdited to false
            Me.Text = title & fileName
            fileEdited = False
        End If
    End Sub

    ''' <summary>
    ''' subroutine to confirm the user wants to use new/open/exit with unsaved changes to their file
    ''' </summary>
    ''' <returns></returns>
    Function ConfirmClose() As Boolean
        ' select case with message box to inform user there are unsaved changes and offers yes/no/cancel options
        Select Case MsgBox("The file has unsaved changes. Would you like to save the changes?", MsgBoxStyle.YesNoCancel)
            ' if user selects yes attempt to save the file. if the file saves then return true
            Case MsgBoxResult.Yes
                If filePath = String.Empty Then
                    If saveDialog.ShowDialog() = DialogResult.OK Then
                        filePath = saveDialog.FileName
                        SaveFile(filePath)
                        Return True
                    Else
                        ConfirmClose()
                    End If
                Else
                    SaveFile(filePath)
                    Return True
                End If
            ' if user selects no then return true
            Case MsgBoxResult.No
                Return True
            ' if user selects cancel return false
            Case MsgBoxResult.Cancel
                Return False
        End Select
    End Function ' i think this warning is because line 62 doesnt have a return with the else? i dont know but it doesnt cause any problems that i noticed

    ''' <summary>
    ''' subroutine to create a new file
    ''' </summary>
    Sub NewFile()
        ' empty the textbox, set all variables to default, reflect new file in form text
        txtBox.Text = String.Empty
        fileName = String.Empty
        filePath = String.Empty
        fileEdited = False
        Me.Text = title & "NewFile"
    End Sub
#End Region

#Region "Event Handlers"

    ''' <summary>
    ''' handles clicking the new menu item
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub NewClick(sender As Object, e As EventArgs) Handles mnuNew.Click
        ' if content is not edited then call NewFile sub
        If Not fileEdited Then
            NewFile()
            ' if content is edited then call ConfirmClose sub, if ConfirmClose returns true call NewFile sub
        ElseIf ConfirmClose() Then
            NewFile()
        End If
    End Sub

    ''' <summary>
    ''' handles clicking the open menu item
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub OpenClick(sender As Object, e As EventArgs) Handles mnuOpen.Click
        ' if content is not edited then call OpenFile sub
        If Not fileEdited Then
            OpenFile()
            ' if content is edited then call ConfirmClose sub, if ConfirmClose returns true call OpenFile sub
        ElseIf ConfirmClose() Then
            OpenFile()
        End If
    End Sub

    ''' <summary>
    ''' handles clicking the save menu item
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub SaveClick(sender As Object, e As EventArgs) Handles mnuSave.Click
        ' if filePath variable is empty then show savedialog to save as a new file
        If filePath = String.Empty Then
            If saveDialog.ShowDialog() = DialogResult.OK Then
                filePath = saveDialog.FileName
                SaveFile(filePath)
            End If
            ' if filePath variable isnt empty then save the file with the same path
        Else
            SaveFile(filePath)
        End If
    End Sub

    ''' <summary>
    ''' handles clocking the save as menu item
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub SaveAsClick(sender As Object, e As EventArgs) Handles mnuSaveAs.Click
        ' show savedialog to save as a new file
        If saveDialog.ShowDialog() = DialogResult.OK Then
            filePath = saveDialog.FileName
            SaveFile(filePath)
        End If
    End Sub

    ''' <summary>
    ''' handles clicking the exit menu item
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ExitClick(sender As Object, e As EventArgs) Handles mnuExit.Click
        ' if content isnt edited then close the application
        If Not fileEdited Then
            Application.Exit()
            ' if content is edited call the ConfirmClose sub, if ConfirmClose sub returns true then close the application
        ElseIf ConfirmClose() Then
            Application.Exit()
        End If
    End Sub

    ''' <summary>
    ''' handles clicking the about menu item
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub AboutClick(sender As Object, e As EventArgs) Handles mnuAbout.Click
        'create variable for the AboutForm and display it
        Dim aboutModal As New frmAbout
        aboutModal.ShowDialog()
    End Sub

    ''' <summary>
    ''' handles clicking the copy menu item
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CopyClick(sender As Object, e As EventArgs) Handles mnuCopy.Click
        ' call the CopyCut sub using the selected text in the textbox
        CopyCut(txtBox.SelectedText)
    End Sub

    ''' <summary>
    ''' handles clicking the cut menu item
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CutClick(sender As Object, e As EventArgs) Handles mnuCut.Click
        ' call the CopyCut sub using the selected text in the textbox then clear the selected text
        CopyCut(txtBox.SelectedText)
        txtBox.SelectedText = String.Empty
    End Sub

    ''' <summary>
    ''' handles clicking the paste menu item
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub PasteClick(sender As Object, e As EventArgs) Handles mnuPaste.Click
        ' paste the clipboard content into the textbox
        txtBox.Paste()
    End Sub

    ''' <summary>
    ''' handles the text in txtBox being changed
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub txtBoxTextChanged(sender As Object, e As EventArgs) Handles txtBox.TextChanged
        ' if content hasnt already been edited
        If Not fileEdited Then
            ' set fileEdited to true and reflect unsaved changes in the form text
            fileEdited = True
            Me.Text += " - Unsaved"
        End If
    End Sub
#End Region

End Class
