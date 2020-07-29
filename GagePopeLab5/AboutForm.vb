Option Strict On
' NETD2202 - Lab 5
' Gage Pope
' July 29, 2020
' Description: The about form for the main lab 5 form.

Public Class frmAbout
    ''' <summary>
    ''' handles clicking the ok button
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub OkClick(sender As Object, e As EventArgs) Handles btnOk.Click
        ' close the AboutForm only
        Me.Close()
    End Sub
End Class