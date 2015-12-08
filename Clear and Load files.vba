Option Compare Database
Option Explicit

Function ClearAndLoadFile(MyTableName As String, MySpecificationName As String, Optional Description As String = "file") As Integer
On Error GoTo ClearAndLoadFile_Err

Dim filePath As String
Dim choice As Integer
Dim sSql As String
Dim fd As FileDialog
Dim vrtSelectedItem As Variant

    ClearAndLoadFile = -1

    'Check the user really wants to clear the table down and start again
    choice = SlamMSG("This will clear out the [" & MyTableName & "] table and reload it. All data in it will be deleted." _
                        & vbCrLf & vbCrLf & "Continue?", vbYesNo + vbQuestion)
    
    If choice = vbYes Then
        
    'Create a FileDialog object as a File Picker dialog box.
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    'Use a With...End With block to reference the FileDialog object.
    With fd
         .AllowMultiSelect = False
         .Title = "Select " & Description
        
        If .Show = True Then

            filePath = .SelectedItems(1)
        Else
            SlamMSG ("No file selected - no changes made.")
            Exit Function
            
        End If
    End With

    'Set the object variable to Nothing.
    Set fd = Nothing
          
    'Run the delete query that clears down the table
    DoCmd.SetWarnings False
    sSql = "DELETE * FROM " & MyTableName
    DoCmd.RunSQL sSql
    
    'Load new file, from selection in File Picker dialog box
    DoCmd.TransferText acImportDelim, MySpecificationName, MyTableName, filePath, True
    
    'Update audit table to show who loaded the file, from where and when
    sSql = "INSERT INTO Filter_load_audit ([dateAdded], [filePath], [AddedBy])VALUES (Date(), '" + filePath + " [" & MyTableName & "]', GetLoggedOnUser())"
    DoCmd.RunSQL sSql
    
    DoCmd.SetWarnings True
        
    Else
        SlamMSG ("No changes were made.")
        Exit Function
    End If
    
SlamMSG "File loaded into table."

ClearAndLoadFile_Exit:
    ClearAndLoadFile = 0
    Exit Function

ClearAndLoadFile_Err:
    ClearAndLoadFile = Err.Number
    SlamMSG Error$
End Function

Sub ClearLoad_CtxUser()

    ClearAndLoadFile "ctxUser", "ctxUser Import", "ctxUser extract from BFC"
    
End Sub

Sub ClearLoad_Security()

    If ClearAndLoadFile("Security_extract", "Security_extract_import", "User Access (Treasury) extract from BFC") = 0 Then
        SplitSecurityTable
    End If
    
End Sub
