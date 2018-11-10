Attribute VB_Name = "Main"
Option Explicit

Dim swApp As Object
Dim currentDoc As ModelDoc2

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub Main()
    Dim doc As ModelDoc2
    Dim oldName As String
    Dim newName As String
    Dim resultRename As Boolean
    Dim isOtherDoc As Boolean
    
    Set swApp = Application.SldWorks
    Set currentDoc = swApp.ActiveDoc
    If currentDoc Is Nothing Then Exit Sub

    Set doc = GetDocument
    isOtherDoc = Not doc Is currentDoc
    oldName = doc.GetPathName
    newName = GetNewName(oldName)
    If newName = "" Then Exit Sub

    If isOtherDoc Then
        If Not ActivateDoc(oldName) Then
            MsgBox "Не удалось открыть модель.", vbCritical
            Exit Sub
        End If
    End If
    
    resultRename = RenameDoc(oldName, newName, doc)
    If isOtherDoc Then
        swApp.CloseDoc doc.GetPathName
    End If
    If resultRename Then
        If isOtherDoc Then
            currentDoc.ForceRebuild3 (False)
            Sleep 500
        End If
        RemoveFile oldName
    Else
        MsgBox "Cannot to save file"
    End If
End Sub
      
Function GetDocument() As ModelDoc2
    Dim thisDoc As ModelDoc2
    Dim selected As Component2

    Set thisDoc = swApp.ActiveDoc
    If thisDoc.GetType = swDocASSEMBLY Then
        Set selected = thisDoc.SelectionManager.GetSelectedObjectsComponent4(1, -1)
        If Not selected Is Nothing Then
            Set GetDocument = selected.GetModelDoc2
            Exit Function
        End If
    End If
    Set GetDocument = thisDoc
End Function

Function GetNewName(oldName As String) As String
    Dim fso As Object
    Dim baseOldName As String
    Dim baseNewName As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    baseOldName = fso.GetBaseName(oldName)  'without extension
    baseNewName = InputBox("Input new file name:", , baseOldName)
    GetNewName = IIf(baseNewName <> "" And baseNewName <> baseOldName, _
                     fso.GetParentFolderName(oldName) & "\" & _
                     baseNewName & "." & fso.GetExtensionName(oldName), _
                     "")
End Function

Function RenameDoc(oldName As String, newName As String, doc As ModelDoc2) As Boolean
    Dim error As Long, Warning As Long
    
    RenameDoc = doc.Extension.SaveAs(newName, swSaveAsCurrentVersion, swSaveAsOptions_AvoidRebuildOnSave, _
                                     Nothing, error, Warning)
End Function

Sub RemoveFile(filename As String)
    On Error GoTo Warning
    Kill filename
    Exit Sub
Warning:
    MsgBox "Cannot to remove file"
End Sub

Function ActivateDoc(name As String) As Boolean
    Dim err As swActivateDocError_e
    
    swApp.ActivateDoc3 name, True, swDontRebuildActiveDoc, err
    ActivateDoc = (err <> swGenericActivateError)
End Function
