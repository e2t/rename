Attribute VB_Name = "Main"
Option Explicit

Dim swApp As Object
Dim CurrentDoc As ModelDoc2
Dim gFso As FileSystemObject
Dim PreviousOpenDocs As Collection 'of ModelDoc2

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub Main()

  Dim Doc As ModelDoc2
  Dim OldName As String
  Dim NewName As String
  Dim IsRenameDone As Boolean
  Dim IsOtherDoc As Boolean
  Dim Drawings As Collection
  Dim I As Variant
  Dim RelevantDrawing As String
  
  Set swApp = Application.SldWorks
  Set CurrentDoc = swApp.ActiveDoc
  If CurrentDoc Is Nothing Then Exit Sub
  
  Set gFso = New FileSystemObject
  Set PreviousOpenDocs = New Collection
  
  For Each I In swApp.Frame.ModelWindows
    PreviousOpenDocs.Add I.ModelDoc
  Next
  
  Set Doc = GetDocument
  IsOtherDoc = Not Doc Is CurrentDoc
  OldName = Doc.GetPathName
  NewName = GetNewName(OldName)
  If NewName = "" Then Exit Sub
  
  If IsOtherDoc Then
    If Not ActivateDoc(OldName) Then
      MsgBox "Не удалось открыть модель.", vbCritical
      Exit Sub
    End If
  End If
  
  If Doc.GetType <> swDocDRAWING Then
    Set Drawings = GetReferencedDrawings(Doc)
  End If
  IsRenameDone = RenameDoc(OldName, NewName, Doc)
  If IsOtherDoc Then
    For Each I In PreviousOpenDocs
      If I Is Doc Then
        GoTo ReturnToOriginalDoc
      End If
    Next
      swApp.CloseDoc Doc.GetPathName
  End If
  
ReturnToOriginalDoc:

  ActivateDoc CurrentDoc.GetPathName
  If IsOtherDoc Then
    CurrentDoc.ForceRebuild3 (False)
    Sleep 500
  End If
  RemoveFile OldName
  
  If Doc.GetType <> swDocDRAWING Then
    ReplaceLinksInDrawings Drawings, OldName, NewName
    RelevantDrawing = GetDrawingName(OldName)
    If SearchRelevantDrawing(RelevantDrawing, Drawings) Then
      RenameRelevantDrawing RelevantDrawing, GetDrawingName(NewName)
    End If
  End If
    
End Sub

Function GetDocument() As ModelDoc2

  Dim ThisDoc As ModelDoc2
  Dim Selected As Component2
  
  Set ThisDoc = swApp.ActiveDoc
  If ThisDoc.GetType = swDocASSEMBLY Then
    Set Selected = ThisDoc.SelectionManager.GetSelectedObjectsComponent4(1, -1)
    If Not Selected Is Nothing Then
      Set GetDocument = Selected.GetModelDoc2
      Exit Function
    End If
  End If
  Set GetDocument = ThisDoc
    
End Function

Function GetNewName(OldName As String) As String
  
  Dim BaseOldName As String
  Dim BaseNewName As String
  
  BaseOldName = gFso.GetBaseName(OldName)  'without extension
  BaseNewName = InputBox("Input new file name:", , BaseOldName)
  GetNewName = IIf(BaseNewName <> "" And BaseNewName <> BaseOldName, _
                   gFso.GetParentFolderName(OldName) & "\" & _
                   BaseNewName & "." & gFso.GetExtensionName(OldName), _
                   "")
                     
End Function

Function GetDrawingName(ModelName As String) As String

  GetDrawingName = gFso.BuildPath( _
    gFso.GetParentFolderName(ModelName), gFso.GetBaseName(ModelName) + ".SLDDRW")

End Function

Function RenameDoc(OldName As String, NewName As String, Doc As ModelDoc2) As Boolean

  Dim Error As Long
  Dim Warning As Long
  
  RenameDoc = Doc.Extension.SaveAs( _
    NewName, swSaveAsCurrentVersion, swSaveAsOptions_AvoidRebuildOnSave, _
    Nothing, Error, Warning)
                                     
End Function

Function SearchRelevantDrawing(DrawingName As String, Drawings As Collection) As Boolean

  Dim I As Variant
  
  DrawingName = LCase(DrawingName)
  SearchRelevantDrawing = False
  For Each I In Drawings
    If LCase(I) = DrawingName Then
      SearchRelevantDrawing = True
      Exit Function
    End If
  Next

End Function

Sub RenameRelevantDrawing(OldName As String, NewName As String)

  Dim I As Variant
  Dim Doc As ModelDoc2
  
  OldName = LCase(OldName)
  For Each I In PreviousOpenDocs
    Set Doc = I
    If LCase(Doc.GetPathName) = OldName Then
      RenameDoc OldName, NewName, Doc
      RemoveFile OldName
      Exit Sub
    End If
  Next
  
  On Error Resume Next
  Name OldName As NewName

End Sub

Sub RemoveFile(FileName As String)

  On Error GoTo Warning
  Kill FileName
  Exit Sub
  
Warning:
  MsgBox "Cannot to remove file " + FileName, vbCritical
    
End Sub

Function ActivateDoc(Name As String) As Boolean

  Dim Err As swActivateDocError_e
  
  swApp.ActivateDoc3 Name, True, swDontRebuildActiveDoc, Err
  ActivateDoc = (Err <> swGenericActivateError)
    
End Function

Function GetReferencedDrawings(Doc As ModelDoc2) As Collection

  Dim OldStateUseFolderSearchRules As Boolean
  Dim PackGo As PackAndGo
  Dim Refs As Variant
  Dim I As Variant
  Dim Drawings As Collection
  
  OldStateUseFolderSearchRules = swApp.GetUserPreferenceToggle(swUseFolderSearchRules)
  swApp.SetUserPreferenceToggle swUseFolderSearchRules, False
  
  Set PackGo = Doc.Extension.GetPackAndGo
  PackGo.IncludeDrawings = True
  PackGo.GetDocumentNames Refs
  Set Drawings = New Collection
  For Each I In Refs
    If LCase(I) Like "*.slddrw" Then
      Drawings.Add I
    End If
  Next
  
  swApp.SetUserPreferenceToggle swUseFolderSearchRules, OldStateUseFolderSearchRules
  Set GetReferencedDrawings = Drawings
    
End Function

Sub ReplaceLinksInDrawings(Drawings As Collection, OldName As String, NewName As String)

  Dim I As Variant
  
  For Each I In Drawings
    swApp.ReplaceReferencedDocument I, OldName, NewName
  Next
    
End Sub
