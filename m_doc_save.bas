Attribute VB_Name = "m_doc_save"
Option Explicit
Sub doc_save()
Dim docpfad As String
Dim docname As String
Dim speichern As Boolean
Dim pfadVorlage As String
Dim datei As FileDialog
Dim Res As Integer
Dim dlg As Dialog
Dim backupdocpfad As String
Dim Antwort As String

 
   docpfad = ActiveDocument.CustomDocumentProperties("DokumentZWSPfad").Value
   backupdocpfad = ActiveDocument.CustomDocumentProperties("DokumentBackupPfad").Value

    ' Prüft ob das Dokument bereits zwischengespeichert wurde, wenn ja, dann wird es einfach gespeichert.
    ' Wenn es noch nicht zwischengespeichert wurde,dann wird der Name vergeben.
     
     If ActiveDocument.CustomDocumentProperties("DokumentZWS").Value = True Then
     ActiveDocument.Save
     Exit Sub
      
      Else
     ' Prüfung, ob das Verzeichnis vorhanden ist.
        If IsdocPfad(docpfad) = False Then
 
            If IsdocPfad(backupdocpfad) = False Then
                MkDir (backupdocpfad)
                docpfad = backupdocpfad
        
            Else
            docpfad = backupdocpfad
            End If
        End If
    
    'Setzt den Status des Dokument auf zwischengespeichert und baut den Namen.
    ActiveDocument.CustomDocumentProperties("DokumentZWS").Value = True
    docname = ""
    docname = NameBauen
    pfadVorlage = ActiveDocument.FullName
    
    
    Set dlg = Application.Dialogs(wdDialogFileSaveAs)
     
    With dlg
        .Name = docpfad & docname
        .Format = wdFormatXMLDocumentMacroEnabled
         If dlg.Show <> -1 Then
    ' Dialog Save as wird geöffnet, falls Abbrechen gedrückt wird, kommt ein Hinweistext und der Status wird zurückgesetzt. Die Daten bleiben erhalten.
    
         MsgBox "Die Datei wurde nicht zwischengespeichert, da Sie auf Abrechen geklickt haben. "
         ActiveDocument.CustomDocumentProperties("DokumentZWS").Value = False
         
         Exit Sub
        End If
          
    End With
        
      
     ' Speichert die Datei unter dem alten Namen ab und setzt den Status zurück.
    Set datei = Application.FileDialog(FileDialogType:=msoFileDialogSaveAs)
   
        
    ActiveDocument.CustomDocumentProperties("DokumentZWS").Value = False
    ActiveDocument.SaveAs2 FileName:=pfadVorlage
 
    End If
    
    ' Aktualisiert die Felder und setzt die Felder zurück.
    FelderLeeren
    FelderFüllen
    ActiveDocument.CustomDocumentProperties("DokumentZWS").Value = False

   Exit Sub

End Sub
Sub DocSpeichern(docname As String)
Dim pfadVorlage As String
Dim backupdocpfad As String
Dim dlg As Dialog
Dim docpfad As String

 
docpfad = ActiveDocument.CustomDocumentProperties("DokumentPfad").Value
backupdocpfad = ActiveDocument.CustomDocumentProperties("DokumentBackupPfad").Value
ActiveDocument.CustomDocumentProperties("DokumentZWS").Value = True

    If IsdocPfad(docpfad) = False Then
 
        If IsdocPfad(backupdocpfad) = False Then
        MkDir (backupdocpfad)
        docpfad = backupdocpfad
        Else
        docpfad = backupdocpfad
        End If
     End If
   
    pfadVorlage = ActiveDocument.FullName
    ActiveDocument.SaveAs2 FileName:=docpfad & docname
    ActiveDocument.SaveAs2 FileName:=pfadVorlage
    ActiveDocument.SaveAs2 FileFormat:=wdFormatXMLDocumentMacroEnabled
On Error GoTo Errorhandler

    If docpfad = backupdocpfad Then
 
    MsgBox "Die Datei wurde in einem temporären Ordner auf ihrem Rechner gespeichert. " & "Sie finden die Datei im Ordner " & backupdocpfad & "." & Chr(13) & "Bitte kontaktieren Sie:" & Chr(13) & Chr(13) & "Jan Riener" & Chr(13) & "Telefon: 0611-815-2275" & Chr(13) & "E-Mail: jan.riener@wirtschaft.hessen.de", vbOKOnly, "Die Netzlaufwerke sind nicht verfügbar!"
    
    End If

On Error GoTo Errorhandler
       
   Exit Sub


Errorhandler:
   'Falls etwas schief geht, wird der normale Speicherdialog aufgerufen
    Set dlg = Application.Dialogs(wdDialogFileSaveAs)
    With dlg
        .Name = docpfad & docname
        .Format = wdFormatXMLDocumentMacroEnabled
        .Show
    End With
    
    
End Sub
Function NameBauen()
Dim feld As ContentControl
Dim docname As String
Dim Datum As String
Dim Uhrzeit As String
   
        
    docname = ""
   
      
    For Each feld In ActiveDocument.ContentControls
        If feld.Tag = "Datum" Then
            docname = docname & Format(feld.Range.text, "yyyy_mm_dd") & "_"
        End If
        If feld.Tag = "Uhrzeit" Then
            docname = docname & "_" & Format(feld.Range.text, "hh_mm")
        End If
               
        If feld.Tag = "AnruferName" Then
            docname = docname & "_" & feld.Range.text
        End If
        If feld.Tag = "Unternehmensart" Then
            docname = docname & "_" & feld.Range.text
        End If
        If feld.Tag = "Soforthilfe" Then
            If feld.Checked Then docname = docname & "_Soforthilfe"
         End If
        
        
    Next feld
    
    
     ' Wenn das Dokument zwischengespeichert wird, kommt ein Zusatz an den Namen.
    If ActiveDocument.CustomDocumentProperties("DokumentZWS").Value = True Then
        docname = docname & "_ZWS"
    End If
    
    If ActiveDocument.CustomDocumentProperties("DokumentZWS").Value = False Then
        For Each feld In ActiveDocument.ContentControls
        
        If feld.Tag = "Beantwortet" Then
           If feld.Checked = "False" Then docname = docname & "_EMAIL"
        End If
        
    Next feld
       
       End If
            
    docname = Replace(docname, vbTab, "")
    docname = Replace(docname, "Art des Unternehmens", "")
    docname = konvert(docname)
    docname = Replace(docname, "__", "_")
    docname = docname & ".docm"

    NameBauen = docname

End Function
