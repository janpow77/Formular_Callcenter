Attribute VB_Name = "m_main"
Option Explicit
Sub Zwischenspeichern()
Call doc_save
End Sub
Sub Abschliesen()
Dim docpfad As String
Dim docname As String
Dim email As String
Dim sikAbfrage As Integer
Dim speichern As Boolean
Dim Backoffice As Boolean
Dim Betrug As Boolean
Dim soforthilfe_htai As Boolean
Dim soforthilfe_rpks As Boolean
Dim feld As ContentControl
Dim header As String
Dim adressat As String


  
    sikAbfrage = 0
       ActiveDocument.CustomDocumentProperties("DokumentZWS").Value = False
    docname = ""
    docname = NameBauen
    
     
     
    For Each feld In ActiveDocument.ContentControls
        If feld.Tag = "Beantwortet" Then
            speichern = feld.Checked
                   
        End If
        
        If feld.Tag = "Backoffice" Then
            Backoffice = feld.Checked
                        
        End If
        
        
        If feld.Tag = "Soforthilfe_HTAI" Then
            soforthilfe_htai = feld.Checked
                        
        End If
        
        If feld.Tag = "Soforthilfe_RPKS" Then
            soforthilfe_rpks = feld.Checked
         End If
                  
        If feld.Tag = "Betrug" Then
        Betrug = feld.Checked
        End If
    Next feld

       
       If Backoffice = True Then
              email = ActiveDocument.CustomDocumentProperties("DokumentEmail").Value
                header = "Weiterleitung an das Backoffice "
                            
            End If
       
                    
            If soforthilfe_htai = True Then
                email = ActiveDocument.CustomDocumentProperties("DokumentEmail2").Value
                header = "Allgemeine Frage zur Soforthilfe"
                adressat = "die Hessen Trade & Invest GmbH (HTAI)"
            End If
        
            If soforthilfe_rpks = True Then
            email = ActiveDocument.CustomDocumentProperties("DokumentEmail3").Value
            header = "Sachstandsanfrage zur Soforthilfe"
            adressat = "das Regierungspräsidium Kassel"
            End If
        
          
        
        If Betrug = True Then
          sikAbfrage = MsgBox("Sie haben einen Hinweis auf einen Betrugsverdacht erhalten. Der Vorgang wird gespeichert und zur Information per E-Mail weitergeleitet. ", vbOKCancel, "Abgeschlossen!")
           
            email = "lukas.boelke@wirtschaft.hessen.de;katja.kuemmel@wirtschaft.hessen.de;erin.polster@wirtschaft.hessen.de;olaf.hossfeld@wirtschaft.hessen.de"
            header = "Hinweis auf einen Betrugsverdacht im Soforthilfeprogramm"
        
             If Betrug And (sikAbfrage = 1) Then
                DocSpeichern docname
                DocEmail docname, email, header
                FelderLeeren
               FelderFüllen
               ActiveDocument.CustomDocumentProperties("DokumentZWS").Value = False
               FelderFüllen
                Exit Sub
                Else
                Exit Sub
            End If
        Exit Sub
        End If
        
        
        
     
        If speichern Then
            sikAbfrage = MsgBox("Sie haben mündlich geantwortet. Die Datei wird gespeichert und abgelegt.", vbOKCancel, "Abgeschlossen!")
        Else
            sikAbfrage = MsgBox("Sie haben nicht mündlich geantwortet, die Anfragedaten werden zur weiteren Bearbeitung per E-Mail an " & adressat & " weitergeleitet! ", vbOKCancel, header)
        End If
        
        If speichern And (sikAbfrage = 1) Then
                       
            
           DocSpeichern docname
           FelderLeeren
           FelderFüllen
           ActiveDocument.CustomDocumentProperties("DokumentZWS").Value = False
        End If
        
        If (Not speichern) And (sikAbfrage <> 2) And soforthilfe_rpks = False Then
            
           'Emails werden nicht als Dokument gespeichert!
           'DocSpeichern docname
            DocEmail docname, email, header
              FelderLeeren
              FelderFüllen
              ActiveDocument.CustomDocumentProperties("DokumentZWS").Value = False
            
            End If
        
        If (Not speichern) And (sikAbfrage <> 2) And soforthilfe_rpks = True Then
            
           'Emails werden nicht als Dokument gespeichert!
            DocEmailRPKS docname, email, exportfile
             FelderLeeren
              FelderFüllen
              ActiveDocument.CustomDocumentProperties("DokumentZWS").Value = False
             
            End If
        
        
        
        
        
        'If (sikAbfrage <> 2) Then
         '   FelderLeeren
          '  FelderFüllen
        'End If
        
        If (Not speichern) And (sikAbfrage = 2) Then
        MsgBox "Die Anfragedaten wurden weder per Email versandt noch gespeichert. Bitte treffen Sie eine Entscheidung.", , "Achtung!"
        
        Exit Sub
        End If
  
End Sub


Sub FelderLeeren()
Dim feld As ContentControl
    
    For Each feld In ActiveDocument.ContentControls
        If Not feld.Type = wdContentControlCheckBox Then
            feld.Range.text = ""
        Else
            feld.Checked = False
        End If
    Next feld

End Sub
Sub FelderFüllen()
Dim feld As ContentControl
   ' Die Felder werden nur dann gefüllt bzw. aktualisiert wenn die Datei neu erstellt wird. Die zwischengespeicherten Dateien werden nicht akualisiert.
   
If ActiveDocument.CustomDocumentProperties("DokumentZWS").Value = False Then
        
       
    For Each feld In ActiveDocument.ContentControls
        If feld.Tag = "Agent" Then
            feld.Range.text = Application.UserName
            
        End If
        If feld.Tag = "Datum" Then
            feld.Range.text = Date
        End If
        If feld.Tag = "Uhrzeit" Then
            feld.Range.text = Time
        End If
    Next feld
End If


End Sub

Sub Support()
    MsgBox "Sollten Sie Fragen haben oder weitere Informationen benötigen, kontaktieren Sie:" & Chr(13) & Chr(13) & "Jan Riener" & Chr(13) & "Telefon: 0611-815-2275" & Chr(13) & "E-Mail: jan.riener@wirtschaft.hessen.de", vbOKOnly, "Support"
End Sub


