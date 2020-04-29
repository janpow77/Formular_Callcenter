Attribute VB_Name = "m_EMAIL"
Option Explicit
Sub DocEmail(docname As String, email As String, header As String)
Dim OutApp As Object
Dim OutMail As Object
' Anstatt des ürsprünglichen Bindings wird das late Binding eingesetzt.
' Outlook wird gleich gestartet.

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
           

    With OutMail
        .To = email
        .Subject = header & docname
        .Body = MailErstellen
        .Display
       

    End With
    

    Set OutApp = Nothing
    Set OutMail = Nothing

End Sub
Sub DocEmailRPKS(docname As String, email As String, exportfile)
Dim OutApp As Object
Dim OutMail As Object
' Anstatt des ürsprünglichen Bindings wird das late Binding eingesetzt.
' Outlook wird gleich gestartet. Es wird eine temporäre Datei erzeugt
' und als Anhang der Email beigefügt. Am Ende des Subs wird die csv Datei gelöscht.

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
           


    With OutMail
        .To = email
        .Subject = HeaderErstellenRPKS
        .Body = MailErstellenRPKS
        .Display
        .Attachments.Add exportfile

    End With
    

    Set OutApp = Nothing
    Set OutMail = Nothing

' Aufräumen
Kill exportfile

End Sub
Function HeaderErstellenRPKS()
Dim feld As ContentControl
Dim text As String
 text = "Anfrage an der Hotline des Landesregierung zu einem Antrag auf Soforthilfe. "
    
    For Each feld In ActiveDocument.ContentControls
        
            
        If feld.Tag = "RPKS_Az" Then
            text = text & "  " & "Aktenzeichen: " & feld.Range.text & ". "
        
        End If
        
        If feld.Tag = "RPKS_Thema" Then
            text = text & "Thema: " & feld.Range.text
        End If
        
              
        
    Next feld
    HeaderErstellenRPKS = text
    
End Function
Function MailErstellenRPKS()
Dim feld As ContentControl
Dim text As String
text = ""
' Hier wird der Emailtext für das RP Kassel erstellt. Das Anschreiben wird aus den Formularfeldern generiert.
  text = "Sehr geehrte Damen und Herren," & Chr(13) & " an der Hotline der Landesregierung Hessen für Fragen, Anliegen und Informationen zum Corona-Virus ist ein Anruf eingegangen, der einen Antrag auf Soforthilfe betrifft." & Chr(13) & "Zuständigkeitshalber leite ich Ihnen die Telefonnotiz der Anfrage weiter und Bitte um weitere Bearbeitung. "
text = text & "Eine Rückmeldung an die Hotline des HMWEVW ist nicht erforderlich." & Chr(13) & Chr(13) & "Mit freundlichen Grüßen" & Chr(13) & "Im Auftrag" & Chr(13)
  
    For Each feld In ActiveDocument.ContentControls
        If feld.Tag = "Agent" Then
           text = text & Chr(13) & feld.Range.text & Chr(13)
            
        End If
           
        If feld.Tag = "Agent" Then
            text = text & "---------------------" & Chr(13) & "Anruf" & Chr(13) & "Agent: " & feld.Range.text
        End If
        
        If feld.Tag = "Datum" Then
            text = text & Chr(13) & "Datum: " & feld.Range.text
        End If
        
        If feld.Tag = "Zeit" Then
            text = text & "Zeit: " & feld.Range.text
        End If
        
        If feld.Tag = "RPKS_Az" Then
            text = text & Chr(13) & "---------------------" & Chr(13) & "Aktenzeichen des Antrags auf Soforthilfe"
            text = text & Chr(13) & "Aktenzeichen: " & feld.Range.text
        End If
            
        If feld.Tag = "RPKS_Thema" Then
            text = text & Chr(13) & "---------------------" & Chr(13) & "Thema der Anfrage:"
            text = text & Chr(13) & feld.Range.text
        End If
        
        If feld.Tag = "Anliegen" Then
            text = text & Chr(13) & "---------------------" & Chr(13) & "Der Anrufer hat folgendes Anliegen / Frage"
            text = text & Chr(13) & feld.Range.text
        End If
                
        If feld.Tag = "AnruferName" Then
            text = text & Chr(13) & "---------------------" & Chr(13) & "Daten des Anrufers"
            text = text & Chr(13) & "Nachname des Anrufers: " & feld.Range.text
        End If
        
        If feld.Tag = "Unternehmen" Then
            text = text & Chr(13) & "Unternehmen: " & feld.Range.text
        End If
        
        If feld.Tag = "Unternehmensart" Then
            text = text & Chr(13) & "Unternehmensart: " & feld.Range.text
        End If
        
        If feld.Tag = "Telefon" Then
            text = text & Chr(13) & "Telefon: " & feld.Range.text
        End If
        
        If feld.Tag = "Email" Then
            text = text & Chr(13) & "E-Mail: " & feld.Range.text
        End If
        
        If feld.Tag = "Weiteres" Then
            text = text & Chr(13) & "Weiteres: " & feld.Range.text
        End If

    Next feld
    
    
    MailErstellenRPKS = text
    text = ""
    
End Function
Function MailErstellen()
Dim feld As ContentControl
Dim text As String

    For Each feld In ActiveDocument.ContentControls
        If feld.Tag = "Agent" Then
            text = text & "Anruf" & Chr(13) & "Agent: " & feld.Range.text
        End If
        If feld.Tag = "Datum" Then
            text = text & Chr(13) & "Datum: " & feld.Range.text
        End If
        If feld.Tag = "Zeit" Then
            text = text & "Zeit: " & feld.Range.text
        End If
        If feld.Tag = "AnruferName" Then
            text = text & Chr(13) & "---------------------" & Chr(13) & "Anrufer"
            text = text & Chr(13) & "Nachname des Anrufers: " & feld.Range.text
        End If
        If feld.Tag = "Unternehmen" Then
            text = text & Chr(13) & "Unternehmen: " & feld.Range.text
        End If
        If feld.Tag = "Unternehmensart" Then
            text = text & Chr(13) & "Unternehmensart: " & feld.Range.text
        End If
        If feld.Tag = "Telefon" Then
            text = text & Chr(13) & "Telefon: " & feld.Range.text
        End If
        If feld.Tag = "Email" Then
            text = text & Chr(13) & "E-Mail: " & feld.Range.text
        End If
        If feld.Tag = "Weiteres" Then
            text = text & Chr(13) & "Weiteres: " & feld.Range.text
        End If
        If feld.Tag = "Soforthilfe" Then
            text = text & Chr(13) & "---------------------" & Chr(13) & "Frage zur Soforthilfe"
            text = text & Chr(13) & "Soforthilfe: " & IIf(feld.Checked, "JA", "NEIN")
        End If
        If feld.Tag = "Anliegen" Then
            text = text & Chr(13) & "---------------------" & Chr(13) & "Anliegen / Frage"
            text = text & Chr(13) & "Anliegen: " & feld.Range.text
        End If
        If feld.Tag = "Beantwortet" Then
            text = text & Chr(13) & "---------------------" & Chr(13) & "Bearbeitung"
            text = text & Chr(13) & "Anfrage mündlich beantwortet: " & IIf(feld.Checked, "Ja", "Die Anfrage konnte während des Anrufs nicht beantwortet werden.")
        End If
        If feld.Tag = "Backoffice_Hinweise" Then
            text = text & Chr(13) & "---------------------" & Chr(13) & "Anfrage nicht beantwortet, bitte weiterleiten an:"
            text = text & Chr(13) & "Hinweis an das Backoffice: " & feld.Range.text
        End If
        
    Next feld
    MailErstellen = text
    text = ""
End Function
