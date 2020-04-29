Attribute VB_Name = "m_RPKS_txt"
Option Explicit

Function exportfile()
Dim fSo As Object
Dim txtfile As Object
Dim savetxtpath As String
Dim savetxtfile As String

savetxtpath = ActiveDocument.CustomDocumentProperties("DokumentBackupPfad").Value

savetxtfile = savetxtpath & TXTNameBauen
Set fSo = CreateObject("Scripting.FileSystemObject")

Set txtfile = fSo.CreateTextFile(savetxtfile, True)

exportfile = savetxtfile

txtfile.WriteLine TXTBefüllen
txtfile.Close
   
Set fSo = Nothing

End Function
Function TXTBefüllen()
Dim feld As ContentControl
Dim text As String
text = ""

' Der Inhalt der csv Datei wird hier definiert. Zuerst die 1. Zeile als Header und im Anschluss die Werte. Übergabe der txtbefüllen an die Emailfunktion.
text = "Agent" & ";" & "Datum" & ";" & "Uhrzeit" & ";" & "AnruferName" & ";" & "Unternehmen" & ";" & "Unternehmensart" & ";" & "Telefon" & ";" & "Email" & ";" & "Weiteres" & ";" & "Anliegen" & ";" & "RPKS_Az" & ";" & "RPKS_Thema" & Chr(13) & Chr(10)
    
    
   For Each feld In ActiveDocument.ContentControls
        
        If feld.Tag = "Agent" Then
            text = text & feld.Range.text & ";"
        End If
        If feld.Tag = "Datum" Then
            text = text & Format(feld.Range.text, "dd.mm.yyyy") & ";"
        End If
        If feld.Tag = "Uhrzeit" Then
        text = text & Format(feld.Range.text, "hh:mm") & ";"
        End If
        If feld.Tag = "AnruferName" Then
              text = text & feld.Range.text & ";"
        End If
        If feld.Tag = "Unternehmen" Then
            text = text & feld.Range.text & ";"
        End If
        If feld.Tag = "Unternehmensart" Then
            text = text & feld.Range.text & ";"
        End If
        If feld.Tag = "Telefon" Then
            text = text & feld.Range.text & ";"
        End If
        If feld.Tag = "Email" Then
            text = text & feld.Range.text & ";"
        End If
        If feld.Tag = "Weiteres" Then
            text = text & feld.Range.text & ";"
        End If
      
        If feld.Tag = "Anliegen" Then
            text = text & feld.Range.text & ";"
        End If
        If feld.Tag = "RPKS_Az" Then
            text = text & feld.Range.text & ";"
        End If
                
        If feld.Tag = "RPKS_Thema" Then
         text = text & feld.Range.text & ";"
        End If
                     
       
    'Next feld
    
    TXTBefüllen = text
    
    text = ""
    
End Function
Function TXTNameBauen()
Dim feld As ContentControl
Dim txtname As String
'Dim Datum As String
'Dim Uhrzeit As String
   
        
    txtname = ""
         
    For Each feld In ActiveDocument.ContentControls
        If feld.Tag = "Datum" Then
            txtname = txtname & Format(feld.Range.text, "yyyy_mm_dd") & "_"
        End If
        If feld.Tag = "Uhrzeit" Then
            txtname = txtname & "_" & Format(feld.Range.text, "hh_mm") & "_Soforthilfe"
        End If
               
        
               
        If feld.Tag = "AnruferName" Then
            txtname = txtname & "_" & feld.Range.text
        End If
                             
        If feld.Tag = "RPKS_Az" Then
           txtname = txtname & "_" & feld.Range.text
                              
        End If
        
        
    Next feld
    
             
    txtname = Replace(txtname, vbTab, "")
    txtname = Replace(txtname, "Art des Unternehmens", "")
    txtname = Replace(txtname, "Art des Unternehmens", "")
    txtname = konvert(txtname)
    txtname = Replace(txtname, "__", "_")
    txtname = txtname & ".csv"

    TXTNameBauen = txtname
    
    
End Function

