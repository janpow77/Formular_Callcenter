Attribute VB_Name = "m_tools"
Option Explicit

Sub ResetForm()
' Diese Prozdedur setzt das Formular vollst�ndig zur�ck.
' Die Inhalte k�nnen �ber Angaben aktualsieren �berschrieben werden.
    ActiveDocument.CustomDocumentProperties("DokumentZWS").Value = False
    
End Sub



Sub test()
'Schreibt die Dokumenteneigenschaften

With ActiveDocument.CustomDocumentProperties
.Add Name:="DokumentEmail3", _
        LinkToContent:=False, _
        Type:=msoPropertyTypeString, _
        Value:="jan.riener@vwvg.de"
        End With
End Sub


