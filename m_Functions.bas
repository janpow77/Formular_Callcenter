Attribute VB_Name = "m_Functions"
Option Explicit

Function IsdocPfad(ByVal fName As String) As Boolean
    'gibt True aus, wenn der Ordner existiert

    If (Dir(fName, vbDirectory) <> "") Then
        IsdocPfad = True
    Else
        IsdocPfad = False
    End If
End Function
 Function konvert(strIn As String) As String
' Funktion, die illegale Buchstaben konvertiert.
    Dim i As Integer
    Const str = "\,/,:,*,?,"",<,>,|,&,$,%,§,"
    konvert = strIn
    For i = 1 To Len(str)
        konvert = Replace(konvert, Mid$(str, i, 1), "_")
    Next i
End Function
