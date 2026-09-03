Sub JelszóKitöltés()
    Dim ws As Worksheet
    Dim oszlop As String
    Dim oszlopIndex As Long
    Dim utolsóSor As Long
    Dim i As Long
    Dim jelszó As String
    Dim hossz As Integer
    Dim speciális As Boolean
    
    Set ws = ActiveSheet
    
    ' Oszlop bekérése
    oszlop = InputBox("Melyik oszlopba szeretnéd a jelszót?" & vbCrLf & _
                      "Pl.: A, B, C, D, stb.", "Oszlop kiválasztása")
    
    If oszlop = "" Then Exit Sub
    
    ' Oszlop indexének meghatározása
    On Error Resume Next
    oszlopIndex = ws.Columns(oszlop).Column
    On Error GoTo 0
    
    If oszlopIndex = 0 Then
        MsgBox "Érvénytelen oszlop! Kérjük, adj meg egy érvényes oszlopjelet (A, B, C, stb.)", vbCritical
        Exit Sub
    End If
    
    ' Jelszó hosszának bekérése
    hossz = InputBox("Mekkora legyen a jelszó hossza? (alapértelmezett: 12)", "Jelszó hossza", 12)
    
    If hossz < 4 Then
        MsgBox "A jelszó hossza legalább 4 karakter legyen!", vbExclamation
        Exit Sub
    End If
    
    ' Speciális karakterek bekérése
    If MsgBox("Szeretnél-e speciális karaktereket (@, #, $, stb.) a jelszóban?", vbYesNo, "Speciális karakterek") = vbYes Then
        speciális = True
    Else
        speciális = False
    End If
    
    ' Utolsó sor megkeresése - JAVÍTÁS: az első oszlop helyett az aktuális oszlopban keresünk
    ' de csak a 2. sortól kezdve (fejléc után)
    utolsóSor = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Ha nincs adat, vagy csak fejléc van
    If utolsóSor < 2 Then
        MsgBox "Nincs adat a munkalap alatt (csak fejléc van)!", vbExclamation
        Exit Sub
    End If
    
    ' Kitöltés megkezdése
    Dim kitöltöttek As Long
    Dim kihagyottak As Long
    kitöltöttek = 0
    kihagyottak = 0
    
    Application.ScreenUpdating = False
    
    For i = 2 To utolsóSor ' A 2. sortól kezdünk (1. sor a fejléc)
        ' Csak az üres cellákba írunk
        If IsEmpty(ws.Cells(i, oszlopIndex)) Then
            jelszó = JelszóGenerálás(hossz, speciális)
            ws.Cells(i, oszlopIndex).Value = jelszó
            kitöltöttek = kitöltöttek + 1
        Else
            ' Ha már van tartalom, kihagyjuk
            kihagyottak = kihagyottak + 1
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox "Kész!" & vbCrLf & _
           "Oszlop: " & UCase(oszlop) & vbCrLf & _
           "Feldolgozott sorok: " & (utolsóSor - 1) & vbCrLf & _
           "Új jelszavak: " & kitöltöttek & vbCrLf & _
           "Kihagyott (már van adat): " & kihagyottak & vbCrLf & _
           "Jelszó hossz: " & hossz & vbCrLf & _
           "Speciális karakterek: " & IIf(speciális, "Igen", "Nem"), vbInformation
    
End Sub

Function JelszóGenerálás(hossz As Integer, Optional speciálisKarakter As Boolean = False) As String
    Dim karakterek As String
    Dim jelszó As String
    Dim i As Integer
    Dim randomIndex As Integer
    
    ' Alapvető karakterek (számok és betűk)
    karakterek = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
    
    ' Speciális karakterek hozzáadása, ha szükséges
    If speciálisKarakter Then
        karakterek = karakterek & "!@#$%^&*()_+-=[]{}|;:,.<>?"
    End If
    
    jelszó = ""
    
    ' Jelszó generálása
    For i = 1 To hossz
        randomIndex = Int((Len(karakterek) - 1) * Rnd()) + 1
        jelszó = jelszó & Mid(karakterek, randomIndex, 1)
    Next i
    
    JelszóGenerálás = jelszó
End Function
