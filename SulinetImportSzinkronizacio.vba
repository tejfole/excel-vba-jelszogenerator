Sub SulinetImportSzinkronizacio()
    Dim ws As Worksheet
    Dim wsForras As Worksheet
    Dim forrasFile As String
    Dim forrasWorkbook As Workbook
    Dim fd As FileDialog
    Dim ujSor As Long
    Dim i As Long, j As Long
    Dim nev As String, om As String, uid As String
    Dim uidIndex As Long, nevIndex As Long, omIndex As Long
    Dim actionIndex As Long, snIndex As Long, givennameIndex As Long
    Dim displaynameIndex As Long, passwordIndex As Long, affiliationIndex As Long
    Dim statusIndex As Long, mailIndex As Long, eduroamIndex As Long, eduidIndex As Long
    Dim forrasNevIndex As Long, forrasOmIndex As Long
    Dim uidLista As Collection
    Dim ujUID As String
    Dim counter As Integer
    
    Set ws = ActiveSheet
    
    ' Forrás fájl kiválasztása FileDialog-gal
    Set fd = Application.FileDialog(4)
    With fd
        .Title = "Válassz forrás fájlt (Tanulók listája)"
        .AllowMultiSelect = False
        .Filters.Add "Excel Files", "*.xlsx;*.xls"
        If .Show = -1 Then
            forrasFile = .SelectedItems(1)
        Else
            MsgBox "Nem választottál fájlt!", vbExclamation
            Exit Sub
        End If
    End With
    
    ' Forrás fájl megnyitása
    On Error Resume Next
    Set forrasWorkbook = Workbooks.Open(forrasFile)
    On Error GoTo 0
    
    If forrasWorkbook Is Nothing Then
        MsgBox "Nem tudtam megnyitni a fájlt!", vbCritical
        Exit Sub
    End If
    
    ' Munkalap keresése a forrásban
    On Error Resume Next
    Set wsForras = forrasWorkbook.Sheets("Tanulók listája")
    On Error GoTo 0
    
    If wsForras Is Nothing Then
        MsgBox "Nem találom a 'Tanulók listája' munkalapot!", vbCritical
        forrasWorkbook.Close
        Exit Sub
    End If
    
    ' Fejlécek keresése a cél fájlban
    ujSor = 1
    Do While ujSor <= 100
        If wsForras.Cells(ujSor, 1).Value = "Neve" Then Exit Do
        ujSor = ujSor + 1
    Loop
    
    If ujSor > 100 Then
        MsgBox "Nem találom az oszlopfejléceket a forrás fájlban!", vbCritical
        forrasWorkbook.Close
        Exit Sub
    End If
    
    ' Oszlopindexek keresése forrásban
    forrasNevIndex = 0
    forrasOmIndex = 0
    For j = 1 To 20
        If wsForras.Cells(ujSor, j).Value = "Neve" Then forrasNevIndex = j
        If wsForras.Cells(ujSor, j).Value = "Oktatási azonosító" Then forrasOmIndex = j
    Next j
    
    If forrasNevIndex = 0 Or forrasOmIndex = 0 Then
        MsgBox "Nem találom az oszlopokat a forrás fájlban!", vbCritical
        forrasWorkbook.Close
        Exit Sub
    End If
    
    ' Fejlécek keresése a cél fájlban (Worksheet)
    ujSor = 1
    Do While ujSor <= 100
        If ws.Cells(ujSor, 1).Value = "UID" Then Exit Do
        ujSor = ujSor + 1
    Loop
    
    If ujSor > 100 Then
        MsgBox "Nem találom az oszlopfejléceket a cél fájlban!", vbCritical
        forrasWorkbook.Close
        Exit Sub
    End If
    
    ' Oszlopindexek keresése célban
    uidIndex = 0
    actionIndex = 0
    omIndex = 0
    snIndex = 0
    givennameIndex = 0
    displaynameIndex = 0
    passwordIndex = 0
    affiliationIndex = 0
    statusIndex = 0
    mailIndex = 0
    eduroamIndex = 0
    eduidIndex = 0
    
    For j = 1 To 20
        If ws.Cells(ujSor, j).Value = "UID" Then uidIndex = j
        If ws.Cells(ujSor, j).Value = "ACTION" Then actionIndex = j
        If ws.Cells(ujSor, j).Value = "OM" Then omIndex = j
        If ws.Cells(ujSor, j).Value = "SN" Then snIndex = j
        If ws.Cells(ujSor, j).Value = "GIVENNAME" Then givennameIndex = j
        If ws.Cells(ujSor, j).Value = "DISPLAYNAME" Then displaynameIndex = j
        If ws.Cells(ujSor, j).Value = "PASSWORD" Then passwordIndex = j
        If ws.Cells(ujSor, j).Value = "AFFILIATION" Then affiliationIndex = j
        If ws.Cells(ujSor, j).Value = "STATUS" Then statusIndex = j
        If ws.Cells(ujSor, j).Value = "MAIL" Then mailIndex = j
        If ws.Cells(ujSor, j).Value = "EDUROAM" Then eduroamIndex = j
        If ws.Cells(ujSor, j).Value = "EDUID" Then eduidIndex = j
    Next j
    
    If uidIndex = 0 Or actionIndex = 0 Or omIndex = 0 Or snIndex = 0 Or givennameIndex = 0 Then
        MsgBox "Hiányznak a szükséges oszlopok a cél fájlban!", vbCritical
        forrasWorkbook.Close
        Exit Sub
    End If
    
    ' Meglévő UID-k gyűjtése
    Set uidLista = New Collection
    Dim ujCelSor As Long
    ujCelSor = ujSor + 1
    
    Do While ws.Cells(ujCelSor, uidIndex).Value <> ""
        uidLista.Add ws.Cells(ujCelSor, uidIndex).Value
        ujCelSor = ujCelSor + 1
    Loop
    
    Application.ScreenUpdating = False
    
    Randomize
    
    ' Feldolgozás
    Dim forrasSor As Long
    forrasSor = ujSor + 1
    Dim szinkronizaltRowCount As Long
    szinkronizaltRowCount = 0
    
    ' Tanulók feldolgozása a forrásból
    Do While wsForras.Cells(forrasSor, forrasNevIndex).Value <> ""
        nev = wsForras.Cells(forrasSor, forrasNevIndex).Value
        om = wsForras.Cells(forrasSor, forrasOmIndex).Value
        
        If nev <> "" Then
            ujUID = GeneralUID(nev, uidLista)
            
            ' Megnézzük, van-e már ilyen UID
            Dim var As Variant
            Dim uidLetezik As Boolean
            uidLetezik = False
            
            On Error Resume Next
            For Each var In uidLista
                If var = ujUID Then
                    uidLetezik = True
                    Exit For
                End If
            Next var
            On Error GoTo 0
            
            ujCelSor = ujSor + 1
            Dim ujRowSzukseges As Boolean
            ujRowSzukseges = True
            
            ' Keresés meglévő UID-ban
            Do While ws.Cells(ujCelSor, uidIndex).Value <> ""
                If ws.Cells(ujCelSor, uidIndex).Value = ujUID Then
                    ujRowSzukseges = False
                    ' UPDATE: csak az ACTION-t módosítjuk UPDATE-re
                    ws.Cells(ujCelSor, actionIndex).Value = "UPDATE"
                    Exit Do
                End If
                ujCelSor = ujCelSor + 1
            Loop
            
            ' Új sor hozzáadása, ha szükséges
            If ujRowSzukseges Then
                Dim utolsoSor As Long
                utolsoSor = ujSor + 1
                Do While ws.Cells(utolsoSor, uidIndex).Value <> ""
                    utolsoSor = utolsoSor + 1
                Loop
                
                ' Új sor adatai
                ws.Cells(utolsoSor, uidIndex).Value = ujUID
                ws.Cells(utolsoSor, actionIndex).Value = "ADD"
                ws.Cells(utolsoSor, omIndex).Value = om
                
                ' Név szétszedése
                Dim sn As String, givenname As String
                SzetnezNev nev, sn, givenname
                ws.Cells(utolsoSor, snIndex).Value = sn
                ws.Cells(utolsoSor, givennameIndex).Value = givenname
                ws.Cells(utolsoSor, displaynameIndex).Value = nev
                
                ' Jelszó generálása csak ADD-nél!
                ws.Cells(utolsoSor, passwordIndex).Value = JelszóGenerálás(12, False)
                
                ' Alapértelmezett értékek
                If affiliationIndex > 0 Then ws.Cells(utolsoSor, affiliationIndex).Value = "student"
                If statusIndex > 0 Then ws.Cells(utolsoSor, statusIndex).Value = "active"
                If mailIndex > 0 Then ws.Cells(utolsoSor, mailIndex).Value = "disabled"
                If eduroamIndex > 0 Then ws.Cells(utolsoSor, eduroamIndex).Value = "enabled"
                If eduidIndex > 0 Then ws.Cells(utolsoSor, eduidIndex).Value = "enabled"
                
                uidLista.Add ujUID
            End If
            
            szinkronizaltRowCount = szinkronizaltRowCount + 1
        End If
        
        forrasSor = forrasSor + 1
    Loop
    
    ' DELETE-ek jelölése - akik a régi listában vannak, de az új listában nincsenek
    Dim forrasUidLista As Collection
    Set forrasUidLista = New Collection
    
    forrasSor = ujSor + 1
    Do While wsForras.Cells(forrasSor, forrasNevIndex).Value <> ""
        nev = wsForras.Cells(forrasSor, forrasNevIndex).Value
        ujUID = GeneralUID(nev, forrasUidLista)
        forrasUidLista.Add ujUID
        forrasSor = forrasSor + 1
    Loop
    
    ' Célban végigmegyünk és jelöljük a törlendőket
    ujCelSor = ujSor + 1
    Do While ws.Cells(ujCelSor, uidIndex).Value <> ""
        Dim vanAForrasban As Boolean
        vanAForrasban = False
        
        On Error Resume Next
        For Each var In forrasUidLista
            If var = ws.Cells(ujCelSor, uidIndex).Value Then
                vanAForrasban = True
                Exit For
            End If
        Next var
        On Error GoTo 0
        
        ' Ha nincs a forrásban ÉS student, akkor DELETE
        If Not vanAForrasban And ws.Cells(ujCelSor, affiliationIndex).Value = "student" Then
            ws.Cells(ujCelSor, actionIndex).Value = "DELETE"
        End If
        
        ujCelSor = ujCelSor + 1
    Loop
    
    Application.ScreenUpdating = True
    
    MsgBox "Szinkronizálás kész!" & vbCrLf & _
           "Feldolgozott sorok: " & szinkronizaltRowCount, vbInformation
    
    ' Forrás fájl bezárása
    forrasWorkbook.Close SaveChanges:=False
    
End Sub

Function GeneralUID(nev As String, uidLista As Collection) As String
    Dim uid As String
    Dim ujUID As String
    Dim counter As Integer
    
    ' Név tisztítása: ékezetek eltávolítása
    uid = nev
    uid = Replace(uid, "á", "a")
    uid = Replace(uid, "é", "e")
    uid = Replace(uid, "í", "i")
    uid = Replace(uid, "ó", "o")
    uid = Replace(uid, "ö", "o")
    uid = Replace(uid, "ő", "o")
    uid = Replace(uid, "ú", "u")
    uid = Replace(uid, "ü", "u")
    uid = Replace(uid, "ű", "u")
    uid = Replace(uid, " ", "")
    uid = Replace(uid, "-", "")
    uid = LCase(uid)
    
    ' Csak betűk és számok
    Dim i As Integer, newUID As String
    newUID = ""
    For i = 1 To Len(uid)
        Dim ch As String
        ch = Mid(uid, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then
            newUID = newUID & ch
        End If
    Next i
    
    ujUID = newUID
    
    ' Duplikáció kezelése
    counter = 1
    Dim var As Variant
    Do
        Dim letezik As Boolean
        letezik = False
        
        On Error Resume Next
        For Each var In uidLista
            If var = ujUID Then
                letezik = True
                Exit For
            End If
        Next var
        On Error GoTo 0
        
        If letezik Then
            ujUID = newUID & CStr(counter)
            counter = counter + 1
        Else
            Exit Do
        End If
    Loop
    
    GeneralUID = ujUID
End Function

Sub SzetnezNev(teljesNev As String, ByRef sn As String, ByRef givenname As String)
    Dim parts() As String
    Dim i As Integer
    
    ' Szóközzel szétszedés
    parts = Split(teljesNev, " ")
    
    If UBound(parts) >= 1 Then
        ' Utolsó szó = keresztnév
        givenname = parts(UBound(parts))
        
        ' Az előtte lévő összes = vezetéknév
        sn = ""
        For i = 0 To UBound(parts) - 1
            If sn <> "" Then sn = sn & " "
            sn = sn & parts(i)
        Next i
    Else
        ' Ha csak egy szó van
        sn = teljesNev
        givenname = ""
    End If
End Sub

Function JelszóGenerálás(hossz As Integer, Optional speciálisKarakter As Boolean = False) As String
    Dim karakterek As String
    Dim jelszó As String
    Dim i As Integer
    Dim randomIndex As Integer
    
    karakterek = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
    
    If speciálisKarakter Then
        karakterek = karakterek & "!@#$%^&*()_+-=[]{}|;:,.<>?"
    End If
    
    jelszó = ""
    
    For i = 1 To hossz
        randomIndex = Int(Rnd() * Len(karakterek)) + 1
        jelszó = jelszó & Mid(karakterek, randomIndex, 1)
    Next i
    
    JelszóGenerálás = jelszó
End Function
