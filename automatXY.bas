Attribute VB_Name = "testowy_1"
    Public wiersz As Variant
    Public kolumnaNazwaFirmy As Variant
    Public kolumnaNazwaPliku As Variant
    Public kolumnawlacz As Variant
    Public kolumnaplaceholder20 As Variant
    Public kolumnaplaceholder21 As Variant
    Public kolumnaColumnPrice As Variant
    Public kolumnplaceholder20IND As Variant
    Public nrkolumny As Variant
    Public kolumna1Stock As Variant
    Public kolumnaL2Stock As Variant
    Public kolumnaL3Stock As Variant
    Public kolumnaplaceholder1 As Variant
    Public kolumnaplaceholder2 As Variant
    Public kolumnaplaceholder3 As Variant
    Public kolumnaplaceholder4Ribbon As Variant
    Public kolumnaplaceholder5 As Variant
    Public kolumnaplaceholder12 As Variant
    Public kolumnaplaceholder4 As Variant
    Public kolumnaplaceholder6 As Variant
    Public kolumnaplaceholder16 As Variant
    Public kolumnaplaceholder7 As Variant
    Public kolumnaplaceholder8 As Variant
    Public kolumnaplaceholder10 As Variant
    Public kolumnaplaceholder11 As Variant
    Public kolumnaplaceholder15 As Variant
    Public kolumnaplaceholder17 As Variant
    Public kolumnaplaceholder18 As Variant
    Public kolumnaplaceholder27 As Variant
    Public nazwasheetuIND As Variant
    Public kolumnaSheetIND As Variant
    Public kolumnaZapisXLS As Variant
    Public kolumnaplaceholder4NiedoPL As Variant
    Public kolumnaplaceholder14 As Variant
    Public kolumnaUsuwanieNaglowkow As Variant
    Public kolumnaUsuwaniePustegoStocku As Variant
    Public kolumnaUsuwanieplaceholder12placeholder28 As Variant
    Public kolumnaUsuwanieplaceholder4NieRibbon As Variant
    Public kolumnaZostawiaTylkoStockplaceholder6 As Variant
    Public kolumnaUsuwaplaceholder13 As Variant
    Public kolumnaColumnStock As Variant
    Public kolumnaStockIND As Variant
    Public kolumnaStockDo100 As Variant
    Public StaraNazwa As String, NowaNazwa As String, Sciezka As String
    

    
    Sub tworzeniecennika()
    
    'wy��czenie aktualizacji widoku okna excal podczas trwania makra
    Application.ScreenUpdating = False
    
    'Aby zapobiec b��dom usuwa sheet o nazwie pricelist je�li taki istnieje
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("pricelist").Delete
    Application.DisplayAlerts = True
    
    Dim StartTime As Double
    Dim EndTime As String
    
    
    'zmienna zapisuje czas rozpoczecia dzialania makra
    StartTime = Timer
    '�cie�ka zapisu plik�w
    Sciezka = "C:\cennikiautomat" & "\"
        
    'Sta�e odpowiadaj�ce numerom konkretnych kolumn w sheecie modu�
    kolumnaNazwaPliku = 4
    kolumnaNazwaFirmy = 8
    kolumnawlacz = 10
    kolumnaplaceholder20 = 11
    kolumnaplaceholder21 = 12
    kolumnaColumnPrice = 13
    kolumnplaceholder20IND = 14
    kolumnaplaceholder19 = 15
    kolumna1Stock = 16
    kolumnaL2Stock = 17
    kolumnaL3Stock = 18
    kolumnaplaceholder25 = 19
    kolumnaplaceholder1 = 20
    kolumnaplaceholder2 = 21
    kolumnaplaceholder3 = 22
    kolumnaplaceholder4Ribbon = 23
    kolumnaplaceholder5 = 24
    kolumnaplaceholder12 = 25
    kolumnaplaceholder4 = 26
    kolumnaplaceholder6 = 27
    kolumnaplaceholder16 = 28
    kolumnaplaceholder7 = 29
    kolumnaplaceholder8 = 30
    kolumnaplaceholder10 = 31
    kolumnaplaceholder11 = 32
    kolumnaplaceholder15 = 33
    kolumnaplaceholder17 = 34
    kolumnaplaceholder18 = 35
    kolumnaplaceholder27 = 36
    kolumnaCzySheetIND = 37
    kolumnaSheetIND = 38
    kolumnaZapisXLS = 39
    kolumnaZapisCSV = 40
    kolumnaplaceholder4NiedoPL = 41
    kolumnaplaceholder14 = 42
    kolumnaUsuwanieNaglowkow = 43
    kolumnaUsuwaniePustegoStocku = 44
    kolumnaUsuwanieplaceholder12placeholder28 = 45
    kolumnaUsuwanieplaceholder4NieRibbon = 46
    kolumnaZostawiaTylkoStockplaceholder6 = 47
    kolumnaUsuwaplaceholder13 = 48
    kolumnaStockDo100 = 49
    kolumnaColumnStock = 50
    kolumnaStockIND = 51
       
    
    'zmienna kt�ra odpowiada ka�demu wierszowi z nazw� firmy (najpierw zlciza ich ilo��)
    r = Cells(Rows.Count, "D").End(xlUp).Row
    
    'P�tla kt�ra wykonuje si� dla ka�dej firmy
    For wiersz = 3 To r
    
    'Warunek sprawdzaj�cy czy makro dla konkretnej firmy ma by� uruchomione
    If Sheets("modu�").Cells(wiersz, kolumnawlacz).Value = "+" Then
        
    'Zmienna pobieraj�ca nazw� sheetu indywidualnego z kt�rego maj� by� usuni�te konkretne produkty
    nazwasheetuIND = Sheets("modu�").Cells(wiersz, kolumnaSheetIND).Value
    
    'Zmienna pobieraj�ca nazw� firmy
    komorkaznazwa = Sheets("modu�").Cells(wiersz, kolumnaNazwaFirmy)
    
'W przypadku b��du uruchamia prplaceholder10dur� obs�ugi b��d�w
On Error GoTo ErrHandler
        
        
    'Makro kolejno sprawdza czy zaznaczone s� kom�rki w sheecie modu�. W przypadk uzaznaczenia ich znakiem "+" uruchamia kolejno wybrane prplaceholder10dury.
    'Po ich zako�czeniu wszystko zaczyna si� ponownie dla nast�pnej firmy.
        Call kopiowanie_SL
        If Sheets("modu�").Cells(wiersz, kolumnaplaceholder19).Value = "+" Then Call placeholder19
        If Sheets("modu�").Cells(wiersz, kolumna1Stock).Value = "+" Then Call placeholder22
        If Sheets("modu�").Cells(wiersz, kolumnaL2Stock).Value = "+" Then Call placeholder23
        If Sheets("modu�").Cells(wiersz, kolumnaL3Stock).Value = "+" Then Call placeholder24
        If Sheets("modu�").Cells(wiersz, kolumnaplaceholder25).Value = "+" Then Call placeholder25
        If Sheets("modu�").Cells(wiersz, kolumnaColumnStock).Value = "+" Then Call ColumnStock
        If Sheets("modu�").Cells(wiersz, kolumnaplaceholder20).Value = "+" Then Call placeholder20
        If Sheets("modu�").Cells(wiersz, kolumnaplaceholder21).Value = "+" Then Call placeholder21
        If Sheets("modu�").Cells(wiersz, kolumnaColumnPrice).Value = "+" Then Call ColumnPrice
        If Sheets("modu�").Cells(wiersz, kolumnaplaceholder1).Value = "+" Then Call Usuwaplaceholder1
        If Sheets("modu�").Cells(wiersz, kolumnaplaceholder2).Value = "+" Then Call Usuwaplaceholder2
        If Sheets("modu�").Cells(wiersz, kolumnaplaceholder3).Value = "+" Then Call Usuwaplaceholder3
        If Sheets("modu�").Cells(wiersz, kolumnaplaceholder4Ribbon).Value = "+" Then Call Usuwaplaceholder4Ribbon
        If Sheets("modu�").Cells(wiersz, kolumnaplaceholder5).Value = "+" Then Call Usuwaplaceholder5
        If Sheets("modu�").Cells(wiersz, kolumnaplaceholder12).Value = "+" Then Call Usuwaplaceholder12
        If Sheets("modu�").Cells(wiersz, kolumnaplaceholder4).Value = "+" Then Call Usuwaplaceholder4
        If Sheets("modu�").Cells(wiersz, kolumnaplaceholder6).Value = "+" Then Call Usuwaplaceholder6
        If Sheets("modu�").Cells(wiersz, kolumnaplaceholder16).Value = "+" Then Call Usuwaplaceholder16
        If Sheets("modu�").Cells(wiersz, kolumnaplaceholder7).Value = "+" Then Call Usuwaplaceholder7
        If Sheets("modu�").Cells(wiersz, kolumnaplaceholder8).Value = "+" Then Call Usuwaplaceholder8
        If Sheets("modu�").Cells(wiersz, kolumnaplaceholder10).Value = "+" Then Call Usuwaplaceholder10
        If Sheets("modu�").Cells(wiersz, kolumnaplaceholder11).Value = "+" Then Call Usuwaplaceholder11
        If Sheets("modu�").Cells(wiersz, kolumnaplaceholder15).Value = "+" Then Call Usuwaplaceholder15
        If Sheets("modu�").Cells(wiersz, kolumnaplaceholder17).Value = "+" Then Call Usuwaplaceholder17
        If Sheets("modu�").Cells(wiersz, kolumnaplaceholder18).Value = "+" Then Call Usuwaplaceholder18
        If Sheets("modu�").Cells(wiersz, kolumnaplaceholder27).Value = "+" Then Call Usuwaplaceholder27
        If Sheets("modu�").Cells(wiersz, kolumnaCzySheetIND).Value = "+" Then Call UsuwaIND
        If Sheets("modu�").Cells(wiersz, kolumnaplaceholder4NiedoPL).Value = "+" Then Call Usuwaplaceholder4NiedoPL
        If Sheets("modu�").Cells(wiersz, kolumnaplaceholder14).Value = "+" Then Call placeholder14
        If Sheets("modu�").Cells(wiersz, kolumnaUsuwanieNaglowkow).Value = "+" Then Call UsuwanienNaglowkow
        If Sheets("modu�").Cells(wiersz, kolumnaUsuwaniePustegoStocku).Value = "+" Then Call UsuwaniePustegoStocku
        If Sheets("modu�").Cells(wiersz, kolumnaUsuwanieplaceholder12placeholder28).Value = "+" Then Call Usuwaplaceholder12placeholder28
        If Sheets("modu�").Cells(wiersz, kolumnaUsuwanieplaceholder4NieRibbon).Value = "+" Then Call Usuwanieplaceholder4NieRibbon
        If Sheets("modu�").Cells(wiersz, kolumnaZostawiaTylkoStockplaceholder6).Value = "+" Then Call ZostawiaTylkoStockplaceholder6
        If Sheets("modu�").Cells(wiersz, kolumnaUsuwaplaceholder13).Value = "+" Then Call Usuwaplaceholder13
        If Sheets("modu�").Cells(wiersz, kolumnaStockDo100).Value = "+" Then Call OgraniczDo100
        Call Usuwanie_puste
        If Sheets("modu�").Cells(wiersz, kolumnaZapisXLS).Value = "+" Then Call ZapisXLS
        If Sheets("modu�").Cells(wiersz, kolumnaZapisCSV).Value = "+" Then Call ZapisCSV
    
    End If
    
Next
    
    'w��czenie aktualizacji widoku okna excal podczas trwania makra
    Application.ScreenUpdating = True
    EndTime = Format((Timer - StartTime) / 86400, "hh:mm:ss")
    
    ' Wiadomo�� po zako�czeniu makra wraz z czasem jaki zaj�o jego wykonanie
    MsgBox "Gotowe. Czas trwania: " & EndTime & " minut", vbInformation


Exit Sub

'Obs�uga b��d�w
ErrHandler:
    'Je�eli zosta� ju� utworzony sheet "pricelist" prplaceholder10dura go usunie aby nie wywo�ywa� dalszych b��d�w
    If ActiveSheet.Name = "pricelist" Then ActiveSheet.Delete
    'Wy�wietla info dla jakiego cennika wyst�pi� b��d
    MsgBox "Wyst�pi� bl�d dla cennika: " & vbNewLine & komorkaznazwa & vbNewLine & "Sprawd� poprawno�� wprowadzonych danych lub skontaktuj si� z administratoerm.", vbInformation
    'Po wy�wietleniu zostanie uruchomione makro dla kolejnego cennika tak aby tylko ten na kt�rym wyst�pi� b��d nie zsota� stworzony
    wiersz = wiersz + 1
    Resume
End Sub

Sub kopiowanie_SL()
    
    'Kopiuje kolumny od A do E z sheetu "szablon cen" i umieszcza je w nowym sheecie

    Sheets("szablon cen").Range("A:E").Copy
    Sheets.Add.Name = "pricelist"
    ActiveSheet.Paste

End Sub
Sub placeholder20()

    'Kopiowanie ... i umieszczanie ich w arkuszu z nowym cennikiem
    Sheets("szablon cen").Select
    Range("F:F").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("pricelist").Select
    Columns("F:F").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Ustawienie odpowiedniego formatu walutowego, pogrubienie czcionki, i ustawienie szeroko�ci kolumny
    Columns("F:F").NumberFormat = "#,##0.00 [$�-x-euro1]"
    Columns("F:H").ColumnWidth = 8
    Range("f4").Font.Bold = True
    
End Sub
Sub placeholder21()

    'Kopiowanie ... i umieszczanie ich w arkuszu z nowym cennikiem
    Sheets("szablon cen").Select
    Range("G:G").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("pricelist").Select
    Columns("F:F").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Ustawienie odpowiedniego formatu walutowego, pogrubienie czcionki, i ustawienie szeroko�ci kolumny
    Columns("F:F").NumberFormat = "#,##0.00 [$�-x-euro1]"
    Columns("F:H").ColumnWidth = 8
    Range("f4").Font.Bold = True
    Range("f4").Value = "price"

End Sub
Sub ColumnPrice()

    'Kopiowanie cen z kolumny wpisanej do kolumny "Wpisz symbol kolumny" i umieszczanie ich w arkuszu z nowym cennikiem
    kolumnplaceholder20IND = 14
    kolumnaColumnPrice = 13
    nrkolumny = Sheets("modu�").Cells(wiersz, kolumnplaceholder20IND).Value

    Sheets("szablon cen").Select
    Range(Cells(1, nrkolumny), Cells(5000, nrkolumny)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("pricelist").Select
    Columns("F:F").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Ustawienie odpowiedniego formatu walutowego, pogrubienie czcionki, i ustawienie szeroko�ci kolumny
    Columns("F:F").NumberFormat = "#,##0.00 [$�-x-euro1]"
    Columns("F:H").ColumnWidth = 8
    Range("f4").Font.Bold = True
    Range("f4").Value = "price"
    
    ' Usuwa wiersz je�eli w kom�rce z cen� jest wpsiane "x"
    ow = Cells(Rows.Count, "F").End(xlUp).Row
    For r = ow To 1 Step -1
        If Cells(r, "F").Value = "x" Then Rows(r).Delete
    Next
    
End Sub
Sub ColumnStock()

    'Kopiowanie stocku z kolumny wpisanej do kolumny "Wpisz symbol kolumny" i umieszczanie ich arkuszu z nowym cennikiem
    kolumnaStockIND = 51
    kolumnaColumnStock = 50
    nrkolumny = Sheets("modu�").Cells(wiersz, kolumnaStockIND).Value

    Sheets("szablon cen").Select
    Range(Cells(1, nrkolumny), Cells(5000, nrkolumny)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("pricelist").Select
    Columns("G:G").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("F:H").ColumnWidth = 8
    Range("G4").Font.Bold = True
    Range("G4").Value = "stock"
    
End Sub

Sub placeholder19()
    
    'Kopiowanie zwyk�ego stocku i umieszczanie go w arkuszu z nowym cennikiem
    Sheets("szablon cen").Select
    Range("H:H").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("pricelist").Select
    Columns("G:G").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("F:G").ColumnWidth = 8
    Range("F4:G4").Font.Bold = True
End Sub
Sub placeholder22()
    
    'Kopiowanie Stocku zwyk�ego z ... i umieszczanie go w arkuszu z nowym cennikiem
    Sheets("szablon cen").Select
    Range("BN:BN").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("pricelist").Select
    Columns("G:G").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("F:G").ColumnWidth = 8
    Range("F4:G4").Font.Bold = True
    Range("G4").Value = "stock"
End Sub
Sub placeholder23()
    
    'Kopiowanie Stocku zwyk�ego z placeholder7 N i umieszczanie go w arkuszu z nowym cennikiem
    Sheets("szablon cen").Select
    Range("BO:BO").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("pricelist").Select
    Columns("G:G").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("F:G").ColumnWidth = 8
    Range("F4:G4").Font.Bold = True
    Range("G4").Value = "stock"
End Sub
Sub placeholder24()
    
    'Kopiowanie Stocku z Lex i umieszczanie go w arkuszu z nowym cennikiem
    Sheets("szablon cen").Select
    Range("BP:BP").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("pricelist").Select
    Columns("G:G").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("F:G").ColumnWidth = 8
    Range("F4:G4").Font.Bold = True
    Range("G4").Value = "stock"
End Sub
Sub placeholder25()
    
    'Kopiowanie Stocku ... i umieszczanie go w arkuszu z nowym cennikiem
    Sheets("szablon cen").Select
    Range("BQ:BQ").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("pricelist").Select
    Columns("G:G").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("F:G").ColumnWidth = 8
    Range("F4:G4").Font.Bold = True
    Range("G4").Value = "stock"
End Sub
Sub Usuwaplaceholder1()
    'Usuwanie wszystkich placeholder26 placeholder1
    Sheets("pricelist").Select

    ow = Cells(Rows.Count, "D").End(xlUp).Row
        For r = ow To 1 Step -1
        If Cells(r, "D") Like "*placeholder1*" Then Rows(r).Delete
    Next
End Sub
Sub Usuwaplaceholder2()
    'Usuwanie wszystkich placeholder26 placeholder2
    Sheets("pricelist").Select

    ow = Cells(Rows.Count, "D").End(xlUp).Row
    For r = ow To 1 Step -1
        If Cells(r, "A") Like "*placeholder27*" And Cells(r, "D") Like "* W *" Then Rows(r).Delete
    Next
End Sub
Sub Usuwaplaceholder3()
    
    'Usuwanie wszystkich placeholder18�w Spare Parts
    Sheets("pricelist").Select

    ow = Cells(Rows.Count, "D").End(xlUp).Row
    For r = ow To 1 Step -1
        If Cells(r, "A") Like "*placeholder18*" And Not Cells(r, "D") Like "*Toner*" Then Rows(r).Delete
        If Cells(r, "A") Like "*placeholder18*" And Cells(r, "D") Like "*Waste*" Then Rows(r).Delete
    Next
End Sub
Sub Usuwaplaceholder4Ribbon()

'Usuwa placeholder6 Ribbon
    Sheets("pricelist").Select
    ow = Cells(Rows.Count, "D").End(xlUp).Row
    For r = ow To 1 Step -1
        If Cells(r, "A").Text = "placeholder4" And Cells(r, "D") Like "*Ribbon*" Then Rows(r).Delete
    Next
    ow = Cells(Rows.Count, "A").End(xlUp).Row
    For r = ow To 5 Step -1
        If Cells(r, "C").Text = "Original placeholder4 print supplies" And Cells(r + 1, "C").Value = "" Then Rows(r).Resize(2).Delete
    Next
End Sub
Sub Usuwaplaceholder5()
    'Usuwa mark� placeholder5 wraz za nag��wkiem/nag��kami
    Sheets("pricelist").Select
    ow = Cells(Rows.Count, "A").End(xlUp).Row
    'Usuwa placeholder5
    For r = ow To 1 Step -1
        If Cells(r, "A").Value = "placeholder5" Then Rows(r).Delete
    Next
    ' Usuwa nag��wek placeholder5
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "placeholder5 Supplies" Then Rows(r).Resize(2).Delete
    Next
End Sub
Sub Usuwaplaceholder12()
    
    'Usuwa mark� placeholder12 wraz za nag��wkiem/nag��kami
    Sheets("pricelist").Select
    ow = Cells(Rows.Count, "A").End(xlUp).Row
    'Usuwa placeholder12
    For r = ow To 1 Step -1
        If Cells(r, "A").Value = "placeholder12" Then Rows(r).Delete
    Next
    ' Usuwa nag��wki placeholder12
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "placeholder12 Color Copier Supplies" Then Rows(r).Resize(2).Delete
    Next
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "placeholder12 B&W Copier Supplies" Then Rows(r).Resize(2).Delete
    Next
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "placeholder12 Color Laser Supplies" Then Rows(r).Resize(2).Delete
    Next
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "placeholder12 B&W Laser Supplies" Then Rows(r).Resize(2).Delete
    Next
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "placeholder12 Large Format Supplies" Then Rows(r).Resize(2).Delete
    Next
End Sub
Sub Usuwaplaceholder4()
    
    'Usuwa mark� placeholder4 wraz za nag��wkiem/nag��kami
    Sheets("pricelist").Select
    ow = Cells(Rows.Count, "A").End(xlUp).Row
    'Usuwa placeholder4
    For r = ow To 1 Step -1
        If Cells(r, "A").Value = "placeholder4" Then Rows(r).Delete
    Next
    ' Usuwa nag��wek placeholder4
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "Original placeholder4 print supplies" Then Rows(r).Resize(2).Delete
    Next

End Sub
Sub Usuwaplaceholder6()
    
    'Usuwa mark� placeholder6 wraz za nag��wkiem/nag��kami
    Sheets("pricelist").Select
    ow = Cells(Rows.Count, "A").End(xlUp).Row
    'Usuwa placeholder6
    For r = ow To 1 Step -1
        If Cells(r, "A").Value = "placeholder6" Then Rows(r).Delete
    Next
    ' Usuwa nag��wek placeholder6
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "Original placeholder6 print supplies" Then Rows(r).Resize(2).Delete
    Next

End Sub
Sub Usuwaplaceholder16()
    
    'Usuwa mark� placeholder9 wraz za nag��wkiem/nag��kami
    Sheets("pricelist").Select
    ow = Cells(Rows.Count, "A").End(xlUp).Row
    'Usuwa placeholder16
    For r = ow To 1 Step -1
        If Cells(r, "A").Value = "placeholder9" Then Rows(r).Delete
    Next
    ' Usuwa nag��wek placeholder16
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "Original placeholder9 print supplies" Then Rows(r).Resize(2).Delete
    Next

End Sub
Sub Usuwaplaceholder7()

    'Usuwa mark� placeholder7 wraz za nag��wkiem/nag��kami
    Sheets("pricelist").Select
    ow = Cells(Rows.Count, "A").End(xlUp).Row
    'Usuwa placeholder7
    For r = ow To 1 Step -1
        If Cells(r, "A").Value = "placeholder7" Then Rows(r).Delete
    Next
    ' Usuwa nag��wek placeholder7
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "Original placeholder7 print supplies" Then Rows(r).Resize(2).Delete
    Next

End Sub
Sub Usuwaplaceholder8()

    'Usuwa mark� placeholder8 wraz za nag��wkiem/nag��kami
    Sheets("pricelist").Select
    ow = Cells(Rows.Count, "A").End(xlUp).Row
    'Usuwa placeholder8
    For r = ow To 1 Step -1
        If Cells(r, "A").Value = "placeholder8" Then Rows(r).Delete
    Next
    ' Usuwa nag��wek placeholder8
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "Original placeholder8 print supplies" Then Rows(r).Resize(2).Delete
    Next

End Sub
Sub Usuwaplaceholder10()

    'Usuwa mark� placeholder10 wraz za nag��wkiem/nag��kami
    Sheets("pricelist").Select
    ow = Cells(Rows.Count, "A").End(xlUp).Row
    'Usuwa placeholder10
    For r = ow To 1 Step -1
        If Cells(r, "A").Value = "placeholder10" Then Rows(r).Delete
    Next
    ' Usuwa nag��wek placeholder10
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "placeholder10 Print Supplies" Then Rows(r).Resize(2).Delete
    Next

End Sub
Sub Usuwaplaceholder11()

    'Usuwa mark� placeholder11 wraz za nag��wkiem/nag��kami
    Sheets("pricelist").Select
    ow = Cells(Rows.Count, "A").End(xlUp).Row
    'Usuwa placeholder11
    For r = ow To 1 Step -1
        If Cells(r, "A").Value = "placeholder11" Then Rows(r).Delete
    Next
    ' Usuwa nag��wek placeholder11
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "placeholder11 Supplies" Then Rows(r).Resize(2).Delete
    Next

End Sub
Sub Usuwaplaceholder15()

    'Usuwa mark� placeholder15 wraz za nag��wkiem/nag��kami
    Sheets("pricelist").Select
    ow = Cells(Rows.Count, "A").End(xlUp).Row
    'Usuwa placeholder15
    For r = ow To 1 Step -1
        If Cells(r, "A").Value = "placeholder15" Then Rows(r).Delete
    Next
    ' Usuwa nag��wek placeholder15
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "Original placeholder15 Supplies" Then Rows(r).Resize(2).Delete
    Next

End Sub
Sub Usuwaplaceholder17()

    'Usuwa mark� placeholder17 wraz za nag��wkiem/nag��kami
    Sheets("pricelist").Select
    ow = Cells(Rows.Count, "A").End(xlUp).Row
    'Usuwa placeholder17
    For r = ow To 1 Step -1
        If Cells(r, "A").Value = "placeholder17" Then Rows(r).Delete
    Next
    ' Usuwa nag��wek placeholder17
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "Original placeholder17 print supplies" Then Rows(r).Resize(2).Delete
    Next

End Sub
Sub Usuwaplaceholder18()

    'Usuwa mark� placeholder18 wraz za nag��wkiem/nag��kami
    Sheets("pricelist").Select
    ow = Cells(Rows.Count, "A").End(xlUp).Row
    'Usuwa placeholder18
    For r = ow To 1 Step -1
        If Cells(r, "A").Value = "placeholder18" Then Rows(r).Delete
    Next
    ' Usuwa nag��wek placeholder18
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "Original placeholder18 print supplies" Then Rows(r).Resize(2).Delete
    Next

End Sub
Sub Usuwaplaceholder27()

    'Usuwa mark� placeholder27 wraz za nag��wkiem/nag��kami
    Sheets("pricelist").Select
    ow = Cells(Rows.Count, "A").End(xlUp).Row
    'Usuwa placeholder27
    For r = ow To 1 Step -1
        If Cells(r, "A").Value = "placeholder27" Then Rows(r).Delete
    Next
    'Usuwa placeholder27 SP
    For r = ow To 1 Step -1
        If Cells(r, "A").Value = "placeholder27 Spare Parts" Then Rows(r).Delete
    Next
    ' Usuwa nag��wek placeholder27
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "placeholder27 Color Machines" Or Cells(r, "C").Value = "placeholder27 B&W" Or Cells(r, "C").Value = "placeholder27 placeholder29 models" Or Cells(r, "C").Value = "placeholder27 Wide Format" Or Cells(r, "C").Value = "placeholder27 Original Spare Parts " Or Cells(r, "C").Value = "PLEASE NOTE: Prices and leadtime are to be confirmed before ordering" Then Rows(r).Resize(2).Delete
    Next

End Sub
Sub UsuwaIND()
'Usuwa z Sheetu IND.
    
    Sheets(nazwasheetuIND).Select
    Range("A:A").Copy
    Sheets("pricelist").Select
    Range("L1").PasteSpecial
    
    
        ow = Cells(Rows.Count, "B").End(xlUp).Row
        For r = ow To 1 Step -1
            For p = 1 To 200
                If Cells(r, "B").Value = Cells(p, "L") And Cells(p, "L") <> 0 Then Rows(r).Delete
            Next
        Next
    Range("L:L").Clear
End Sub
Sub ZapisXLS()
    Sheets("modu�").Select
    Cells(wiersz, kolumnaNazwaPliku).Copy
    Sheets("pricelist").Select
    Range("H1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Workbooks.Add
    ThisWorkbook.Sheets("pricelist").Copy before:=ActiveWorkbook.Sheets(1)
    Application.DisplayAlerts = False
    ActiveWorkbook.Sheets("Arkusz1").Delete
    Application.DisplayAlerts = True
    ActiveWorkbook.SaveAs Filename:=Sciezka & ActiveCell.Value & ".xlsx", FileFormat:=xlOpenXMLWorkbook
' Usuwa nazwe pliku i czysli sheet podstawowoy
    Range("H1").Delete
    Range("A1").Select
    Application.DisplayAlerts = False
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    ActiveSheet.Delete
    Application.DisplayAlerts = True
    Sheets("modu�").Select
End Sub
Sub ZapisCSV()
    Sheets("modu�").Select
    Cells(wiersz, kolumnaNazwaPliku).Copy
    Sheets("pricelist").Select
    Range("H1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Workbooks.Add
    ThisWorkbook.Sheets("pricelist").Copy before:=ActiveWorkbook.Sheets(1)
    Application.DisplayAlerts = False
    ActiveWorkbook.Sheets("Arkusz1").Delete
    Application.DisplayAlerts = True
    ActiveWorkbook.SaveAs Filename:=Sciezka & ActiveCell.Value & ".csv", FileFormat:=xlOpenXMLWorkbook
' Usuwa nazwe pliku i czysli sheet podstawowoy
    Range("H1").Delete
    Range("A1").Select
    Application.DisplayAlerts = False
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    ActiveSheet.Delete
    Application.DisplayAlerts = True
    Sheets("modu�").Select
End Sub
Sub Usuwaplaceholder4NiedoPL()
    
    Sheets("placeholder4 NIE DO PL").Select
    Range("B:B").Copy
    Sheets("pricelist").Select
    Range("L1").PasteSpecial
    
    
        ow = Cells(Rows.Count, "B").End(xlUp).Row
        For r = ow To 1 Step -1
            For p = 1 To 200
                If Cells(r, "B").Value = Cells(p, "L") And Cells(p, "L") <> 0 Then Rows(r).Delete
            Next
        Next
    Range("L:L").Clear
End Sub

Sub TestInputs()
    kolumnawlacz = 10
    kolumnaplaceholder20 = 11
    kolumnaplaceholder21 = 12
    kolumnaColumnPrice = 13
    kolumnaNazwaFirmy = 8
    
    kolumnaplaceholder19 = 15
    kolumna1Stock = 16
    kolumnaL2Stock = 17
    kolumnaL3Stock = 18
    kolumnaplaceholder25 = 19
    kolumnaColumnStock = 50
   
    kolumnaCzySheetIND = 37
    kolumnaSheetIND = 38
    
    kolumnplaceholder20IND = 14
    
    kolumnaplaceholder14 = 42
    
    kolumnaStockIND = 51
    
    For wiersz = 3 To 100
    
        'Sprawdza poprawno�� cen
        komorkaznazwa = Sheets("modu�").Cells(wiersz, kolumnaNazwaFirmy)
            If Sheets("modu�").Cells(wiersz, kolumnawlacz).Value = "+" Then
                If Sheets("modu�").Cells(wiersz, kolumnaplaceholder20).Value = "+" Then
                ElseIf Sheets("modu�").Cells(wiersz, kolumnaplaceholder21).Value = "+" Then
                ElseIf Sheets("modu�").Cells(wiersz, kolumnaColumnPrice).Value = "+" Then
                ElseIf Sheets("modu�").Cells(wiersz, kolumnaplaceholder14).Value = "+" Then
                Else: MsgBox "Ceny nie okre�lone dla: " & komorkaznazwa, vbInformation
            End If
            End If
        'Sprawdza poparawno�� stock�w
            If Sheets("modu�").Cells(wiersz, kolumnawlacz).Value = "+" Then
                If Sheets("modu�").Cells(wiersz, kolumnaplaceholder19).Value = "+" Then
                ElseIf Sheets("modu�").Cells(wiersz, kolumna1Stock).Value = "+" Then
                ElseIf Sheets("modu�").Cells(wiersz, kolumnaL2Stock).Value = "+" Then
                ElseIf Sheets("modu�").Cells(wiersz, kolumnaL3Stock).Value = "+" Then
                ElseIf Sheets("modu�").Cells(wiersz, kolumnaplaceholder25).Value = "+" Then
                ElseIf Sheets("modu�").Cells(wiersz, kolumnaColumnStock).Value = "+" Then
                Else: MsgBox "Stock nie okre�lony dla: " & komorkaznazwa, vbInformation
            End If
            End If
        'Sprawdza czy wpisana jest nazwa sheet IND.
            If Sheets("modu�").Cells(wiersz, kolumnawlacz).Value = "+" Then
                If Sheets("modu�").Cells(wiersz, kolumnaCzySheetIND).Value = "+" Then
                If Sheets("modu�").Cells(wiersz, kolumnaSheetIND) = "" Then MsgBox "Sheet IND. nie okre�lony dla: " & komorkaznazwa, vbInformation
            End If
            End If
        'Sprawdza czy wpisana jest kolumna z cenami
            If Sheets("modu�").Cells(wiersz, kolumnawlacz).Value = "+" Then
                If Sheets("modu�").Cells(wiersz, kolumnaColumnPrice).Value = "+" Then
                If Sheets("modu�").Cells(wiersz, kolumnplaceholder20IND) = "" Then MsgBox "Sheet IND. nie okre�lony dla: " & komorkaznazwa, vbInformation
            End If
            End If
        'Sprawdza czy wpisany jest kolumna ze stockiem
            If Sheets("modu�").Cells(wiersz, kolumnawlacz).Value = "+" Then
                If Sheets("modu�").Cells(wiersz, kolumnaColumnStock).Value = "+" Then
                If Sheets("modu�").Cells(wiersz, kolumnaStockIND) = "" Then MsgBox "Sheet IND. nie okre�lony dla: " & komorkaznazwa, vbInformation
            End If
            End If
        'Sprawdza czy nie ma zaznaczonych dwuch cen
            If Sheets("modu�").Cells(wiersz, kolumnawlacz).Value = "+" Then
                If Sheets("modu�").Cells(wiersz, kolumnaplaceholder20).Value = "+" And Sheets("modu�").Cells(wiersz, kolumnaplaceholder21).Value = "+" _
                Or Sheets("modu�").Cells(wiersz, kolumnaplaceholder20).Value = "+" And Sheets("modu�").Cells(wiersz, kolumnaColumnPrice).Value = "+" _
                Or Sheets("modu�").Cells(wiersz, kolumnaplaceholder21).Value = "+" And Sheets("modu�").Cells(wiersz, kolumnaColumnPrice).Value = "+" _
                Or Sheets("modu�").Cells(wiersz, kolumnaplaceholder20).Value = "+" And Sheets("modu�").Cells(wiersz, kolumnaplaceholder14).Value = "+" _
                Or Sheets("modu�").Cells(wiersz, kolumnaplaceholder21).Value = "+" And Sheets("modu�").Cells(wiersz, kolumnaplaceholder14).Value = "+" _
                Or Sheets("modu�").Cells(wiersz, kolumnaColumnPrice).Value = "+" And Sheets("modu�").Cells(wiersz, kolumnaplaceholder14).Value = "+" _
                Then MsgBox "Cena wprowadzona dwukrotnie dla: " & komorkaznazwa, vbInformation

            End If
            If Sheets("modu�").Cells(wiersz, kolumnawlacz).Value = "+" Then
                If Sheets("modu�").Cells(wiersz, kolumnaColumnPrice).Value = "+" And Sheets("modu�").Cells(wiersz, kolumnplaceholder20IND).Value = "" _
                Then MsgBox "Okre�l kolumn� w Price IND. dla : " & komorkaznazwa, vbInformation
            End If
    
        'Sprawdza czy nie ma zaznaczonych dw�ch stock�w
            If Sheets("modu�").Cells(wiersz, kolumnawlacz).Value = "+" Then
                If Sheets("modu�").Cells(wiersz, kolumnaplaceholder19).Value = "+" And Sheets("modu�").Cells(wiersz, kolumna1Stock).Value = "+" _
                Or Sheets("modu�").Cells(wiersz, kolumnaplaceholder19).Value = "+" And Sheets("modu�").Cells(wiersz, kolumnaL2Stock).Value = "+" _
                Or Sheets("modu�").Cells(wiersz, kolumnaplaceholder19).Value = "+" And Sheets("modu�").Cells(wiersz, kolumnaColumnStock).Value = "+" _
                Or Sheets("modu�").Cells(wiersz, kolumnaplaceholder19).Value = "+" And Sheets("modu�").Cells(wiersz, kolumnaL3Stock).Value = "+" _
                Or Sheets("modu�").Cells(wiersz, kolumnaplaceholder19).Value = "+" And Sheets("modu�").Cells(wiersz, kolumnaplaceholder25).Value = "+" _
                Or Sheets("modu�").Cells(wiersz, kolumna1Stock).Value = "+" And Sheets("modu�").Cells(wiersz, kolumnaL2Stock).Value = "+" _
                Or Sheets("modu�").Cells(wiersz, kolumna1Stock).Value = "+" And Sheets("modu�").Cells(wiersz, kolumnaL3Stock).Value = "+" _
                Or Sheets("modu�").Cells(wiersz, kolumna1Stock).Value = "+" And Sheets("modu�").Cells(wiersz, kolumnaplaceholder25).Value = "+" _
                Or Sheets("modu�").Cells(wiersz, kolumna1Stock).Value = "+" And Sheets("modu�").Cells(wiersz, kolumnaColumnStock).Value = "+" _
                Or Sheets("modu�").Cells(wiersz, kolumnaL2Stock).Value = "+" And Sheets("modu�").Cells(wiersz, kolumnaL3Stock).Value = "+" _
                Or Sheets("modu�").Cells(wiersz, kolumnaL2Stock).Value = "+" And Sheets("modu�").Cells(wiersz, kolumnaplaceholder25).Value = "+" _
                Or Sheets("modu�").Cells(wiersz, kolumnaL2Stock).Value = "+" And Sheets("modu�").Cells(wiersz, kolumnaColumnStock).Value = "+" _
                Or Sheets("modu�").Cells(wiersz, kolumnaL3Stock).Value = "+" And Sheets("modu�").Cells(wiersz, kolumnaplaceholder25).Value = "+" _
                Or Sheets("modu�").Cells(wiersz, kolumnaL3Stock).Value = "+" And Sheets("modu�").Cells(wiersz, kolumnaColumnStock).Value = "+" _
                Or Sheets("modu�").Cells(wiersz, kolumnaColumnStock).Value = "+" And Sheets("modu�").Cells(wiersz, kolumnaplaceholder25).Value = "+" _
                Then MsgBox "Stock wprowadzony kilkukrotnie dla: " & komorkaznazwa, vbInformation
            End If
    Next

    MsgBox "Nie wykryto wi�cej b��d�w: ", vbInformation
End Sub
Sub UsuwanieSheetu()
    
    Application.DisplayAlerts = False
    Sheets("pricelist").Select
    ActiveSheet.Delete
    Sheets("modu�").Select
    Application.DisplayAlerts = True

End Sub
Sub Usuwanie_puste()
'Usuwa wiersze pod pustymi nag��wkami
ow = Cells(Rows.Count, "F").End(xlUp).Row
For r = ow To 5 Step -1
    If Cells(r, "C").Text = "" And Cells(r + 1, "C").Value = "" Then Rows(r).Delete
Next

End Sub
Sub placeholder14()
    
    Sheets("szablon cen").Select
    Range("X:X").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("pricelist").Select
    Columns("F:F").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("F:F").NumberFormat = "#,##0.00 [$�-x-euro1]"
    Columns("F:H").ColumnWidth = 8
    Range("f4").Font.Bold = True
    Range("f4").Value = "price"
    
End Sub
Sub UsuwanienNaglowkow()
' Usuwa wszystkie nag��wki
ow = Cells(Rows.Count, "C").End(xlUp).Row
For r = ow To 4 Step -1
    If Cells(r, "C") = "placeholder5 Supplies" Or Cells(r, "C").Value = "Original placeholder6 print supplies" _
    Or Cells(r, "C").Value = "placeholder27 Original Spare Parts " Or Cells(r, "C").Value = "PLEASE NOTE: Prices and leadtime are to be confirmed before ordering" _
    Or Cells(r, "C").Value = "Spare parts are NON CANCELLABLE" Or Cells(r, "C").Value = "placeholder27 Wide Format" _
    Or Cells(r, "C").Value = "placeholder27 placeholder29 models" Or Cells(r, "C").Value = "placeholder27 B&W" _
    Or Cells(r, "C").Value = "Original placeholder18 print supplies" Or Cells(r, "C").Value = "Original placeholder4 print supplies" _
    Or Cells(r, "C").Value = "placeholder27 Color Machines" Or Cells(r, "C").Value = "Original placeholder17 print supplies" _
    Or Cells(r, "C").Value = "placeholder10 Original Spare Parts " Or Cells(r, "C").Value = "placeholder10 Print Supplies" Or Cells(r, "C").Value = "placeholder11 Supplies" _
    Or Cells(r, "C").Value = "Original placeholder15 Supplies" Or Cells(r, "C").Value = "Original placeholder8 print supplies" _
    Or Cells(r, "C").Value = "Original placeholder7 print supplies" Or Cells(r, "C").Value = "Original placeholder9 print supplies" _
    Or Cells(r, "C").Value = "placeholder12 inkjets" Or Cells(r, "C").Value = "placeholder12 Large Format Supplies" _
    Or Cells(r, "C").Value = "placeholder12 B&W Laser Supplies" Or Cells(r, "C").Value = "placeholder12 Color Laser Supplies" _
    Or Cells(r, "C").Value = "placeholder12 B&W Copier Supplies" Or Cells(r, "C").Value = "placeholder12 Color Copier Supplies" Then Rows(r).Delete
Next

End Sub
Sub UsuwaniePustegoStocku()
    ow = Cells(Rows.Count, "C").End(xlUp).Row
    For r = ow To 5 Step -1
         If Cells(r, "G").Value = "" Then Rows(r).Delete
    Next
End Sub

Sub Usuwaplaceholder12placeholder28()
ow = Cells(Rows.Count, "A").End(xlUp).Row
For r = ow To 1 Step -1
    If Cells(r, "A").Value = "placeholder12" And Cells(r, "D") Like "*placeholder28*" Then Rows(r).Delete
Next

End Sub

Sub Usuwanieplaceholder4NieRibbon()
ow = Cells(Rows.Count, "A").End(xlUp).Row
For r = ow To 1 Step -1
    If Cells(r, "A").Text = "placeholder4" And Not Cells(r, "D") Like "*Ribbon*" Then Rows(r).Delete
Next
End Sub
Sub ZostawiaTylkoStockplaceholder6()
ow = Cells(Rows.Count, "A").End(xlUp).Row
For r = ow To 6 Step -1

    If Not Cells(r, "A").Text = "placeholder6" Then Rows(r).Delete

Next
Range("C5").Value = "Original placeholder6 print supplies"
For r = ow To 6 Step -1
    If Not Cells(r, "G").Value <> "" Then Rows(r).Delete
Next
For r = ow To 6 Step -1
    If Cells(r, "G").Value > 100 Then Cells(r, "G").Value = 100
Next

End Sub
Sub Usuwaplaceholder13()
    ow = Cells(Rows.Count, "A").End(xlUp).Row
    'Znajduje kom�rk� z placeholder27 placeholder29 models i usuwa wszystko co jest pod ni� a� do znalezienia pustej kom�rki
    For r = 1 To 3000
        If Cells(r, "C").Value = "placeholder27 placeholder29 models" And Cells(r + 1, "A") <> "" Then
        Rows(r + 1).Delete
        r = r - 1
        End If
    Next
    'Usuwa nag��wek placeholder27 placeholder29 models i pust� kom�rk� pod nim
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "placeholder27 placeholder29 models" Then Rows(r).Resize(2).Delete
    Next
    ' Usuwa pozycje bez stocku dla placeholder12a placeholder4a i placeholder10
    For r = ow To 1 Step -1
            If Cells(r, "G").Value = "" And Cells(r, "A").Value = "placeholder12" Then Rows(r).Delete
    Next
    For r = ow To 1 Step -1
            If Cells(r, "G").Value = "" And Cells(r, "A").Value = "placeholder4" Then Rows(r).Delete
    Next
    For r = ow To 1 Step -1
            If Cells(r, "G").Value = "" And Cells(r, "A").Value = "placeholder10" Then Rows(r).Delete
    Next
    
    For r = ow To 1 Step -1
        If Cells(r, "C").Text = "Original placeholder4 print supplies" And Cells(r + 1, "C").Value = "" Then Rows(r).Resize(2).Delete
    Next

End Sub
Sub OgraniczDo100()


' Ogranicza stock do warto�ci 100
ow = Cells(Rows.Count, "A").End(xlUp).Row
For r = ow To 6 Step -1
    If Cells(r, "G").Value > 100 Then Cells(r, "G").Value = 100
Next

End Sub
Sub czy_sheetpricelist()


End Sub

