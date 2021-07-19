Attribute VB_Name = "automat"
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
    
    'wy³¹czenie aktualizacji widoku okna excal podczas trwania makra
    Application.ScreenUpdating = False
    
    'Aby zapobiec b³êdom usuwa sheet o nazwie pricelist jeœli taki istnieje
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("pricelist").Delete
    Application.DisplayAlerts = True
    
    Dim StartTime As Double
    Dim EndTime As String
    
    
    'zmienna zapisuje czas rozpoczecia dzialania makra
    StartTime = Timer
    'Œcie¿ka zapisu plików
    Sciezka = "C:\cennikiautomat" & "\"
        
    'Sta³e odpowiadaj¹ce numerom konkretnych kolumn w sheecie modu³
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
       
    
    'zmienna która odpowiada ka¿demu wierszowi z nazw¹ firmy (najpierw zlciza ich iloœæ)
    r = Cells(Rows.Count, "D").End(xlUp).Row
    
    'Pêtla która wykonuje siê dla ka¿dej firmy
    For wiersz = 3 To r
    
    'Warunek sprawdzaj¹cy czy makro dla konkretnej firmy ma byæ uruchomione
    If Sheets("modu³").Cells(wiersz, kolumnawlacz).Value = "+" Then
        
    'Zmienna pobieraj¹ca nazwê sheetu indywidualnego z którego maj¹ byæ usuniête konkretne produkty
    nazwasheetuIND = Sheets("modu³").Cells(wiersz, kolumnaSheetIND).Value
    
    'Zmienna pobieraj¹ca nazwê firmy
    komorkaznazwa = Sheets("modu³").Cells(wiersz, kolumnaNazwaFirmy)
    
'W przypadku b³êdu uruchamia prplaceholder10durê obs³ugi b³êdów
On Error GoTo ErrHandler
        
        
    'Makro kolejno sprawdza czy zaznaczone s¹ komórki w sheecie modu³. W przypadk uzaznaczenia ich znakiem "+" uruchamia kolejno wybrane prplaceholder10dury.
    'Po ich zakoñczeniu wszystko zaczyna siê ponownie dla nastêpnej firmy.
        Call kopiowanie_SL
        If Sheets("modu³").Cells(wiersz, kolumnaplaceholder19).Value = "+" Then Call placeholder19
        If Sheets("modu³").Cells(wiersz, kolumna1Stock).Value = "+" Then Call placeholder22
        If Sheets("modu³").Cells(wiersz, kolumnaL2Stock).Value = "+" Then Call placeholder23
        If Sheets("modu³").Cells(wiersz, kolumnaL3Stock).Value = "+" Then Call placeholder24
        If Sheets("modu³").Cells(wiersz, kolumnaplaceholder25).Value = "+" Then Call placeholder25
        If Sheets("modu³").Cells(wiersz, kolumnaColumnStock).Value = "+" Then Call ColumnStock
        If Sheets("modu³").Cells(wiersz, kolumnaplaceholder20).Value = "+" Then Call placeholder20
        If Sheets("modu³").Cells(wiersz, kolumnaplaceholder21).Value = "+" Then Call placeholder21
        If Sheets("modu³").Cells(wiersz, kolumnaColumnPrice).Value = "+" Then Call ColumnPrice
        If Sheets("modu³").Cells(wiersz, kolumnaplaceholder1).Value = "+" Then Call Usuwaplaceholder1
        If Sheets("modu³").Cells(wiersz, kolumnaplaceholder2).Value = "+" Then Call Usuwaplaceholder2
        If Sheets("modu³").Cells(wiersz, kolumnaplaceholder3).Value = "+" Then Call Usuwaplaceholder3
        If Sheets("modu³").Cells(wiersz, kolumnaplaceholder4Ribbon).Value = "+" Then Call Usuwaplaceholder4Ribbon
        If Sheets("modu³").Cells(wiersz, kolumnaplaceholder5).Value = "+" Then Call Usuwaplaceholder5
        If Sheets("modu³").Cells(wiersz, kolumnaplaceholder12).Value = "+" Then Call Usuwaplaceholder12
        If Sheets("modu³").Cells(wiersz, kolumnaplaceholder4).Value = "+" Then Call Usuwaplaceholder4
        If Sheets("modu³").Cells(wiersz, kolumnaplaceholder6).Value = "+" Then Call Usuwaplaceholder6
        If Sheets("modu³").Cells(wiersz, kolumnaplaceholder16).Value = "+" Then Call Usuwaplaceholder16
        If Sheets("modu³").Cells(wiersz, kolumnaplaceholder7).Value = "+" Then Call Usuwaplaceholder7
        If Sheets("modu³").Cells(wiersz, kolumnaplaceholder8).Value = "+" Then Call Usuwaplaceholder8
        If Sheets("modu³").Cells(wiersz, kolumnaplaceholder10).Value = "+" Then Call Usuwaplaceholder10
        If Sheets("modu³").Cells(wiersz, kolumnaplaceholder11).Value = "+" Then Call Usuwaplaceholder11
        If Sheets("modu³").Cells(wiersz, kolumnaplaceholder15).Value = "+" Then Call Usuwaplaceholder15
        If Sheets("modu³").Cells(wiersz, kolumnaplaceholder17).Value = "+" Then Call Usuwaplaceholder17
        If Sheets("modu³").Cells(wiersz, kolumnaplaceholder18).Value = "+" Then Call Usuwaplaceholder18
        If Sheets("modu³").Cells(wiersz, kolumnaplaceholder27).Value = "+" Then Call Usuwaplaceholder27
        If Sheets("modu³").Cells(wiersz, kolumnaCzySheetIND).Value = "+" Then Call UsuwaIND
        If Sheets("modu³").Cells(wiersz, kolumnaplaceholder4NiedoPL).Value = "+" Then Call Usuwaplaceholder4NiedoPL
        If Sheets("modu³").Cells(wiersz, kolumnaplaceholder14).Value = "+" Then Call placeholder14
        If Sheets("modu³").Cells(wiersz, kolumnaUsuwanieNaglowkow).Value = "+" Then Call UsuwanienNaglowkow
        If Sheets("modu³").Cells(wiersz, kolumnaUsuwaniePustegoStocku).Value = "+" Then Call UsuwaniePustegoStocku
        If Sheets("modu³").Cells(wiersz, kolumnaUsuwanieplaceholder12placeholder28).Value = "+" Then Call Usuwaplaceholder12placeholder28
        If Sheets("modu³").Cells(wiersz, kolumnaUsuwanieplaceholder4NieRibbon).Value = "+" Then Call Usuwanieplaceholder4NieRibbon
        If Sheets("modu³").Cells(wiersz, kolumnaZostawiaTylkoStockplaceholder6).Value = "+" Then Call ZostawiaTylkoStockplaceholder6
        If Sheets("modu³").Cells(wiersz, kolumnaUsuwaplaceholder13).Value = "+" Then Call Usuwaplaceholder13
        If Sheets("modu³").Cells(wiersz, kolumnaStockDo100).Value = "+" Then Call OgraniczDo100
        Call Usuwanie_puste
        If Sheets("modu³").Cells(wiersz, kolumnaZapisXLS).Value = "+" Then Call ZapisXLS
        If Sheets("modu³").Cells(wiersz, kolumnaZapisCSV).Value = "+" Then Call ZapisCSV
    
    End If
    
Next
    
    'w³¹czenie aktualizacji widoku okna excal podczas trwania makra
    Application.ScreenUpdating = True
    EndTime = Format((Timer - StartTime) / 86400, "hh:mm:ss")
    
    ' Wiadomoœæ po zakoñczeniu makra wraz z czasem jaki zajê³o jego wykonanie
    MsgBox "Gotowe. Czas trwania: " & EndTime & " minut", vbInformation


Exit Sub

'Obs³uga b³êdów
ErrHandler:
    'Je¿eli zosta³ ju¿ utworzony sheet "pricelist" prplaceholder10dura go usunie aby nie wywo³ywaæ dalszych b³êdów
    If ActiveSheet.Name = "pricelist" Then ActiveSheet.Delete
    'Wyœwietla info dla jakiego cennika wyst¹pi³ b³¹d
    MsgBox "Wyst¹pi³ bl¹d dla cennika: " & vbNewLine & komorkaznazwa & vbNewLine & "SprawdŸ poprawnoœæ wprowadzonych danych lub skontaktuj siê z administratoerm.", vbInformation
    'Po wyœwietleniu zostanie uruchomione makro dla kolejnego cennika tak aby tylko ten na którym wyst¹pi³ b³¹d nie zsota³ stworzony
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
    'Ustawienie odpowiedniego formatu walutowego, pogrubienie czcionki, i ustawienie szerokoœci kolumny
    Columns("F:F").NumberFormat = "#,##0.00 [$€-x-euro1]"
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
    'Ustawienie odpowiedniego formatu walutowego, pogrubienie czcionki, i ustawienie szerokoœci kolumny
    Columns("F:F").NumberFormat = "#,##0.00 [$€-x-euro1]"
    Columns("F:H").ColumnWidth = 8
    Range("f4").Font.Bold = True
    Range("f4").Value = "price"

End Sub
Sub ColumnPrice()

    'Kopiowanie cen z kolumny wpisanej do kolumny "Wpisz symbol kolumny" i umieszczanie ich w arkuszu z nowym cennikiem
    kolumnplaceholder20IND = 14
    kolumnaColumnPrice = 13
    nrkolumny = Sheets("modu³").Cells(wiersz, kolumnplaceholder20IND).Value

    Sheets("szablon cen").Select
    Range(Cells(1, nrkolumny), Cells(5000, nrkolumny)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("pricelist").Select
    Columns("F:F").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Ustawienie odpowiedniego formatu walutowego, pogrubienie czcionki, i ustawienie szerokoœci kolumny
    Columns("F:F").NumberFormat = "#,##0.00 [$€-x-euro1]"
    Columns("F:H").ColumnWidth = 8
    Range("f4").Font.Bold = True
    Range("f4").Value = "price"
    
    ' Usuwa wiersz je¿eli w komórce z cen¹ jest wpsiane "x"
    ow = Cells(Rows.Count, "F").End(xlUp).Row
    For r = ow To 1 Step -1
        If Cells(r, "F").Value = "x" Then Rows(r).Delete
    Next
    
End Sub
Sub ColumnStock()

    'Kopiowanie stocku z kolumny wpisanej do kolumny "Wpisz symbol kolumny" i umieszczanie ich arkuszu z nowym cennikiem
    kolumnaStockIND = 51
    kolumnaColumnStock = 50
    nrkolumny = Sheets("modu³").Cells(wiersz, kolumnaStockIND).Value

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
    
    'Kopiowanie zwyk³ego stocku i umieszczanie go w arkuszu z nowym cennikiem
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
    
    'Kopiowanie Stocku zwyk³ego z ... i umieszczanie go w arkuszu z nowym cennikiem
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
    
    'Kopiowanie Stocku zwyk³ego z placeholder7 N i umieszczanie go w arkuszu z nowym cennikiem
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
    
    'Usuwanie wszystkich placeholder18ów Spare Parts
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
    'Usuwa markê placeholder5 wraz za nag³ówkiem/nag³ókami
    Sheets("pricelist").Select
    ow = Cells(Rows.Count, "A").End(xlUp).Row
    'Usuwa placeholder5
    For r = ow To 1 Step -1
        If Cells(r, "A").Value = "placeholder5" Then Rows(r).Delete
    Next
    ' Usuwa nag³ówek placeholder5
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "placeholder5 Supplies" Then Rows(r).Resize(2).Delete
    Next
End Sub
Sub Usuwaplaceholder12()
    
    'Usuwa markê placeholder12 wraz za nag³ówkiem/nag³ókami
    Sheets("pricelist").Select
    ow = Cells(Rows.Count, "A").End(xlUp).Row
    'Usuwa placeholder12
    For r = ow To 1 Step -1
        If Cells(r, "A").Value = "placeholder12" Then Rows(r).Delete
    Next
    ' Usuwa nag³ówki placeholder12
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
    
    'Usuwa markê placeholder4 wraz za nag³ówkiem/nag³ókami
    Sheets("pricelist").Select
    ow = Cells(Rows.Count, "A").End(xlUp).Row
    'Usuwa placeholder4
    For r = ow To 1 Step -1
        If Cells(r, "A").Value = "placeholder4" Then Rows(r).Delete
    Next
    ' Usuwa nag³ówek placeholder4
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "Original placeholder4 print supplies" Then Rows(r).Resize(2).Delete
    Next

End Sub
Sub Usuwaplaceholder6()
    
    'Usuwa markê placeholder6 wraz za nag³ówkiem/nag³ókami
    Sheets("pricelist").Select
    ow = Cells(Rows.Count, "A").End(xlUp).Row
    'Usuwa placeholder6
    For r = ow To 1 Step -1
        If Cells(r, "A").Value = "placeholder6" Then Rows(r).Delete
    Next
    ' Usuwa nag³ówek placeholder6
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "Original placeholder6 print supplies" Then Rows(r).Resize(2).Delete
    Next

End Sub
Sub Usuwaplaceholder16()
    
    'Usuwa markê placeholder9 wraz za nag³ówkiem/nag³ókami
    Sheets("pricelist").Select
    ow = Cells(Rows.Count, "A").End(xlUp).Row
    'Usuwa placeholder16
    For r = ow To 1 Step -1
        If Cells(r, "A").Value = "placeholder9" Then Rows(r).Delete
    Next
    ' Usuwa nag³ówek placeholder16
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "Original placeholder9 print supplies" Then Rows(r).Resize(2).Delete
    Next

End Sub
Sub Usuwaplaceholder7()

    'Usuwa markê placeholder7 wraz za nag³ówkiem/nag³ókami
    Sheets("pricelist").Select
    ow = Cells(Rows.Count, "A").End(xlUp).Row
    'Usuwa placeholder7
    For r = ow To 1 Step -1
        If Cells(r, "A").Value = "placeholder7" Then Rows(r).Delete
    Next
    ' Usuwa nag³ówek placeholder7
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "Original placeholder7 print supplies" Then Rows(r).Resize(2).Delete
    Next

End Sub
Sub Usuwaplaceholder8()

    'Usuwa markê placeholder8 wraz za nag³ówkiem/nag³ókami
    Sheets("pricelist").Select
    ow = Cells(Rows.Count, "A").End(xlUp).Row
    'Usuwa placeholder8
    For r = ow To 1 Step -1
        If Cells(r, "A").Value = "placeholder8" Then Rows(r).Delete
    Next
    ' Usuwa nag³ówek placeholder8
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "Original placeholder8 print supplies" Then Rows(r).Resize(2).Delete
    Next

End Sub
Sub Usuwaplaceholder10()

    'Usuwa markê placeholder10 wraz za nag³ówkiem/nag³ókami
    Sheets("pricelist").Select
    ow = Cells(Rows.Count, "A").End(xlUp).Row
    'Usuwa placeholder10
    For r = ow To 1 Step -1
        If Cells(r, "A").Value = "placeholder10" Then Rows(r).Delete
    Next
    ' Usuwa nag³ówek placeholder10
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "placeholder10 Print Supplies" Then Rows(r).Resize(2).Delete
    Next

End Sub
Sub Usuwaplaceholder11()

    'Usuwa markê placeholder11 wraz za nag³ówkiem/nag³ókami
    Sheets("pricelist").Select
    ow = Cells(Rows.Count, "A").End(xlUp).Row
    'Usuwa placeholder11
    For r = ow To 1 Step -1
        If Cells(r, "A").Value = "placeholder11" Then Rows(r).Delete
    Next
    ' Usuwa nag³ówek placeholder11
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "placeholder11 Supplies" Then Rows(r).Resize(2).Delete
    Next

End Sub
Sub Usuwaplaceholder15()

    'Usuwa markê placeholder15 wraz za nag³ówkiem/nag³ókami
    Sheets("pricelist").Select
    ow = Cells(Rows.Count, "A").End(xlUp).Row
    'Usuwa placeholder15
    For r = ow To 1 Step -1
        If Cells(r, "A").Value = "placeholder15" Then Rows(r).Delete
    Next
    ' Usuwa nag³ówek placeholder15
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "Original placeholder15 Supplies" Then Rows(r).Resize(2).Delete
    Next

End Sub
Sub Usuwaplaceholder17()

    'Usuwa markê placeholder17 wraz za nag³ówkiem/nag³ókami
    Sheets("pricelist").Select
    ow = Cells(Rows.Count, "A").End(xlUp).Row
    'Usuwa placeholder17
    For r = ow To 1 Step -1
        If Cells(r, "A").Value = "placeholder17" Then Rows(r).Delete
    Next
    ' Usuwa nag³ówek placeholder17
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "Original placeholder17 print supplies" Then Rows(r).Resize(2).Delete
    Next

End Sub
Sub Usuwaplaceholder18()

    'Usuwa markê placeholder18 wraz za nag³ówkiem/nag³ókami
    Sheets("pricelist").Select
    ow = Cells(Rows.Count, "A").End(xlUp).Row
    'Usuwa placeholder18
    For r = ow To 1 Step -1
        If Cells(r, "A").Value = "placeholder18" Then Rows(r).Delete
    Next
    ' Usuwa nag³ówek placeholder18
    For r = ow To 1 Step -1
        If Cells(r, "C").Value = "Original placeholder18 print supplies" Then Rows(r).Resize(2).Delete
    Next

End Sub
Sub Usuwaplaceholder27()

    'Usuwa markê placeholder27 wraz za nag³ówkiem/nag³ókami
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
    ' Usuwa nag³ówek placeholder27
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
    Sheets("modu³").Select
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
    Sheets("modu³").Select
End Sub
Sub ZapisCSV()
    Sheets("modu³").Select
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
    Sheets("modu³").Select
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
    
        'Sprawdza poprawnoœæ cen
        komorkaznazwa = Sheets("modu³").Cells(wiersz, kolumnaNazwaFirmy)
            If Sheets("modu³").Cells(wiersz, kolumnawlacz).Value = "+" Then
                If Sheets("modu³").Cells(wiersz, kolumnaplaceholder20).Value = "+" Then
                ElseIf Sheets("modu³").Cells(wiersz, kolumnaplaceholder21).Value = "+" Then
                ElseIf Sheets("modu³").Cells(wiersz, kolumnaColumnPrice).Value = "+" Then
                ElseIf Sheets("modu³").Cells(wiersz, kolumnaplaceholder14).Value = "+" Then
                Else: MsgBox "Ceny nie okreœlone dla: " & komorkaznazwa, vbInformation
            End If
            End If
        'Sprawdza poparawnoœæ stocków
            If Sheets("modu³").Cells(wiersz, kolumnawlacz).Value = "+" Then
                If Sheets("modu³").Cells(wiersz, kolumnaplaceholder19).Value = "+" Then
                ElseIf Sheets("modu³").Cells(wiersz, kolumna1Stock).Value = "+" Then
                ElseIf Sheets("modu³").Cells(wiersz, kolumnaL2Stock).Value = "+" Then
                ElseIf Sheets("modu³").Cells(wiersz, kolumnaL3Stock).Value = "+" Then
                ElseIf Sheets("modu³").Cells(wiersz, kolumnaplaceholder25).Value = "+" Then
                ElseIf Sheets("modu³").Cells(wiersz, kolumnaColumnStock).Value = "+" Then
                Else: MsgBox "Stock nie okreœlony dla: " & komorkaznazwa, vbInformation
            End If
            End If
        'Sprawdza czy wpisana jest nazwa sheet IND.
            If Sheets("modu³").Cells(wiersz, kolumnawlacz).Value = "+" Then
                If Sheets("modu³").Cells(wiersz, kolumnaCzySheetIND).Value = "+" Then
                If Sheets("modu³").Cells(wiersz, kolumnaSheetIND) = "" Then MsgBox "Sheet IND. nie okreœlony dla: " & komorkaznazwa, vbInformation
            End If
            End If
        'Sprawdza czy wpisana jest kolumna z cenami
            If Sheets("modu³").Cells(wiersz, kolumnawlacz).Value = "+" Then
                If Sheets("modu³").Cells(wiersz, kolumnaColumnPrice).Value = "+" Then
                If Sheets("modu³").Cells(wiersz, kolumnplaceholder20IND) = "" Then MsgBox "Sheet IND. nie okreœlony dla: " & komorkaznazwa, vbInformation
            End If
            End If
        'Sprawdza czy wpisany jest kolumna ze stockiem
            If Sheets("modu³").Cells(wiersz, kolumnawlacz).Value = "+" Then
                If Sheets("modu³").Cells(wiersz, kolumnaColumnStock).Value = "+" Then
                If Sheets("modu³").Cells(wiersz, kolumnaStockIND) = "" Then MsgBox "Sheet IND. nie okreœlony dla: " & komorkaznazwa, vbInformation
            End If
            End If
        'Sprawdza czy nie ma zaznaczonych dwuch cen
            If Sheets("modu³").Cells(wiersz, kolumnawlacz).Value = "+" Then
                If Sheets("modu³").Cells(wiersz, kolumnaplaceholder20).Value = "+" And Sheets("modu³").Cells(wiersz, kolumnaplaceholder21).Value = "+" _
                Or Sheets("modu³").Cells(wiersz, kolumnaplaceholder20).Value = "+" And Sheets("modu³").Cells(wiersz, kolumnaColumnPrice).Value = "+" _
                Or Sheets("modu³").Cells(wiersz, kolumnaplaceholder21).Value = "+" And Sheets("modu³").Cells(wiersz, kolumnaColumnPrice).Value = "+" _
                Or Sheets("modu³").Cells(wiersz, kolumnaplaceholder20).Value = "+" And Sheets("modu³").Cells(wiersz, kolumnaplaceholder14).Value = "+" _
                Or Sheets("modu³").Cells(wiersz, kolumnaplaceholder21).Value = "+" And Sheets("modu³").Cells(wiersz, kolumnaplaceholder14).Value = "+" _
                Or Sheets("modu³").Cells(wiersz, kolumnaColumnPrice).Value = "+" And Sheets("modu³").Cells(wiersz, kolumnaplaceholder14).Value = "+" _
                Then MsgBox "Cena wprowadzona dwukrotnie dla: " & komorkaznazwa, vbInformation

            End If
            If Sheets("modu³").Cells(wiersz, kolumnawlacz).Value = "+" Then
                If Sheets("modu³").Cells(wiersz, kolumnaColumnPrice).Value = "+" And Sheets("modu³").Cells(wiersz, kolumnplaceholder20IND).Value = "" _
                Then MsgBox "Okreœl kolumnê w Price IND. dla : " & komorkaznazwa, vbInformation
            End If
    
        'Sprawdza czy nie ma zaznaczonych dwóch stocków
            If Sheets("modu³").Cells(wiersz, kolumnawlacz).Value = "+" Then
                If Sheets("modu³").Cells(wiersz, kolumnaplaceholder19).Value = "+" And Sheets("modu³").Cells(wiersz, kolumna1Stock).Value = "+" _
                Or Sheets("modu³").Cells(wiersz, kolumnaplaceholder19).Value = "+" And Sheets("modu³").Cells(wiersz, kolumnaL2Stock).Value = "+" _
                Or Sheets("modu³").Cells(wiersz, kolumnaplaceholder19).Value = "+" And Sheets("modu³").Cells(wiersz, kolumnaColumnStock).Value = "+" _
                Or Sheets("modu³").Cells(wiersz, kolumnaplaceholder19).Value = "+" And Sheets("modu³").Cells(wiersz, kolumnaL3Stock).Value = "+" _
                Or Sheets("modu³").Cells(wiersz, kolumnaplaceholder19).Value = "+" And Sheets("modu³").Cells(wiersz, kolumnaplaceholder25).Value = "+" _
                Or Sheets("modu³").Cells(wiersz, kolumna1Stock).Value = "+" And Sheets("modu³").Cells(wiersz, kolumnaL2Stock).Value = "+" _
                Or Sheets("modu³").Cells(wiersz, kolumna1Stock).Value = "+" And Sheets("modu³").Cells(wiersz, kolumnaL3Stock).Value = "+" _
                Or Sheets("modu³").Cells(wiersz, kolumna1Stock).Value = "+" And Sheets("modu³").Cells(wiersz, kolumnaplaceholder25).Value = "+" _
                Or Sheets("modu³").Cells(wiersz, kolumna1Stock).Value = "+" And Sheets("modu³").Cells(wiersz, kolumnaColumnStock).Value = "+" _
                Or Sheets("modu³").Cells(wiersz, kolumnaL2Stock).Value = "+" And Sheets("modu³").Cells(wiersz, kolumnaL3Stock).Value = "+" _
                Or Sheets("modu³").Cells(wiersz, kolumnaL2Stock).Value = "+" And Sheets("modu³").Cells(wiersz, kolumnaplaceholder25).Value = "+" _
                Or Sheets("modu³").Cells(wiersz, kolumnaL2Stock).Value = "+" And Sheets("modu³").Cells(wiersz, kolumnaColumnStock).Value = "+" _
                Or Sheets("modu³").Cells(wiersz, kolumnaL3Stock).Value = "+" And Sheets("modu³").Cells(wiersz, kolumnaplaceholder25).Value = "+" _
                Or Sheets("modu³").Cells(wiersz, kolumnaL3Stock).Value = "+" And Sheets("modu³").Cells(wiersz, kolumnaColumnStock).Value = "+" _
                Or Sheets("modu³").Cells(wiersz, kolumnaColumnStock).Value = "+" And Sheets("modu³").Cells(wiersz, kolumnaplaceholder25).Value = "+" _
                Then MsgBox "Stock wprowadzony kilkukrotnie dla: " & komorkaznazwa, vbInformation
            End If
    Next

    MsgBox "Nie wykryto wiêcej b³êdów: ", vbInformation
End Sub
Sub UsuwanieSheetu()
    
    Application.DisplayAlerts = False
    Sheets("pricelist").Select
    ActiveSheet.Delete
    Sheets("modu³").Select
    Application.DisplayAlerts = True

End Sub
Sub Usuwanie_puste()
'Usuwa wiersze pod pustymi nag³ówkami
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
    Columns("F:F").NumberFormat = "#,##0.00 [$€-x-euro1]"
    Columns("F:H").ColumnWidth = 8
    Range("f4").Font.Bold = True
    Range("f4").Value = "price"
    
End Sub
Sub UsuwanienNaglowkow()
' Usuwa wszystkie nag³ówki
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
    'Znajduje komórkê z placeholder27 placeholder29 models i usuwa wszystko co jest pod ni¹ a¿ do znalezienia pustej komórki
    For r = 1 To 3000
        If Cells(r, "C").Value = "placeholder27 placeholder29 models" And Cells(r + 1, "A") <> "" Then
        Rows(r + 1).Delete
        r = r - 1
        End If
    Next
    'Usuwa nag³ówek placeholder27 placeholder29 models i pust¹ komórkê pod nim
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


' Ogranicza stock do wartoœci 100
ow = Cells(Rows.Count, "A").End(xlUp).Row
For r = ow To 6 Step -1
    If Cells(r, "G").Value > 100 Then Cells(r, "G").Value = 100
Next

End Sub
Sub czy_sheetpricelist()


End Sub

