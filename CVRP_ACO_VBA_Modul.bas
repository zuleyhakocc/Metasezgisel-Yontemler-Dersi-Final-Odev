Option Explicit  ' Degiskenlerin onceden tanimlanmasini zorunlu kilar

' ==========================================================
' Simetrik CVRP - Karinca Kolonisi Optimizasyonu
' Excel VBA Karar Destek Sistemi
'
' Bu sistemde 0 numarali nokta depo olarak kabul edilir.
' Diger noktalar musterileri temsil eder.
'
' Parametre sayfasi : arac kapasitesi ve algoritma ayarlari
' Mesafe sayfasi    : depo ve musteriler arasi mesafe matrisi
' Talep sayfasi     : musterilerin talep miktarlari
' Sonuc sayfasi     : en iyi rota ve toplam maliyet ciktisi
'
' Sistem tahmin yapmaz.
' Girilen mesafe, talep ve kapasite degerlerine gore
' toplam mesafesi en dusuk rota kombinasyonunu arar.
' ==========================================================

Dim Distance() As Double        ' Mesafe matrisini tutar
Dim Demand() As Double          ' Musteri talep degerlerini tutar
Dim Feromon() As Double         ' Yollar uzerindeki feromon degerlerini tutar
Dim Visibility() As Double      ' Mesafenin tersini tutar, 1 / mesafe olarak hesaplanir
Dim EnIyiRota As String         ' Algoritmanin buldugu en iyi rota bilgisini metin olarak tutar
Dim EnIyiMaliyet As Double      ' En iyi rotaya ait toplam mesafe/maliyet degerini tutar

Sub Run_ACO_CVRP()

    Dim n As Long                    ' Nokta sayisini tutar, depo da dahildir
    Dim aracCapacity As Double       ' Aracin tek turda tasiyabilecegi maksimum talep miktari
    Dim karincaCount As Long         ' Algoritmada kullanilacak karinca sayisi
    Dim iterasyonCount As Long       ' Algoritmanin kac kez tekrar calisacagini tutar
    Dim alpha As Double              ' Feromon bilgisinin rota secimindeki etkisi
    Dim beta As Double               ' Mesafe bilgisinin rota secimindeki etkisi
    Dim BuharlasmaRate As Double     ' Her asamada feromonun ne kadar azalacagini tutar
    Dim Q As Double                  ' Iyi bulunan rotalara eklenecek feromon katsayisi

    Dim i As Long, j As Long         ' Dongulerde satir ve sutun takibi icin kullanilir
    Dim iter As Long, ant As Long    ' Iterasyon ve karinca dongulerini takip eder
    Dim rota As String               ' Her karincanin olusturdugu rota bilgisini tutar
    Dim cost As Double               ' Olusturulan rotanin toplam mesafe/maliyet degerini tutar

    Randomize                        ' Rassal secimlerin her calistirmada farkli olmasini saglar

    ' Parametre sayfasindan kullanici girdileri okunur
    aracCapacity = Sheets("Parametre").Range("B2").Value
    karincaCount = Sheets("Parametre").Range("B3").Value
    iterasyonCount = Sheets("Parametre").Range("B4").Value
    alpha = Sheets("Parametre").Range("B5").Value
    beta = Sheets("Parametre").Range("B6").Value
    BuharlasmaRate = Sheets("Parametre").Range("B7").Value
    Q = Sheets("Parametre").Range("B8").Value

    ' Mesafe ve talep verileri okunur
    n = ReadData()

    ' Feromon ve gorunurluk matrisleri olusturulur
    ReDim Feromon(0 To n - 1, 0 To n - 1)
    ReDim Visibility(0 To n - 1, 0 To n - 1)

    For i = 0 To n - 1
        For j = 0 To n - 1

            If i <> j Then
                Feromon(i, j) = 1

                If Distance(i, j) > 0 Then
                    Visibility(i, j) = 1 / Distance(i, j)
                End If
            End If

        Next j
    Next i

    ' Ilk karsilastirma icin en iyi maliyet cok buyuk atanir
    EnIyiMaliyet = 10 ^ 30
    EnIyiRota = ""

    ' ACO ana dongusu
    For iter = 1 To iterasyonCount

        For ant = 1 To karincaCount

            rota = ConstructSolution(n, aracCapacity, alpha, beta, cost)

            If cost < EnIyiMaliyet Then
                EnIyiMaliyet = cost
                EnIyiRota = rota
            End If

            UpdateFeromon rota, cost, Q

        Next ant

        BuharlasmaFeromon n, BuharlasmaRate

    Next iter

    WriteResult EnIyiRota, EnIyiMaliyet

    MsgBox "ACO tamamlandi. En iyi maliyet: " & Format(EnIyiMaliyet, "0.00"), vbInformation

End Sub

Function ReadData() As Long

    Dim wsD As Worksheet
    Dim wsQ As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long

    Set wsD = ThisWorkbook.Worksheets("Mesafe")
    Set wsQ = ThisWorkbook.Worksheets("Talep")

    lastRow = wsD.Cells(wsD.Rows.Count, 1).End(xlUp).Row

    ReDim Distance(0 To lastRow - 1, 0 To lastRow - 1)
    ReDim Demand(0 To lastRow - 1)

    For i = 1 To lastRow
        For j = 1 To lastRow
            Distance(i - 1, j - 1) = CDbl(wsD.Cells(i, j).Value)
        Next j
    Next i

    For i = 1 To lastRow
        Demand(i - 1) = CDbl(wsQ.Cells(i, 1).Value)
    Next i

    ReadData = lastRow

End Function

Function ConstructSolution(n As Long, capacity As Double, alpha As Double, beta As Double, ByRef totalCost As Double) As String

    Dim visited() As Boolean
    Dim currentNode As Long
    Dim nextNode As Long
    Dim load As Double
    Dim visitedCount As Long
    Dim routeText As String

    ReDim visited(0 To n - 1)

    currentNode = 0
    visited(0) = True
    visitedCount = 1
    load = 0
    totalCost = 0
    routeText = "0"

    Do While visitedCount < n

        nextNode = SelectNextNode(currentNode, visited, load, capacity, n, alpha, beta)

        If nextNode = -1 Then
            totalCost = totalCost + Distance(currentNode, 0)
            routeText = routeText & " - 0 | 0"
            currentNode = 0
            load = 0
        Else
            totalCost = totalCost + Distance(currentNode, nextNode)
            routeText = routeText & " - " & nextNode
            currentNode = nextNode
            visited(nextNode) = True
            visitedCount = visitedCount + 1
            load = load + Demand(nextNode)
        End If

    Loop

    If currentNode <> 0 Then
        totalCost = totalCost + Distance(currentNode, 0)
        routeText = routeText & " - 0"
    End If

    ConstructSolution = routeText

End Function

Function SelectNextNode(currentNode As Long, visited() As Boolean, load As Double, capacity As Double, n As Long, alpha As Double, beta As Double) As Long

    Dim probabilities() As Double
    Dim i As Long
    Dim totalProb As Double
    Dim value As Double
    Dim r As Double
    Dim cumulative As Double

    ReDim probabilities(0 To n - 1)
    totalProb = 0

    For i = 1 To n - 1

        If visited(i) = False Then

            If load + Demand(i) <= capacity Then
                value = (Feromon(currentNode, i) ^ alpha) * (Visibility(currentNode, i) ^ beta)
                probabilities(i) = value
                totalProb = totalProb + value
            End If

        End If

    Next i

    If totalProb = 0 Then
        SelectNextNode = -1
        Exit Function
    End If

    r = Rnd() * totalProb
    cumulative = 0

    For i = 1 To n - 1

        cumulative = cumulative + probabilities(i)

        If cumulative >= r Then
            SelectNextNode = i
            Exit Function
        End If

    Next i

    SelectNextNode = -1

End Function

Sub UpdateFeromon(routeText As String, cost As Double, Q As Double)

    Dim parts() As String
    Dim cleanText As String
    Dim i As Long
    Dim fromNode As Long
    Dim toNode As Long
    Dim delta As Double

    If cost <= 0 Then Exit Sub

    delta = Q / cost

    cleanText = Replace(routeText, "|", "-")
    cleanText = Replace(cleanText, " ", "")
    parts = Split(cleanText, "-")

    For i = LBound(parts) To UBound(parts) - 1

        If IsNumeric(parts(i)) And IsNumeric(parts(i + 1)) Then

            fromNode = CLng(parts(i))
            toNode = CLng(parts(i + 1))

            If fromNode <> toNode Then
                Feromon(fromNode, toNode) = Feromon(fromNode, toNode) + delta
                Feromon(toNode, fromNode) = Feromon(toNode, fromNode) + delta
            End If

        End If

    Next i

End Sub

Sub BuharlasmaFeromon(n As Long, BuharlasmaRate As Double)

    Dim i As Long, j As Long

    For i = 0 To n - 1
        For j = 0 To n - 1

            If i <> j Then
                Feromon(i, j) = Feromon(i, j) * (1 - BuharlasmaRate)

                If Feromon(i, j) < 0.0001 Then
                    Feromon(i, j) = 0.0001
                End If
            End If

        Next j
    Next i

End Sub

Sub WriteResult(routeText As String, cost As Double)

    Dim ws As Worksheet
    Dim routeParts() As String
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets("Sonuc")

    ws.Cells.Clear

    ws.Range("A1").Value = "Simetrik CVRP - Karinca Kolonisi Optimizasyonu"
    ws.Range("A1:D1").Merge
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A1:D1").Interior.Color = RGB(31, 58, 95)
    ws.Range("A1:D1").Font.Color = RGB(255, 255, 255)

    ws.Range("A3").Value = "Toplam Mesafe / Maliyet"
    ws.Range("B3").Value = cost

    ws.Range("A4").Value = "Kullanilan Yontem"
    ws.Range("B4").Value = "Karinca Kolonisi Optimizasyonu"

    ws.Range("A5").Value = "Problem Turu"
    ws.Range("B5").Value = "Simetrik CVRP"

    ws.Range("A7").Value = "En Iyi Rota"
    ws.Range("A7").Font.Bold = True

    routeParts = Split(routeText, "|")

    For i = LBound(routeParts) To UBound(routeParts)
        ws.Cells(8 + i, 1).Value = "Arac / Tur " & (i + 1)
        ws.Cells(8 + i, 2).Value = Trim(routeParts(i))
    Next i

    ws.Range("A3:A5").Font.Bold = True
    ws.Columns("A:D").AutoFit

End Sub

Sub Button1_Click()
    Run_ACO_CVRP
End Sub
