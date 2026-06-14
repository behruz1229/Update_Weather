Option Explicit

Sub UpdateWeatherTable()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim minDate As Date, maxDate As Date, curDate As Date
    Dim datesDict As Object
    Dim vKey As Variant
    Dim url As String
    Dim ie As Object
    Dim html As Object
    Dim targetBlock As Object
    Dim tempCell As Range, weatherCell As Range, windCell As Range, formulaCell As Range
    Dim tempVal As String, weatherVal As String, windVal As String
    Dim updatedCount As Long
    Dim totalDates As Long, processed As Long
    Dim foundRf As Boolean

    ' ---- Настройки ----
    Set ws = ActiveWorkbook.Sheets("Лист1")
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = False
    ie.Silent = True

    Dim oldStatusBar As String
    oldStatusBar = Application.StatusBar
    Application.StatusBar = "Инициализация... (поиск дат)"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' ---- Собираем все существующие даты ----
    Set datesDict = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Нет данных в столбце A (начиная со строки 2)", vbExclamation
        GoTo CleanUp
    End If

    minDate = ws.Cells(2, "A").Value
    maxDate = ws.Cells(lastRow, "A").Value
    For i = 2 To lastRow
        curDate = ws.Cells(i, "A").Value
        datesDict(curDate) = i
        If curDate < minDate Then minDate = curDate
        If curDate > maxDate Then maxDate = curDate
    Next i

    ' ---- Определяем вчерашний день ----
    Dim yesterday As Date
    yesterday = Date - 1

    If minDate > yesterday Then
        MsgBox "Минимальная дата в таблице (" & minDate & ") позже вчерашнего дня. Заполнение не требуется.", vbInformation
        GoTo CleanUp
    End If

    ' ---- Вставляем пропущенные даты (от minDate до yesterday) ----
    Dim current As Date
    Dim newRowIndex As Long

    For current = minDate To yesterday
        If Not datesDict.exists(current) Then
            newRowIndex = 0
            For i = 2 To lastRow
                If ws.Cells(i, "A").Value > current Then
                    newRowIndex = i
                    Exit For
                End If
            Next i
            If newRowIndex = 0 Then newRowIndex = lastRow + 1

            ws.Rows(newRowIndex).Insert Shift:=xlDown
            ws.Cells(newRowIndex, "A").Value = current
            datesDict(current) = newRowIndex
            For Each vKey In datesDict.Keys
                If datesDict(vKey) >= newRowIndex And vKey <> current Then
                    datesDict(vKey) = datesDict(vKey) + 1
                End If
            Next vKey
            lastRow = lastRow + 1
        End If
    Next current

    ' ---- Общее количество дат, которые нужно обработать (до yesterday) ----
    totalDates = 0
    For Each vKey In datesDict.Keys
        If CDate(vKey) <= yesterday Then totalDates = totalDates + 1
    Next vKey
    processed = 0
    updatedCount = 0

    ' ---- Обновляем данные для каждой даты (только пустые ячейки) ----
    Dim rowNum As Long
    For Each vKey In datesDict.Keys
        curDate = CDate(vKey)
        If curDate <= yesterday Then
            processed = processed + 1
            Application.StatusBar = "Обработка: " & Format(curDate, "dd.mm.yyyy") & " (" & processed & " из " & totalDates & ")"

            rowNum = datesDict(vKey)
            Set tempCell = ws.Cells(rowNum, "B")
            Set weatherCell = ws.Cells(rowNum, "C")
            Set windCell = ws.Cells(rowNum, "D")

            If Not (IsEmpty(tempCell) Or IsEmpty(weatherCell) Or IsEmpty(windCell)) Then GoTo NextDate

            url = BuildUrl(curDate)
            If url = "" Then GoTo NextDate

            ie.Navigate url
            Do While ie.Busy Or ie.readyState <> 4
                DoEvents
            Loop

            Application.Wait (Now + TimeValue("0:00:02"))

            Set html = ie.Document

            ' Ищем блок rf, содержащий заголовок с нужной датой
            foundRf = FindRfBlockByDate(html, curDate, targetBlock)
            If foundRf Then
                ' Внутри найденного блока ищем 12:00
                Set targetBlock = Find12OClockInBlock(targetBlock)
                If Not targetBlock Is Nothing Then
                    tempVal = ExtractTemperature(targetBlock)
                    weatherVal = ExtractWeather(targetBlock)
                    windVal = ExtractWind(targetBlock)

                    If IsEmpty(tempCell) And tempVal <> "" Then
                        tempCell.Value = tempVal
                        updatedCount = updatedCount + 1
                    End If
                    If IsEmpty(weatherCell) And weatherVal <> "" Then
                        weatherCell.Value = weatherVal
                        updatedCount = updatedCount + 1
                    End If
                    If IsEmpty(windCell) And windVal <> "" Then
                        windCell.Value = windVal
                        updatedCount = updatedCount + 1
                    End If
                Else
                    ' Блок 12:00 не найден (возможно, данные за этот день ещё не полные)
                End If
            Else
                ' Блок rf с этой датой не найден на странице
            End If
        End If
NextDate:
    Next vKey

    ' ---- Вставляем формулу в столбец E, если пусто (русская функция) ----
    For i = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        Set formulaCell = ws.Cells(i, "E")
        If IsEmpty(formulaCell) Then
            formulaCell.FormulaLocal = "=СЦЕПИТЬ(B" & i & "; "" °С, ""; СИМВОЛ(10); C" & i & "; "", ""; СИМВОЛ(10); D" & i & "; "" м/с"")"
        End If
    Next i

    ' ---- Применяем тонкие границы ----
    Dim dataRange As Range
    Dim lastDataRow As Long
    lastDataRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastDataRow >= 2 Then
        Set dataRange = ws.Range("A2:E" & lastDataRow)
        With dataRange.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End If

CleanUp:
    On Error Resume Next
    ie.Quit
    Set ie = Nothing
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False   ' <--- возвращаем стандартный статус-бар

    If updatedCount > 0 Then
        MsgBox "Обновление завершено!" & vbCrLf & _
               "Обновлено ячеек: " & updatedCount & vbCrLf & _
               "Тонкие границы применены ко всем строкам с данными.", vbInformation
    Else
        MsgBox "Обновление завершено! Новых данных не добавлено.", vbInformation
    End If
End Sub

' ============================================================
'  Вспомогательные функции
' ============================================================

' Построение URL по дате
Function BuildUrl(d As Date) As String
    Dim dayPart As String
    Dim monthPart As String
    Dim yearPart As String
    Dim monthNames(1 To 12) As String

    monthNames(1) = "января"
    monthNames(2) = "февраля"
    monthNames(3) = "марта"
    monthNames(4) = "апреля"
    monthNames(5) = "мая"
    monthNames(6) = "июня"
    monthNames(7) = "июля"
    monthNames(8) = "августа"
    monthNames(9) = "сентября"
    monthNames(10) = "октября"
    monthNames(11) = "ноября"
    monthNames(12) = "декабря"

    dayPart = Format(d, "dd")
    If Left(dayPart, 1) = "0" Then dayPart = Right(dayPart, 1)
    monthPart = monthNames(month(d))
    yearPart = year(d)

    BuildUrl = "https://ust-luga.nuipogoda.ru/" & dayPart & "-" & monthPart & "#" & yearPart
End Function

' Ищет блок div.rf, внутри которого есть h1.cg с датой, соответствующей d.
' Возвращает True и устанавливает rfBlock, если найден.
Function FindRfBlockByDate(html As Object, d As Date, ByRef rfBlock As Object) As Boolean
    Dim rfDivs As Object
    Dim rf As Object
    Dim h1 As Object
    Dim titleText As String
    Dim pageDate As Date
    
    Set rfDivs = html.getElementsByClassName("rf")
    For Each rf In rfDivs
        Set h1 = rf.getElementsByTagName("H1")(0)
        If Not h1 Is Nothing Then
            If h1.className = "cg" Then
                titleText = h1.innerText
                pageDate = ParseDateFromHeader(titleText)
                If pageDate = d Then
                    Set rfBlock = rf
                    FindRfBlockByDate = True
                    Exit Function
                End If
            End If
        End If
    Next rf
    Set rfBlock = Nothing
    FindRfBlockByDate = False
End Function

' Парсит дату из строки заголовка, например "Ср, 12 июня 2024"
Function ParseDateFromHeader(headerText As String) As Date
    Dim parts() As String
    Dim day As Integer, month As Integer, year As Integer
    Dim monthNames As Variant
    Dim i As Integer
    
    parts = Split(headerText, " ")
    If UBound(parts) < 3 Then
        ParseDateFromHeader = 0
        Exit Function
    End If
    
    On Error Resume Next
    day = CInt(parts(1))
    year = CInt(parts(3))
    If Err.Number <> 0 Then
        ParseDateFromHeader = 0
        Exit Function
    End If
    On Error GoTo 0
    
    monthNames = Array("января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря")
    month = 0
    For i = 0 To 11
        If monthNames(i) = parts(2) Then
            month = i + 1
            Exit For
        End If
    Next i
    If month = 0 Then
        ParseDateFromHeader = 0
        Exit Function
    End If
    
    ParseDateFromHeader = DateSerial(year, month, day)
End Function

' Ищет внутри блока rf (или любого другого контейнера) div с временем 12:00
Function Find12OClockInBlock(container As Object) As Object
    Dim timeDivs As Object
    Dim timeDiv As Object
    
    Set timeDivs = container.getElementsByClassName("gh")
    For Each timeDiv In timeDivs
        If Trim(timeDiv.innerText) = "12:00" Then
            Set Find12OClockInBlock = timeDiv.ParentNode
            Exit Function
        End If
    Next timeDiv
    
    ' Альтернативный поиск, если класс gh не подошёл
    Set timeDivs = container.getElementsByTagName("DIV")
    For Each timeDiv In timeDivs
        If timeDiv.className = "gh" And InStr(timeDiv.innerText, "12:00") > 0 Then
            Set Find12OClockInBlock = timeDiv.ParentNode
            Exit Function
        End If
    Next timeDiv
    
    Set Find12OClockInBlock = Nothing
End Function

Function ExtractTemperature(block As Object) As String
    Dim ihDiv As Object
    Set ihDiv = block.getElementsByClassName("ih")(0)
    If Not ihDiv Is Nothing Then
        ExtractTemperature = Replace(ihDiv.innerText, "°", "")
    Else
        ExtractTemperature = ""
    End If
End Function

Function ExtractWeather(block As Object) As String
    Dim hhDiv As Object, spanA As Object
    Set hhDiv = block.getElementsByClassName("hh")(0)
    If Not hhDiv Is Nothing Then
        Set spanA = hhDiv.getElementsByTagName("SPAN")(0)
        If Not spanA Is Nothing Then
            Dim rawText As String
            rawText = spanA.innerText
            If Len(rawText) > 0 Then
                ExtractWeather = UCase(Left(rawText, 1)) & LCase(Mid(rawText, 2))
            Else
                ExtractWeather = ""
            End If
        Else
            ExtractWeather = ""
        End If
    Else
        ExtractWeather = ""
    End If
End Function

Function ExtractWind(block As Object) As String
    Dim khDiv As Object, neDivs As Object, ne As Object
    Dim windVal As String
    Set khDiv = block.getElementsByClassName("kh")(0)
    If Not khDiv Is Nothing Then
        Set neDivs = khDiv.getElementsByClassName("ne")
        For Each ne In neDivs
            If InStr(ne.className, "gg") = 0 And InStr(ne.className, "hg") = 0 Then
                windVal = Trim(ne.innerText)
                If IsNumeric(windVal) Then
                    ExtractWind = windVal
                    Exit Function
                End If
            End If
        Next ne
        If InStr(khDiv.innerText, "штиль") > 0 Then
            ExtractWind = "0"
        Else
            ExtractWind = ""
        End If
    Else
        ExtractWind = ""
    End If
End Function

