' Module_PARSER.bas
' Чистая версия модуля парсера (готово к импорту)
Option Explicit

' -------------------- Глобальные переменные (1..20) --------------------
Public InputString As String
Public WorkString As String

Public ParamName(1 To 20) As String
Public ParamValue(1 To 20) As String
Public ParamStart(1 To 20) As Long
Public ParamEnd(1 To 20) As Long
Public ParamErrorCode(1 To 20) As Integer

' Коды ошибок
Private Const ERR_NOT_DEFINED As Integer = 127
Private Const ERR_FOUND_NOT_CLEANED As Integer = 255
Private Const ERR_OK As Integer = 0

' Префикс лога
Private Const LOG_PREFIX As String = "ParserLog: "
Public Index As Integer
Public DebugString As String

Public Sub Parse_B1()
    'On Error GoTo EH
    Debug.Print "Parse_B1: START"

    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets("1")

    ' 1) Очистка A7:V10 включая форматы
    With sh
        .Range("A7:V10").Clear
        .Range("A7:V7").HorizontalAlignment = xlCenter
        Dim hdr As Variant
        hdr = Array("1", "2", "3", "4", "5", "6", "7.1", "7.2", "8", "9.1", "9.2", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20")
        Dim i As Integer
        For i = 0 To UBound(hdr)
            .Cells(7, i + 1).Value = hdr(i)
            .Cells(7, i + 1).HorizontalAlignment = xlCenter
        Next i
    End With

    ' 2) InputString из B1
    Dim sInput As String
    sInput = Trim$(CStr(sh.Range("B1").Value))

    ' 3) Вызов Parser
    Dim ret As Long
    ret = Parser(sInput)

    ' 4) Если была критическая ошибка (ret <> 0), уведомим
    If ret <> 0 Then
        MsgBox "Parser вернул код ошибки: " & CStr(ret), vbExclamation, "Parse_B1"
    End If

    ' 5) Заполнить A8:V8 - ParamName
    For i = 1 To 20
        sh.Cells(8, i).Value = ParamName(i)
        sh.Cells(8, i).HorizontalAlignment = xlCenter
    Next i

    ' 6) A9:V9 - ParamValue
    For i = 1 To 20
        sh.Cells(9, i).Value = ParamValue(i)
        sh.Cells(9, i).HorizontalAlignment = xlCenter
    Next i

    ' 7) A10:V10 - ParamErrorCode, с подсветкой
    For i = 1 To 20
        sh.Cells(10, i).Value = ParamErrorCode(i)
        sh.Cells(10, i).HorizontalAlignment = xlCenter
        If ParamErrorCode(i) = 0 Then
            sh.Cells(10, i).Interior.Color = RGB(198, 239, 206) ' светло-зелёный
        Else
            sh.Cells(10, i).Interior.Color = RGB(255, 199, 206) ' светло-красный
        End If
    Next i

    Debug.Print "Parse_B1: COMPLETED"
    Exit Sub

EH:
    Debug.Print "Parse_B1 ERROR: " & Err.Number & " - " & Err.Description
    MsgBox "Ошибка в Parse_B1: " & Err.Description, vbCritical
End Sub

' NormalizeGOST: нормализация строки

Public Function NormalizeGOST(ByVal ToNormGOST As String) As String
 Dim TempString As String
 TempString = ToNormGOST
 
 ' 1. Заменяем все "плохие" пробелы на ANSI-пробел
 TempString = Replace(TempString, ChrW(160), " ")
 TempString = Replace(TempString, vbTab, " ")
 TempString = Replace(TempString, ChrW(8194), " ")
 TempString = Replace(TempString, ChrW(8195), " ")
 TempString = Replace(TempString, ChrW(8201), " ")
 TempString = Replace(TempString, ChrW(8202), " ")
 TempString = Replace(TempString, ChrW(8239), " ")
 TempString = Replace(TempString, ChrW(8287), " ")
 TempString = Replace(TempString, ChrW(12288), " ")
 
 ' 2. Заменяем длинные тире на "-"
 TempString = Replace(TempString, ChrW(8211), "-") ' EN DASH
 TempString = Replace(TempString, ChrW(8212), "-") ' EM DASH
 
 ' 3. Убираем пробелы вокруг "/"
 Do While InStr(TempString, " /") > 0 Or InStr(TempString, "/ ") > 0
 TempString = Replace(TempString, " /", "/")
 TempString = Replace(TempString, "/ ", "/")
 Loop
 
 ' 4. Убираем множественные пробелы
 Do While InStr(TempString, "  ") > 0
 TempString = Replace(TempString, "  ", " ")
 Loop
 
 ' 5. Обрезаем пробелы по краям
 TempString = Trim(TempString)
 
 ' 6. Посимвольная замена похожих и по ГОСТ
 Dim ReplacementMap As Object
 Set ReplacementMap = CreateObject("Scripting.Dictionary")
 ReplacementMap.Add "К", "K"
 ReplacementMap.Add "Е", "E"
 ReplacementMap.Add "Н", "H"
 ReplacementMap.Add "Х", "X"
 ReplacementMap.Add "В", "B"
 ReplacementMap.Add "А", "A"
 ReplacementMap.Add "Р", "P"
 ReplacementMap.Add "О", "O"
 ReplacementMap.Add "С", "C"
 ReplacementMap.Add "М", "M"
 ReplacementMap.Add "Т", "T"
 ReplacementMap.Add "к", "k"
 ReplacementMap.Add "е", "e"
 ReplacementMap.Add "н", "h"
 ReplacementMap.Add "х", "x"
 ReplacementMap.Add "в", "b"
 ReplacementMap.Add "а", "a"
 ReplacementMap.Add "р", "p"
 ReplacementMap.Add "о", "o"
 ReplacementMap.Add "с", "c"
 ReplacementMap.Add "м", "m"
 ReplacementMap.Add "т", "t"
 ReplacementMap.Add "№", "N"
 
 NormalizeGOST = ""
 Dim IndexChar As Long
 Dim Char As String
 For IndexChar = 1 To Len(TempString)
 Char = Mid(TempString, IndexChar, 1)
 If ReplacementMap.Exists(Char) Then
 NormalizeGOST = NormalizeGOST & ReplacementMap(Char)
 ElseIf AscW(Char) >= 1040 And AscW(Char) <= 1103 Then ' Русские буквы
 Select Case Char
 ' Большие по ГОСТ
 Case "А": NormalizeGOST = NormalizeGOST & "A"
 Case "Б": NormalizeGOST = NormalizeGOST & "B"
 Case "В": NormalizeGOST = NormalizeGOST & "V"
 Case "Г": NormalizeGOST = NormalizeGOST & "G"
 Case "Д": NormalizeGOST = NormalizeGOST & "D"
 Case "Е", "Ё": NormalizeGOST = NormalizeGOST & "E"
 Case "Ж": NormalizeGOST = NormalizeGOST & "ZH"
 Case "З": NormalizeGOST = NormalizeGOST & "Z"
 Case "И": NormalizeGOST = NormalizeGOST & "I"
 Case "Й": NormalizeGOST = NormalizeGOST & "I"
 Case "К": NormalizeGOST = NormalizeGOST & "K"
 Case "Л": NormalizeGOST = NormalizeGOST & "L"
 Case "М": NormalizeGOST = NormalizeGOST & "M"
 Case "Н": NormalizeGOST = NormalizeGOST & "N"
 Case "О": NormalizeGOST = NormalizeGOST & "O"
 Case "П": NormalizeGOST = NormalizeGOST & "P"
 Case "Р": NormalizeGOST = NormalizeGOST & "R"
 Case "С": NormalizeGOST = NormalizeGOST & "S"
 Case "Т": NormalizeGOST = NormalizeGOST & "T"
 Case "У": NormalizeGOST = NormalizeGOST & "U"
 Case "Ф": NormalizeGOST = NormalizeGOST & "F"
 Case "Х": NormalizeGOST = NormalizeGOST & "KH"
 Case "Ц": NormalizeGOST = NormalizeGOST & "TS"
 Case "Ч": NormalizeGOST = NormalizeGOST & "CH"
 Case "Ш": NormalizeGOST = NormalizeGOST & "SH"
 Case "Щ": NormalizeGOST = NormalizeGOST & "SHCH"
 Case "Ъ": NormalizeGOST = NormalizeGOST & ""
 Case "Ы": NormalizeGOST = NormalizeGOST & "Y"
 Case "Ь": NormalizeGOST = NormalizeGOST & ""
 Case "Э": NormalizeGOST = NormalizeGOST & "E"
 Case "Ю": NormalizeGOST = NormalizeGOST & "IU"
 Case "Я": NormalizeGOST = NormalizeGOST & "IA"
 ' Маленькие по ГОСТ
 Case "а": NormalizeGOST = NormalizeGOST & "a"
 Case "б": NormalizeGOST = NormalizeGOST & "b"
 Case "в": NormalizeGOST = NormalizeGOST & "v"
 Case "г": NormalizeGOST = NormalizeGOST & "g"
 Case "д": NormalizeGOST = NormalizeGOST & "d"
 Case "е", "ё": NormalizeGOST = NormalizeGOST & "e"
 Case "ж": NormalizeGOST = NormalizeGOST & "zh"
 Case "з": NormalizeGOST = NormalizeGOST & "z"
 Case "и": NormalizeGOST = NormalizeGOST & "i"
 Case "й": NormalizeGOST = NormalizeGOST & "i"
 Case "к": NormalizeGOST = NormalizeGOST & "k"
 Case "л": NormalizeGOST = NormalizeGOST & "l"
 Case "м": NormalizeGOST = NormalizeGOST & "m"
 Case "н": NormalizeGOST = NormalizeGOST & "n"
 Case "о": NormalizeGOST = NormalizeGOST & "o"
 Case "п": NormalizeGOST = NormalizeGOST & "p"
 Case "р": NormalizeGOST = NormalizeGOST & "r"
 Case "с": NormalizeGOST = NormalizeGOST & "s"
 Case "т": NormalizeGOST = NormalizeGOST & "t"
 Case "у": NormalizeGOST = NormalizeGOST & "u"
 Case "ф": NormalizeGOST = NormalizeGOST & "f"
 Case "х": NormalizeGOST = NormalizeGOST & "kh"
 Case "ц": NormalizeGOST = NormalizeGOST & "ts"
 Case "ч": NormalizeGOST = NormalizeGOST & "ch"
 Case "ш": NormalizeGOST = NormalizeGOST & "sh"
 Case "щ": NormalizeGOST = NormalizeGOST & "shch"
 Case "ъ": NormalizeGOST = NormalizeGOST & ""
 Case "ы": NormalizeGOST = NormalizeGOST & "y"
 Case "ь": NormalizeGOST = NormalizeGOST & ""
 Case "э": NormalizeGOST = NormalizeGOST & "e"
 Case "ю": NormalizeGOST = NormalizeGOST & "iu"
 Case "я": NormalizeGOST = NormalizeGOST & "ia"
 Case Else: NormalizeGOST = NormalizeGOST & Char
 End Select
 Else
 NormalizeGOST = NormalizeGOST & Char
 End If
 Next IndexChar
 
 ' 7. Всё в CAPS
 NormalizeGOST = UCase(NormalizeGOST)
    Debug.Print "NormalizeGOST=" & NormalizeGOST
End Function


' -------------------- Parser: главная функция --------------------
' Возвращает 0 при успешном парсинге, иначе ненулевой код ошибки
Public Function Parser(ByVal sInput As String) As Long
    Dim tStart As Double: tStart = Timer
    Debug.Print LOG_PREFIX & "=== START Parser ==="
    Debug.Print LOG_PREFIX & "Вход: InputString=" & Left$(sInput, 2000)

   'On Error GoTo ErrHandler

    ' Инициализация
    InputString = sInput
    ' Используем вашу рабочую NormalizeGOST (должна быть в проекте)
    On Error Resume Next
    WorkString = NormalizeGOST(sInput)
    If Err.Number <> 0 Then
        Err.Clear
        WorkString = sInput ' fallback
    End If
   'On Error GoTo ErrHandler

    Debug.Print LOG_PREFIX & "После NormalizeGOST: " & Left$(WorkString, 2000)

    Call InitParamsArrays
    Call SetupParamNames

    ' Порядок строго по ТЗ:
    ' 6 -> 1 -> 2 -> 5 -> 4 -> 3 -> 7&8 -> 9 -> 10&11 -> 17 -> 18 -> 12 -> 14 -> 13 -> 15 -> 16 -> 19 -> 20

    ' ----- 6. P_HCX (#6) -----
    Debug.Print LOG_PREFIX & "Шаг #6: Поиск P_HCX..."
    If Not Find_HCX_FromSheet(6) Then
        ParamErrorCode(6) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Не определен НСХ (P_HCX).") Then Parser = 106: Exit Function
    Else
        Debug.Print LOG_PREFIX & "#6 найдено: " & ParamValue(6)
        ReplaceParam 6
    End If

    ' ----- 1. P_TYPE (#1) -----
    Debug.Print LOG_PREFIX & "Шаг #1: P_TYPE..."
    If Not RegexTest("^TC[\-\s]*1388", WorkString, False) Then
        ' добавить дефолт в начало
        ParamValue(1) = "TC-1388"
        ParamStart(1) = 1
        ParamEnd(1) = Len(ParamValue(1))
        ParamErrorCode(1) = ERR_FOUND_NOT_CLEANED
        WorkString = ParamValue(1) & "/" & WorkString
        Debug.Print LOG_PREFIX & "Добавлен дефолт P_TYPE в начало: " & ParamValue(1)
    End If
    If FindRegexInWork("^TC[\-\s]*1388", 1) Then
        ParamValue(1) = "TC-1388"
        Debug.Print LOG_PREFIX & "#1 найден: " & ParamValue(1)
        ReplaceParam 1
    Else
        ParamErrorCode(1) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Не найден шаблон P_TYPE (TC-1388).") Then Parser = 101: Exit Function
    End If

    ' ----- 2. P_ISPOLN (#2) -----
    Debug.Print LOG_PREFIX & "Шаг #2: P_ISPOLN..."
    If Find_ISPOLN(2) Then
        Debug.Print LOG_PREFIX & "#2 найден: " & ParamValue(2)
        ReplaceParam 2
    Else
        ParamErrorCode(2) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Не определен вид исполнения (P_ISPOLN).") Then Parser = 102: Exit Function
    End If

    ' ----- 5. P_EX (#5) -----
    Debug.Print LOG_PREFIX & "Шаг #5: P_EX..."
    If InStr(1, ParamValue(2), "EX", vbTextCompare) > 0 Or InStr(1, ParamValue(2), "EXD", vbTextCompare) > 0 Then
        If Find_EX_Label(5, 6, 2) Then
            Debug.Print LOG_PREFIX & "#5 найден(Ex): " & ParamValue(5)
            ReplaceParam 5
        Else
            ParamErrorCode(5) = ERR_NOT_DEFINED
            If Not Handle_StopOrContinue("Не указана маркировка взрывозащиты для исп. Ex (P_EX).") Then Parser = 105: Exit Function
        End If
    Else
        If FindLiteralDashInRange(5, ParamStart(6), ParamEnd(1)) Then
            Debug.Print LOG_PREFIX & "#5 найден (dash): " & ParamValue(5)
            ReplaceParam 5
        Else
            ParamErrorCode(5) = ERR_NOT_DEFINED
            If Not Handle_StopOrContinue("Параметры взрывозащиты указаны для не Ex исполнения (P_EX).") Then Parser = 150: Exit Function
        End If
    End If

    ' ----- 4. P_KL_AES (#4) -----
    Debug.Print LOG_PREFIX & "Шаг #4: P_KL_AES..."
    If InStr(1, ParamValue(2), "A", vbTextCompare) > 0 Then
        If FindRegexInRange_Longest("(2|3|4)\s*(НУ|HU|HУ|H|Н)?", 4, ParamEnd(2), ParamStart(5)) Then
            Debug.Print LOG_PREFIX & "#4 найден (A): " & ParamValue(4)
            ReplaceParam 4
        Else
            ParamErrorCode(4) = ERR_NOT_DEFINED
            If Not Handle_StopOrContinue("Не найден класс безопасности для исполнения А (P_KL_AES).") Then Parser = 104: Exit Function
        End If
    Else
        If FindLiteralDashInRange(4, ParamStart(5), ParamEnd(1)) Then
            Debug.Print LOG_PREFIX & "#4 найден (dash): " & ParamValue(4)
            ReplaceParam 4
        Else
            ParamErrorCode(4) = ERR_NOT_DEFINED
            If Not Handle_StopOrContinue("Класс безопасности не определен (P_KL_AES).") Then Parser = 140: Exit Function
        End If
    End If

    ' ----- 3. P_MODEL (#3) -----
    Debug.Print LOG_PREFIX & "Шаг #3: P_MODEL..."
    If FindModelBetween(3, ParamEnd(2), ParamStart(4)) Then
        Debug.Print LOG_PREFIX & "#3 найден: " & ParamValue(3)
        ReplaceParam 3
    Else
        ParamErrorCode(3) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Не определен конструктив (P_MODEL).") Then Parser = 103: Exit Function
    End If

    ' ----- 7 & 8. P_T_LOW (#7) и P_T_HIGH (#8) -----
    Debug.Print LOG_PREFIX & "Шаги #7/#8: Темп. диапазон..."
    If FindTemperatureRange(7, 8, ParamEnd(6) + 1, NextSeparatorPos(ParamEnd(6) + 2) - 1) Then
        ' По ТЗ: сначала ReplaceParam(8), затем ReplaceParam(7)
        Debug.Print LOG_PREFIX & "Найден температурный диапазон: " & ParamValue(7) & " .. " & ParamValue(8)
        ReplaceParam 8
        WorkString = Left(WorkString, ParamStart(8) - 2) & "/" & Mid(WorkString, ParamStart(8))
        ParamEnd(7) = ParamStart(8) - 2
        ReplaceParam 7
    Else
        ParamErrorCode(7) = ERR_NOT_DEFINED
        ParamErrorCode(8) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Не найден температурный диапазон (P_T_LOW / P_T_HIGH).") Then Parser = 107: Exit Function
    End If
    
    ' ----- 17. P_CXEMA (#17) -----
    Debug.Print LOG_PREFIX & "Шаг #17: P_CXEMA..."
    If FindRegexInRange_Longest("(?:№|N)([1-6])", 17, ParamEnd(8) + 1, Len(WorkString)) Then
        Debug.Print LOG_PREFIX & "#17 найден: " & ParamValue(17)
        ReplaceParam 17
    Else
        ParamErrorCode(17) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Не найдена схема (P_CXEMA).") Then Parser = 117: Exit Function
    End If

    ' ----- 9. P_DLINA (#9) -----
    Debug.Print LOG_PREFIX & "Шаг #9: P_DLINA..."
    If FindDLINA(9, ParamEnd(8) + 1, ParamStart(17) - 1) Then
        Debug.Print LOG_PREFIX & "#9 найден: " & ParamValue(9)
        ReplaceParam 9
    Else
        ParamErrorCode(9) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Длина не определена (P_DLINA).") Then Parser = 109: Exit Function
    End If

    ' ----- 10 & 11. P_DIAMETR (#10) и P_SHTUCER (#11) -----
    Debug.Print LOG_PREFIX & "Шаги #10/#11: Диаметр / Штуцер..."
    If InStr(1, ParamValue(9), "X", vbTextCompare) > 0 Then
        ParamValue(10) = "-"
        ParamStart(10) = ParamEnd(9) + 1
        ParamEnd(10) = ParamStart(10)
        ParamErrorCode(10) = ERR_OK
        ParamValue(11) = "-"
        ParamStart(11) = ParamEnd(10) + 1
        ParamEnd(11) = ParamStart(11)
        ParamErrorCode(11) = ERR_OK
        Debug.Print LOG_PREFIX & "#10/#11: поставлены '-' (P_DLINA содержит X)"
    Else
        If FindDiameterAndShtucer(10, 11, ParamEnd(8) + 1, ParamStart(17) - 1) Then
            Debug.Print LOG_PREFIX & "#10 найден: " & ParamValue(10)
            ReplaceParam 10
            If ParamValue(11) <> "" Then ReplaceParam 11
        Else
            ParamErrorCode(10) = ERR_NOT_DEFINED
            If Not Handle_StopOrContinue("Диаметр или штуцер не определены (P_DIAMETR / P_SHTUCER).") Then Parser = 110: Exit Function
        End If
    End If

    ' ----- 18. P_GP (#18) -----
    Debug.Print LOG_PREFIX & "Шаг #18: P_GP..."
    If FindRegexInRange_Longest("(?:-|GP)", 18, ParamEnd(17) + 1, Len(WorkString)) Then
        Debug.Print LOG_PREFIX & "#18 найден: " & ParamValue(18)
        ReplaceParam 18
    Else
        ParamErrorCode(18) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Не найдено обозначение GP (P_GP).") Then Parser = 118: Exit Function
    End If

    ' ----- 12. P_L_KAB (#12) -----
    Debug.Print LOG_PREFIX & "Шаг #12: P_L_KAB..."
    If FindRegexInRange_Longest("^\-?\d+(?:,\d+)?$|^-$", 12, ParamEnd(11) + 1, ParamStart(17) - 1) Then
        Debug.Print LOG_PREFIX & "#12 найден: " & ParamValue(12)
        ReplaceParam 12
    Else
        ParamErrorCode(12) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Не найден L_KAB (P_L_KAB).") Then Parser = 112: Exit Function
    End If

    ' ----- 14. P_KLASS (#14) -----
    Debug.Print LOG_PREFIX & "Шаг #14: P_KLASS..."
    If FindRegexInRange_Longest("(AA|A|B|C)", 14, ParamEnd(12) + 1, ParamStart(17) - 1) Then
        Debug.Print LOG_PREFIX & "#14 найден: " & ParamValue(14)
        ReplaceParam 14
    Else
        ParamErrorCode(14) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Не найден класс (P_KLASS).") Then Parser = 114: Exit Function
    End If

    ' ----- 13. P_KABEL (#13) -----
    Debug.Print LOG_PREFIX & "Шаг #13: P_KABEL..."
    If FindBetweenAndReplace(13, ParamEnd(12) + 1, ParamStart(14) - 1) Then
        Debug.Print LOG_PREFIX & "#13 найден: " & ParamValue(13)
        ReplaceParam 13
    Else
        ParamErrorCode(13) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Не найден кабель (P_KABEL).") Then Parser = 113: Exit Function
    End If

    ' ----- 15. P_HEAD (#15) -----
    Debug.Print LOG_PREFIX & "Шаг #15: P_HEAD..."
    If FindRegexInRange_Longest("(\s*\/?-\s*\/?\s*)", 15, ParamEnd(14) + 1, ParamStart(17) - 1) Then
        ' заменяем на '-'
        ParamValue(15) = "-"
        ReplaceParam 15
        Debug.Print LOG_PREFIX & "#15 заменён на '-'"
    Else
        ParamErrorCode(15) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Не найден P_HEAD.") Then Parser = 115: Exit Function
    End If

    ' ----- 16. P_PLUG (#16) -----
    Debug.Print LOG_PREFIX & "Шаг #16: P_PLUG..."
    If FindRegexInRange_Longest("([A-ZА-Я0-9-]+|^-)", 16, ParamEnd(15) + 1, ParamStart(17) - 1) Then
        Debug.Print LOG_PREFIX & "#16 найден: " & ParamValue(16)
        ReplaceParam 16
    Else
        ParamErrorCode(16) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Не найден P_PLUG.") Then Parser = 116: Exit Function
    End If

    ' ----- 19. P_TU (#19) -----
    Debug.Print LOG_PREFIX & "Шаг #19: P_TU..."
    If FindRegexInRange_Longest("TU\s*\d{4}-\d{3}-\d{8}-\d{4}|^-", 19, ParamEnd(18) + 1, Len(WorkString)) Then
        If Trim$(ParamValue(19)) = "-" Then ParamValue(19) = "-"
        ReplaceParam 19
        Debug.Print LOG_PREFIX & "#19 найден: " & ParamValue(19)
    Else
        ParamErrorCode(19) = ERR_NOT_DEFINED
        If Not Handle_StopOrContinue("Не найден P_TU.") Then Parser = 119: Exit Function
    End If

    ' ----- 20. P_H3 (#20) -----
    Debug.Print LOG_PREFIX & "Шаг #20: P_H3..."
    If ParamEnd(17) > 0 And ParamEnd(17) < Len(WorkString) Then
        Dim tailStart As Long
        tailStart = ParamEnd(17) + 1
        ParamStart(20) = tailStart
        ParamEnd(20) = Len(WorkString)
        ParamValue(20) = Trim$(Mid$(WorkString, tailStart))
        ParamErrorCode(20) = ERR_FOUND_NOT_CLEANED
        ReplaceParam 20
        Debug.Print LOG_PREFIX & "#20 найден: " & ParamValue(20)
    Else
        ParamValue(20) = ""
        ParamStart(20) = 0
        ParamEnd(20) = 0
        ParamErrorCode(20) = ERR_NOT_DEFINED
    End If

    ' Финал: вывести краткий лог и вернуть 0
    Debug.Print LOG_PREFIX & "=== Parser completed OK in " & Format$(Timer - tStart, "0.00") & "s ==="
    Dim i As Integer
    For i = 1 To 20
        Debug.Print "  #" & i & " " & ParamName(i) & " = '" & ParamValue(i) & "' (" & ParamErrorCode(i) & ")"
    Next i

    Parser = 0
    Exit Function

ErrHandler:
    Debug.Print LOG_PREFIX & "CRITICAL ERROR: " & Err.Number & " - " & Err.Description
    Parser = 999
End Function

' -------------------- Init и Setup --------------------
Private Sub InitParamsArrays()
    Dim i As Integer
    For i = 1 To 20
        ParamName(i) = ""
        ParamValue(i) = ""
        ParamStart(i) = 0
        ParamEnd(i) = 0
        ParamErrorCode(i) = ERR_NOT_DEFINED
    Next i
End Sub

Private Sub SetupParamNames()
    ParamName(1) = "P_TYPE"
    ParamName(2) = "P_ISPOLN"
    ParamName(3) = "P_MODEL"
    ParamName(4) = "P_KL_AES"
    ParamName(5) = "P_EX"
    ParamName(6) = "P_HCX"
    ParamName(7) = "P_T_LOW"
    ParamName(8) = "P_T_HIGH"
    ParamName(9) = "P_DLINA"
    ParamName(10) = "P_DIAMETR"
    ParamName(11) = "P_SHTUCER"
    ParamName(12) = "P_L_KAB"
    ParamName(13) = "P_KABEL"
    ParamName(14) = "P_KLASS"
    ParamName(15) = "P_HEAD"
    ParamName(16) = "P_PLUG"
    ParamName(17) = "P_CXEMA"
    ParamName(18) = "P_GP"
    ParamName(19) = "P_TU"
    ParamName(20) = "P_H3"
End Sub

' -------------------- ReplaceParam --------------------
Public Sub ReplaceParam(ByVal n As Integer)
    Dim leftPart As String, rightPart As String
    Dim newToken As String
    Dim cleaned As String
    Dim i As Integer
    Dim pos As Long, slashPos As Long

    Debug.Print LOG_PREFIX & "ReplaceParam: start #" & n & " (" & ParamName(n) & ")"
    Debug.Print LOG_PREFIX & "  вход: Start=" & ParamStart(n) & " End=" & ParamEnd(n) & " Value='" & ParamValue(n) & "' Code=" & ParamErrorCode(n)

    If n < 1 Or n > 20 Then Exit Sub
    If ParamValue(n) = "" Or ParamErrorCode(n) <> ERR_FOUND_NOT_CLEANED Then
        Debug.Print LOG_PREFIX & "  ReplaceParam: условие замены не выполнено."
        Exit Sub
    End If

    ' очистка значения по ТЗ (удаляем '/', пробелы, скобки)
    cleaned = ParamValue(n)
    cleaned = Replace(cleaned, "/", "")
    cleaned = Replace(cleaned, " ", "")
    cleaned = Replace(cleaned, "(", "")
    cleaned = Replace(cleaned, ")", "")
    cleaned = Trim$(cleaned)

    newToken = "#" & CStr(n) & "_" & cleaned & "%"

    ' делим WorkString на части
    If ParamStart(n) > 1 Then
        leftPart = Left$(WorkString, ParamStart(n) - 1)
    Else
        leftPart = ""
    End If

    If ParamEnd(n) < Len(WorkString) Then
        rightPart = Mid$(WorkString, ParamEnd(n) + 1)
    Else
        rightPart = ""
    End If

    ' вставка нового токена
    WorkString = leftPart & newToken & rightPart

    ' пересчёт всех ParamStart/ParamEnd через поиск токена #i_
    For i = 1 To 20
        pos = InStr(1, WorkString, "#" & i & "_", vbTextCompare)
        If pos > 0 Then
            ParamStart(i) = pos
            ' ищем конец до ближайшего '/' после #
            slashPos = InStr(pos, WorkString, "/")
            If slashPos = 0 Then slashPos = Len(WorkString) + 1
            ParamEnd(i) = slashPos - 1
        Else
            ParamStart(i) = 0
            ParamEnd(i) = 0
        End If
    Next i

    ' обновляем значение и код
    ParamValue(n) = cleaned
    ParamErrorCode(n) = ERR_OK

    Call GetDebugString
    Debug.Print LOG_PREFIX & "ReplaceParam: done #" & n & " NewToken=" & newToken
    Debug.Print LOG_PREFIX & "  WorkString now: " & Left$(WorkString, 2000)
    Debug.Print "> After replace #" & n & " : Start=" & ParamStart(n) & " points to '" & Mid$(WorkString, ParamStart(n), 1) & "', End=" & ParamEnd(n) & " points to '" & Mid$(WorkString, ParamEnd(n), 1) & "'"
End Sub

' -------------------- Поисковые и вспомогательные функции --------------------

' Поиск NСХ из листа "1" диапазон M12:M19
Private Function Find_HCX_FromSheet(ByVal idx As Integer) As Boolean
   'On Error GoTo EH
    Dim sh As Worksheet: Set sh = ThisWorkbook.Worksheets("1")
    Dim rng As Range: Set rng = sh.Range("M12:M19")
    Dim cell As Range
    Dim posMin As Long: posMin = 0
    Dim selVal As String
    For Each cell In rng
        Dim cv As String
        cv = Trim$(CStr(cell.Value))
        If cv <> "" Then
            Dim pos As Long
            pos = InStr(1, WorkString, cv, vbTextCompare)
            If pos = 0 Then
                pos = RegExpFirstPos(EscapeForRegex(cv), WorkString)
            End If
            If pos > 0 Then
                If posMin = 0 Or pos < posMin Then
                    posMin = pos
                    selVal = cv
                End If
            End If
        End If
    Next cell
    If posMin > 0 Then
        ParamValue(idx) = selVal
        ParamStart(idx) = posMin
        ParamEnd(idx) = posMin + Len(selVal) - 1
        ParamErrorCode(idx) = ERR_FOUND_NOT_CLEANED
        Find_HCX_FromSheet = True
    Else
        Find_HCX_FromSheet = False
    End If
    Exit Function
EH:
    Debug.Print LOG_PREFIX & "Find_HCX_FromSheet ERROR: " & Err.Description
    Find_HCX_FromSheet = False
End Function

' Поиск по RegExp по всему WorkString
Private Function FindRegexInWork(ByVal pattern As String, ByVal idx As Integer) As Boolean
    FindRegexInWork = FindRegexInRange_Longest(pattern, idx, 1, Len(WorkString))
End Function

' Находит самое длинное совпадение в диапазоне (fromPos..toPos)
Private Function FindRegexInRange_Longest(ByVal pattern As String, ByVal idx As Integer, ByVal fromPos As Long, ByVal toPos As Long) As Boolean
    'On Error GoTo EH
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.pattern = pattern
    re.Global = True
    re.IgnoreCase = True
    If fromPos < 1 Then fromPos = 1
    If toPos > Len(WorkString) Then toPos = Len(WorkString)
    If fromPos > toPos Then FindRegexInRange_Longest = False: Exit Function
    Dim subStr As String: subStr = Mid$(WorkString, fromPos, toPos - fromPos + 1)
    Dim matches As Object
    If re.Test(subStr) Then
        Set matches = re.Execute(subStr)
        Dim bestLen As Long: bestLen = 0
        Dim bestMatch As Object
        Dim m As Object
        For Each m In matches
            If Len(m.Value) >= bestLen Then
                bestLen = Len(m.Value)
                Set bestMatch = m
            End If
        Next m
        If Not bestMatch Is Nothing Then
            ParamValue(idx) = Trim$(bestMatch.Value)
            ParamStart(idx) = fromPos + bestMatch.FirstIndex
            ParamEnd(idx) = ParamStart(idx) + Len(bestMatch.Value) - 1
            ParamErrorCode(idx) = ERR_FOUND_NOT_CLEANED
            FindRegexInRange_Longest = True
            Exit Function
        End If
    End If
    FindRegexInRange_Longest = False
    Exit Function
EH:
    Debug.Print LOG_PREFIX & "FindRegexInRange_Longest ERROR: " & Err.Description
    FindRegexInRange_Longest = False
End Function

' Простой вытяг между позициями
Private Function FindBetweenAndReplace(ByVal idx As Integer, ByVal fromPos As Long, ByVal toPos As Long) As Boolean
    If fromPos < 1 Then fromPos = 1
    If toPos > Len(WorkString) Then toPos = Len(WorkString)
    If fromPos > toPos Then FindBetweenAndReplace = False: Exit Function
    Dim s As String: s = Trim$(Mid$(WorkString, fromPos, toPos - fromPos + 1))
    If s <> "" Then
        ParamValue(idx) = s
        ParamStart(idx) = fromPos
        ParamEnd(idx) = toPos
        ParamErrorCode(idx) = ERR_FOUND_NOT_CLEANED
        FindBetweenAndReplace = True
    Else
        FindBetweenAndReplace = False
    End If
End Function

' Поиск P_ISPOLN по ТЗ (между ParamEnd(1) и разделителем "/" или ParamStart(6))
Private Function Find_ISPOLN(ByVal idx As Integer) As Boolean
    'On Error GoTo EH
    Dim fromPos As Long: fromPos = ParamEnd(1) + 1
    If fromPos < 1 Then fromPos = 1
    Dim toPos As Long
    If ParamStart(6) > 0 Then
        toPos = ParamStart(6) - 1
    Else
        toPos = NextSeparatorPos(fromPos) - 1
    End If
    If toPos < fromPos Then toPos = Len(WorkString)
    Dim pattern As String
    pattern = "((A\s*)?(EXD?|EX)\s*(B|BC)?\s*(V3|N3|F3|G2))|(B\s*V3)|(N3)|(F3)|(G2)|-"
    Find_ISPOLN = FindRegexInRange_Longest(pattern, idx, fromPos, toPos)
    Exit Function
EH:
    Debug.Print LOG_PREFIX & "Find_ISPOLN ERROR: " & Err.Description
    Find_ISPOLN = False
End Function

Private Function Find_EX_Label(ByVal idx As Integer, ByVal hcxIdx As Integer, ByVal ispolnIdx As Integer) As Boolean
    'On Error GoTo EH
    Dim leftPos As Long, rightPos As Long
    leftPos = ParamEnd(ispolnIdx) + 1
    If leftPos < 1 Then leftPos = 1
    rightPos = ParamStart(hcxIdx) - 1
    If rightPos < leftPos Then rightPos = Len(WorkString) ' или leftPos, чтобы диапазон был корректным

    Dim pattern As String
    ' Паттерн упрощён для работы без пробелов, но учитывает основные варианты EX/EXD
pattern = "(0|1)\s*([EeЕе][XxХх]\s*IA|EXD\s*DB)\s*II\s*([ABCАВС])\s*T[1-6]\s*G[AaBb]\s*X"

    ' Ищем в диапазоне слева-направо
    Find_EX_Label = FindRegexInRange_Longest(pattern, idx, leftPos, rightPos)
    Exit Function
EH:
    Debug.Print LOG_PREFIX & "Find_EX_Label ERROR: " & Err.Description
    Find_EX_Label = False
End Function


' Ищем литерал '-' в диапазоне и ставим параметр
Private Function FindLiteralDashInRange(ByVal idx As Integer, ByVal fromPos As Long, ByVal toPos As Long) As Boolean
    If fromPos < 1 Then fromPos = 1
    If toPos > Len(WorkString) Then toPos = Len(WorkString)
    If fromPos > toPos Then FindLiteralDashInRange = False: Exit Function
    Dim seg As String: seg = Mid$(WorkString, fromPos, toPos - fromPos + 1)
    Dim p As Long: p = InStr(1, seg, "-", vbTextCompare)
    If p > 0 Then
        ParamValue(idx) = "-"
        ParamStart(idx) = fromPos + p - 1
        ParamEnd(idx) = ParamStart(idx)
        ParamErrorCode(idx) = ERR_FOUND_NOT_CLEANED
        FindLiteralDashInRange = True
    Else
        FindLiteralDashInRange = False
    End If
End Function

' Поиск модели между двумя позициями
Private Function FindModelBetween(ByVal idx As Integer, ByVal posA As Long, ByVal posB As Long) As Boolean
    If posA < 1 Then posA = 1
    If posB > Len(WorkString) Then posB = Len(WorkString)
    If posA > posB Then FindModelBetween = False: Exit Function
    Dim frag As String: frag = Mid$(WorkString, posA + 1, posB - posA - 1)
    Dim pattern As String: pattern = "^/?(\d+)(?:-(\d+))?/?(.+)?/?$"
    If RegexTest(pattern, frag, True) Then
        Dim re As Object: Set re = CreateObject("VBScript.RegExp")
        re.pattern = pattern
        re.IgnoreCase = True
        If re.Test(frag) Then
            Dim m As Object: Set m = re.Execute(frag)(0)
            ParamValue(idx) = Trim$(m.Value)
            ParamStart(idx) = posA + m.FirstIndex
            ParamEnd(idx) = ParamStart(idx) + Len(m.Value) - 1
            ParamErrorCode(idx) = ERR_FOUND_NOT_CLEANED
            FindModelBetween = True
            Exit Function
        End If
    End If
    FindModelBetween = False
End Function

' Поиск температурного диапазона (два числа)
Private Function FindTemperatureRange(ByVal idxLow As Integer, ByVal idxHigh As Integer, ByVal fromPos As Long, ByVal toPos As Long) As Boolean
    If fromPos < 1 Then fromPos = 1
    If toPos > Len(WorkString) Then toPos = Len(WorkString)
    If fromPos > toPos Then FindTemperatureRange = False: Exit Function
    Dim frag As String: frag = Mid$(WorkString, fromPos, toPos - fromPos + 1)
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.pattern = "([+\-]?\d+(?:,\d+)?)\D+([+\-]?\d+(?:,\d+)?)"
    re.Global = False
    re.IgnoreCase = True
    If re.Test(frag) Then
        Dim m As Object: Set m = re.Execute(frag)(0)
        Dim lowVal As String: lowVal = m.SubMatches(0)
        Dim highVal As String: highVal = m.SubMatches(1)
        Dim loc As Long
        loc = InStr(fromPos, WorkString, m.Value, vbTextCompare)
        If loc = 0 Then loc = fromPos
        ParamValue(idxLow) = Trim$(lowVal)
        ParamStart(idxLow) = loc
        ParamEnd(idxLow) = ParamStart(idxLow) + Len(lowVal) - 1
        ParamErrorCode(idxLow) = ERR_FOUND_NOT_CLEANED
        Dim offsetHigh As Long
        offsetHigh = InStr(1, m.Value, m.SubMatches(1), vbTextCompare)
        If offsetHigh > 0 Then
            ParamValue(idxHigh) = Trim$(m.SubMatches(1))
            ParamStart(idxHigh) = ParamStart(idxLow) + offsetHigh - 1
            ParamEnd(idxHigh) = ParamStart(idxHigh) + Len(ParamValue(idxHigh)) - 1
            ParamErrorCode(idxHigh) = ERR_FOUND_NOT_CLEANED
            FindTemperatureRange = True
            Exit Function
        End If
    End If
    FindTemperatureRange = False
End Function

' Поиск длины (DLINA)
Private Function FindDLINA(ByVal idx As Integer, ByVal fromPos As Long, ByVal toPos As Long) As Boolean
    If fromPos < 1 Then fromPos = 1
    If toPos > Len(WorkString) Then toPos = Len(WorkString)
    If fromPos > toPos Then FindDLINA = False: Exit Function
    Dim frag As String: frag = Mid$(WorkString, fromPos, toPos - fromPos + 1)
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.pattern = "\d+(?:,\d+)?(?:\s*[Xxх]\s*\d+(?:,\d+)?(?:\s*[Xxх]\s*\d+(?:,\d+)?)?)?"
    re.Global = False
    re.IgnoreCase = True
    If re.Test(frag) Then
        Dim m As Object: Set m = re.Execute(frag)(0)
        ParamValue(idx) = Trim$(m.Value)
        ParamStart(idx) = fromPos + m.FirstIndex
        ParamEnd(idx) = ParamStart(idx) + Len(m.Value) - 1
        ParamErrorCode(idx) = ERR_FOUND_NOT_CLEANED
        FindDLINA = True
    Else
        FindDLINA = False
    End If
End Function

' Поиск диаметра и штуцера
Private Function FindDiameterAndShtucer(ByVal idxDiam As Integer, ByVal idxSht As Integer, ByVal fromPos As Long, ByVal toPos As Long) As Boolean
    If fromPos < 1 Then fromPos = 1
    If toPos > Len(WorkString) Then toPos = Len(WorkString)
    If fromPos > toPos Then FindDiameterAndShtucer = False: Exit Function
    Dim frag As String: frag = Mid$(WorkString, fromPos, toPos - fromPos + 1)
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.pattern = "(\d+(?:,\d+)?)(?:\s*\(\s*([A-Za-z0-9\-хХ]+)\s*\))?"
    re.Global = False
    re.IgnoreCase = True
    If re.Test(frag) Then
        Dim m As Object: Set m = re.Execute(frag)(0)
        ParamValue(idxDiam) = Trim$(m.SubMatches(0))
        ParamStart(idxDiam) = fromPos + m.FirstIndex
        ParamEnd(idxDiam) = ParamStart(idxDiam) + Len(ParamValue(idxDiam)) - 1
        ParamErrorCode(idxDiam) = ERR_FOUND_NOT_CLEANED
        If m.SubMatches.Count >= 2 Then
            If Trim$(m.SubMatches(1)) <> "" Then
                ParamValue(idxSht) = Trim$(m.SubMatches(1))
                ParamStart(idxSht) = ParamEnd(idxDiam) + 1
                ParamEnd(idxSht) = ParamStart(idxSht) + Len(ParamValue(idxSht)) - 1
                ParamErrorCode(idxSht) = ERR_FOUND_NOT_CLEANED
            End If
        End If
        FindDiameterAndShtucer = True
    Else
        FindDiameterAndShtucer = False
    End If
End Function

' -------------------- Утилиты RegExp и позиции --------------------
Private Function RegexTest(ByVal pattern As String, ByVal s As String, ByVal fullMatch As Boolean) As Boolean
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.pattern = pattern
    re.IgnoreCase = True
    re.Global = False
    RegexTest = re.Test(s)
End Function

Private Function RegExpFirstPos(ByVal pattern As String, ByVal s As String) As Long
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.pattern = pattern
    re.IgnoreCase = True
    re.Global = False
    If re.Test(s) Then
        Dim m As Object: Set m = re.Execute(s)(0)
        RegExpFirstPos = m.FirstIndex + 1
    Else
        RegExpFirstPos = 0
    End If
End Function

Private Function NextSeparatorPos(ByVal pos As Long) As Long
    Dim p As Long
    p = InStr(pos, WorkString, "/", vbTextCompare)
    If p = 0 Then NextSeparatorPos = Len(WorkString) + 1 Else NextSeparatorPos = p
End Function

Private Function EscapeForRegex(ByVal s As String) As String
    Dim chars As Variant
    chars = Array("\", "^", "$", ".", "|", "?", "*", "+", "(", ")", "[", "{", "]", "}")
    Dim i As Integer
    For i = LBound(chars) To UBound(chars)
        s = Replace$(s, chars(i), "\" & chars(i))
    Next i
    EscapeForRegex = s
End Function

' -------------------- Обработка ошибок: MsgBox Stop/Continue --------------------
' Возвращает True если пользователь выбрал "Продолжить", False если "Прекратить"
Private Function Handle_StopOrContinue(ByVal errText As String) As Boolean
    Dim msg As String
    msg = "Ошибка парсинга: " & errText & vbCrLf & vbCrLf & "Выберите действие:" & vbCrLf & "Да - Прекратить парсинг и выйти" & vbCrLf & "Нет - Продолжить парсинг (ошибка будет зафиксирована)"
    Dim res As VbMsgBoxResult
    res = MsgBox(msg, vbYesNo + vbExclamation, "Ошибка парсинга")
    If res = vbYes Then
        Handle_StopOrContinue = False
    Else
        Handle_StopOrContinue = True
    End If
End Function
' -------------------------
' Генерация строки отладки с проверкой соответствия WorkString
Public Function GetDebugString()
    Dim Index As Integer
    Dim dbg As String
    dbg = "ParamStart(i):"

    For Index = 1 To 20
        dbg = dbg & ParamStart(Index) & ","
    Next Index
    dbg = Left(dbg, Len(dbg) - 1) ' убираем последнюю запятую

    dbg = dbg & " ParamEnd(i):"
    For Index = 1 To 20
        dbg = dbg & ParamEnd(Index) & ","
    Next Index
    dbg = Left(dbg, Len(dbg) - 1)

    Debug.Print "DebugString=" & dbg

    ' Дополнительно проверяем фактические позиции в WorkString
    For Index = 1 To 20
        If ParamStart(Index) > 0 And ParamEnd(Index) >= ParamStart(Index) Then
            Dim expectedPos As Long
            Dim foundPos As Long
            Dim tokenPattern As String
            tokenPattern = "#" & Index & "_"

            ' Ищем позицию токена #i_ в WorkString
            foundPos = InStr(1, WorkString, tokenPattern, vbTextCompare)

            If foundPos = 0 Then
                Debug.Print "Warning: Token " & tokenPattern & " not found in WorkString!"
            Else
                ' Сравниваем с ParamStart
                If ParamStart(Index) <> foundPos Then
                    Debug.Print "Warning: ParamStart(" & Index & ")=" & ParamStart(Index) & _
                                " does NOT match actual position=" & foundPos
                End If
                ' Сравниваем с ParamEnd
                Dim actualEnd As Long
                ' Ищем конец токена до ближайшего '/' или конца строки
                Dim slashPos As Long
                slashPos = InStr(foundPos, WorkString, "/")
                If slashPos = 0 Then slashPos = Len(WorkString) + 1
                actualEnd = slashPos - 1
                If ParamEnd(Index) <> actualEnd Then
                    Debug.Print "Warning: ParamEnd(" & Index & ")=" & ParamEnd(Index) & _
                                " does NOT match actual end=" & actualEnd
                End If
            End If
        End If
    Next Index
End Function

