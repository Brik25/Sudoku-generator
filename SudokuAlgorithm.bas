Attribute VB_Name = "SudokuAlgorithm"
'Autor: Gainullin Ramil
Option Explicit

Sub AutoSudoku()
    Dim mainSheet As Worksheet
    Dim i As Integer, j As Integer, Count As Integer, countALL As Integer, NUM As Integer, countBlock As Integer, FullValue As Integer
    Dim RandDict As Integer, RandValue As Integer, RandKey As Integer, TotalLook As Integer, Camultive As Integer
    Dim RangeDictionary As Object, StartEnd As Object, AddressRange As Object, RandNum As Object, AllAddressTable As Object
    Dim existRow As Boolean, existsColumn As Boolean, dictC As Boolean
    Dim RangeName As Variant, RangeCell As Variant, Key As String
    Dim startTable As Range, EndTable As Range
    
    Set mainSheet = ThisWorkbook.Sheets("Sudoku")
    
    With mainSheet
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        
        Set startTable = .Range("StartTable")
        Set EndTable = .Range("EndTable")
INSL:
        Set RangeDictionary = CreateObject("Scripting.Dictionary")
        
        .Range(startTable.Address(False, False) & ":" & EndTable.Address(False, False)).Value = ""
        .Range(startTable.Address(False, False) & ":" & EndTable.Address(False, False)).Font.Color = vbBlack
        
        countALL = 0
        Count = 0
        'Create block's in Table
        For i = startTable.Column To EndTable.Column Step 3
            For j = startTable.Row To EndTable.Row Step 3
                Set StartEnd = CreateObject("Scripting.Dictionary")
                StartEnd.Add "Start", .Cells(j, i).Address(False, False)
                StartEnd.Add "End", .Cells(j + 2, i + 2).Address(False, False)
                RangeDictionary.Add "Range" & Count, StartEnd
                Count = Count + 1
            Next
        Next
        
        'Variation Value
        For Each RangeName In RangeDictionary
INSB:
            If countALL = 10 Then
                GoTo INSL
            End If

            Set AddressRange = CreateObject("Scripting.Dictionary")
            Set RandNum = CreateObject("Scripting.Dictionary")
            For Each RangeCell In .Range(RangeDictionary(RangeName)("Start") & ":" & RangeDictionary(RangeName)("End"))
                AddressRange.Add RangeCell.Address(False, False), ""
            Next RangeCell
            NUM = 8
            
            For i = 1 To 9
                RandNum.Add i, ""
            Next

            countBlock = 0
            countALL = countALL + 1
            
            While AddressRange.Count <> 0
                RandDict = WorksheetFunction.RandBetween(0, NUM)
                RandValue = WorksheetFunction.RandBetween(0, NUM)
    
                existRow = False
                existsColumn = False
                
                Key = AddressRange.keys()(RandDict)
                RandKey = RandNum.keys()(RandValue)
 
                For i = startTable.Row To EndTable.Row
                    If .Cells(i, .Range(Key).Column).Value = RandKey Then
                        existRow = True
                    End If
                Next i
    
                For i = startTable.Column To EndTable.Column
                    If .Cells(.Range(Key).Row, i).Value = RandKey Then
                        existsColumn = True
                    End If
                Next i

                If existRow = False And existsColumn = False Then
                    .Range(Key).Font.Color = vbRed
                    .Range(Key).Value = RandKey
                    RandNum.Remove RandKey
                    AddressRange.Remove Key
                    NUM = NUM - 1
                Else
                    countBlock = countBlock + 1
                End If
                
                If countBlock = 100 Then
                    .Range(RangeDictionary(RangeName)("Start") & ":" & RangeDictionary(RangeName)("End")).Value = ""
                    GoTo INSB
                End If
            Wend
        Next RangeName
        
        'Clear random Cell in Full result
        FullValue = WorksheetFunction.CountA(.Range(startTable.Address(False, False) & ":" & EndTable.Address(False, False)))
        If FullValue = 81 Then
            Set AllAddressTable = CreateObject("Scripting.Dictionary")
    
            For Each RangeCell In .Range(startTable.Address & ":" & EndTable.Address)
                AllAddressTable.Add RangeCell.Address(False, False), ""
            Next RangeCell
    
            TotalLook = Round(81 - (81 * .Range("Percent").Value))
            Camultive = WorksheetFunction.CountA(.Range(startTable.Address(False, False) & ":" & EndTable.Address(False, False)))
            NUM = 80
            
            While FullValue - TotalLook <> Camultive
                RandDict = WorksheetFunction.RandBetween(0, NUM)
                Key = AllAddressTable.keys()(RandDict)
    
                If .Range(Key).Value <> "" Then
                    With .Range(Key)
                        .Value = ""
                        .Font.Color = vbBlack
                    End With
                    AllAddressTable.Remove Key
                    Camultive = Camultive - 1
                    NUM = NUM - 1
                End If
            Wend
    
            Application.EnableEvents = True
            Application.ScreenUpdating = True
        End If
    End With
    
    MsgBox "This is game Ready"
    
End Sub


