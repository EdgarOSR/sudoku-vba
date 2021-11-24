Attribute VB_Name = "modSudokuArrays"
Option Explicit
Global gCount As Integer, gError As Boolean
Private Const maxCount As Integer = 20000
Dim oGrid As Variant
Dim sGrid(9, 9) As Integer

Private Function SdkGet() As Variant
     
     Dim rg As Range, ac As Range
     Set rg = shSudoku.Range("_SdkOrig")
     
     Dim arr As Variant
     arr = rg.Value
     SdkGet = arr
     
End Function

Private Sub SdkSuccess(ByRef sdk() As Integer)

     Dim rg As Range, cel As Range
     Set rg = shSudoku.Range("_SdkBackup")
     
     Dim i As Integer, j As Integer
     
     For i = 1 To 9
          For j = 1 To 9
               rg.Cells(i, j).Value = sdk(i, j)
          Next j
     Next i
     
     shSudoku.Range("_SdkOrig").Value = rg.Value
     
     shSudoku.Range("M12:S12").Value = Application.WorksheetFunction.Sum(rg)

End Sub

Private Function SdkPossible(ByVal i As Integer, ByVal j As Integer, ByVal n As Integer) As Boolean
     
     Dim r0 As Integer, r As Integer
     Dim c0 As Integer, c As Integer
     
     ' Caso já exista algum valor
     If CInt(oGrid(i, j)) <> 0 Then
          SdkPossible = (oGrid(i, j) = n)
          Exit Function
     End If
     
     ' Caso esteja vazia verifica
     ' se é um valor possível (linha)
     For c = 1 To 9
          If CInt(oGrid(i, c)) = n Then
               SdkPossible = False
               Exit Function
          End If
     Next c
     
     ' Caso esteja vazia verifica
     ' se é um valor possível (coluna)
     For r = 1 To 9
          If oGrid(r, j) = n Then
               SdkPossible = False
               Exit Function
          End If
     Next r
     
     ' Caso esteja vazia verifica
     ' se é um valor possível (bloco)
     r0 = i - ((i - 1) Mod 3)
     c0 = j - ((j - 1) Mod 3)
     
     For r = r0 To r0 + 2
          For c = c0 To c0 + 2
               If oGrid(r, c) = n Then
                    SdkPossible = False
                    Exit Function
               End If
          Next c
     Next r
     
     SdkPossible = True
     
End Function

Private Sub SdkSolver(ByVal i As Integer, j As Integer)
     
     gCount = gCount + 1
     
     If gCount > maxCount Then gError = True: Exit Sub
     
     
     Dim n As Integer, nt As Integer, x As Integer, y As Integer
     
     If i > 9 Then
          For x = 1 To 9
               For y = 1 To 9
                    sGrid(x, y) = oGrid(x, y)
               Next y
          Next x
          Exit Sub
     End If
     
     For n = 1 To 9
          If SdkPossible(i, j, n) Then
               nt = oGrid(i, j)
               oGrid(i, j) = n
               If j = 9 Then
                    SdkSolver i + 1, 1
               Else
                    SdkSolver i, j + 1
               End If
               oGrid(i, j) = nt
          End If
     Next n

End Sub

Public Sub HighlightValues(ByRef listOption As Integer, ByRef template As Boolean)
     
     Dim rg As Range
     Set rg = shSudoku.Range("_SdkOrig")
     
     On Error GoTo EndSub
     
     If template Then
          Select Case listOption
               Case 0
                    shSudoku.Range("_SdkOrig").Value = shTemplates.Range("_SdkEasy").Value
               Case 1
                    shSudoku.Range("_SdkOrig").Value = shTemplates.Range("_SdkMedium").Value
               Case 2
                    shSudoku.Range("_SdkOrig").Value = shTemplates.Range("_SdkHard").Value
               Case 3
                    shSudoku.Range("_SdkOrig").Value = shTemplates.Range("_SdkExpert").Value
               Case 4
                    shSudoku.Range("_SdkOrig").Value = shTemplates.Range("_Sdk1To9").Value
          End Select
     End If
     
     With rg.SpecialCells(xlCellTypeConstants)
          .Font.Bold = True
          .Interior.Color = vbYellow
     End With
     
EndSub:
     Set rg = Nothing
     
End Sub

Private Sub SdkSolver1To9()

     Dim n As Integer
     Dim x As Integer, y As Integer
     Dim arr(8, 8) As Integer
     
     n = 3

     For x = 0 To 8
          For y = 0 To 8
               gCount = gCount + 1
               If x Mod n = 2 Then
                    arr(x, y) = (((x * n) + (x / n) + y - 1) Mod 9) + 1
               Else
                    arr(x, y) = (((x * n) + (x / n) + y) Mod 9) + 1
               End If
          Next y
     Next x
     
     shSudoku.Range("_SdkBackup").Value = arr
     shSudoku.Range("_SdkOrig").Value = arr
     shSudoku.Range("M12:S12").Value = Application.WorksheetFunction.Sum(shSudoku.Range("_SdkOrig"))
     
     gError = False
     
     Erase arr
     
End Sub

Public Sub ClearValues()

     shSudoku.Range("_SdkOrig").ClearContents
     shSudoku.Range("_SdkBackup").ClearContents
     shSudoku.Range("M12:S12").ClearContents
     shSudoku.Range("_SdkOrig").Font.Bold = False
     shSudoku.Range("_SdkOrig").Interior.Color = vbWhite

End Sub

Public Sub SdkMain(ByRef listOption As Integer, ByRef template As Boolean)

     Dim sTime As Double
     sTime = CDbl(Now)
     
     oGrid = SdkGet()
     
     If listOption = 4 Then
          Call SdkSolver1To9
     Else
          Call SdkSolver(1, 1)
     End If

     If gError Then
          MsgBox "Not found a solution" & vbCr & _
          "Duration: " & Format(CDate(CDbl(Now) - sTime), "hh:mm:ss:ms") & vbCr & _
          "Interactions: " & CLng(gCount), vbExclamation
     Else
          MsgBox "Solution found" & vbCr & _
          "Duration: " & Format(CDate(CDbl(Now) - sTime), "hh:mm:ss:ms") & vbCr & _
          "Interactions: " & CLng(gCount), vbInformation
          If listOption <> 4 Then Call SdkSuccess(sGrid)
     End If
     
     gError = False
     gCount = 0
     
End Sub

