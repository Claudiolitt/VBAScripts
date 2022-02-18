Attribute VB_Name = "Ocultar_Escalas_Todas"
Sub Ocultar_Turmas_EscalaPadrão_FullTime_TODAS()

Dim rng As Range, celula As Range, linha As Integer, linha_cel1 As Integer, linha_cel2 As Integer

Application.ScreenUpdating = False


'''''''''''1
linha_cel1 = 12
linha_cel2 = 20

Set rng = ActiveSheet.Range(Cells(linha_cel1, 3), Cells(linha_cel2, 3))


For Each celula In rng
    If celula = "0" Or celula = "" Then
    celula.Select
    ActiveCell.EntireRow.Hidden = True
    End If
Next celula

'''''''''''2
linha_cel1 = linha_cel1 + 13

linha_cel2 = linha_cel2 + 13


Set rng = ActiveSheet.Range(Cells(linha_cel1, 3), Cells(linha_cel2, 3))

For Each celula In rng
    If celula = "0" Or celula = "" Then
    celula.Select
    ActiveCell.EntireRow.Hidden = True
    End If
Next celula


''''''''''3
linha_cel1 = linha_cel1 + 13

linha_cel2 = linha_cel2 + 13


Set rng = ActiveSheet.Range(Cells(linha_cel1, 3), Cells(linha_cel2, 3))

For Each celula In rng
    If celula = "0" Or celula = "" Then
    celula.Select
    ActiveCell.EntireRow.Hidden = True
    End If
Next celula

''''''''''''4
linha_cel1 = linha_cel1 + 13

linha_cel2 = linha_cel2 + 13


Set rng = ActiveSheet.Range(Cells(linha_cel1, 3), Cells(linha_cel2, 3))

For Each celula In rng
    If celula = "0" Or celula = "" Then
    celula.Select
    ActiveCell.EntireRow.Hidden = True
    End If
Next celula


'''''''''''5
linha_cel1 = linha_cel1 + 13

linha_cel2 = linha_cel2 + 13


Set rng = ActiveSheet.Range(Cells(linha_cel1, 3), Cells(linha_cel2, 3))

For Each celula In rng
    If celula = "0" Or celula = "" Then
    celula.Select
    ActiveCell.EntireRow.Hidden = True
    End If
Next celula

'''''''''''6
linha_cel1 = linha_cel1 + 13

linha_cel2 = linha_cel2 + 13


Set rng = ActiveSheet.Range(Cells(linha_cel1, 3), Cells(linha_cel2, 3))

For Each celula In rng
    If celula = "0" Or celula = "" Then
    celula.Select
    ActiveCell.EntireRow.Hidden = True
    End If
Next celula

'''''''''''7
linha_cel1 = linha_cel1 + 13

linha_cel2 = linha_cel2 + 13


Set rng = ActiveSheet.Range(Cells(linha_cel1, 3), Cells(linha_cel2, 3))

For Each celula In rng
    If celula = "0" Or celula = "" Then
    celula.Select
    ActiveCell.EntireRow.Hidden = True
    End If
Next celula



Application.ScreenUpdating = True

End Sub
