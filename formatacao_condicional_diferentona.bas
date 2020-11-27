Attribute VB_Name = "Módulo1"
Dim cod_log1 As Integer
Dim cod_log2 As Integer

Dim rg As Range
Dim cs As ColorScale
Sub VerificarCodLog()

cod_log1 = Planilha2.Cells(2, 1).Value
cod_log2 = Planilha2.Cells(3, 1).Value

i = 1
i1 = 2
i2 = 3

comeco_int = 2

Do While cod_log2 <> Empty

i = i + 1
    
    If cod_log2 <> cod_log1 Then
        
        fim_int = i
        
        Range("A" & comeco_int & ":F" & fim_int).Sort Key1:=Range("F" & comeco_int & ":F" & fim_int), Order1:=xlAscending, Header:=xlNo
        
        Set rg = Range("B" & comeco_int & ":B" & fim_int)
        
        rg.FormatConditions.Delete
        'colour scale will have three colours
        Set cs = rg.FormatConditions.AddColorScale(ColorScaleType:=3)
        
        cs.ColorScaleCriteria(1).FormatColor.Color = RGB(51, 255, 51)
        cs.ColorScaleCriteria(2).FormatColor.Color = RGB(255, 230, 153)
        cs.ColorScaleCriteria(3).FormatColor.Color = RGB(255, 51, 0)
        
        Set rg = Range("C" & comeco_int & ":C" & fim_int)
               
        rg.FormatConditions.Delete
        'colour scale will have three colours
        Set cs = rg.FormatConditions.AddColorScale(ColorScaleType:=3)
        
        cs.ColorScaleCriteria(1).FormatColor.Color = RGB(51, 255, 51)
        cs.ColorScaleCriteria(2).FormatColor.Color = RGB(255, 230, 153)
        cs.ColorScaleCriteria(3).FormatColor.Color = RGB(255, 51, 0)
        
        Set rg = Range("D" & comeco_int & ":D" & fim_int)
               
        rg.FormatConditions.Delete
        'colour scale will have three colours
        Set cs = rg.FormatConditions.AddColorScale(ColorScaleType:=3)
        
        cs.ColorScaleCriteria(1).FormatColor.Color = RGB(51, 255, 51)
        cs.ColorScaleCriteria(2).FormatColor.Color = RGB(255, 230, 153)
        cs.ColorScaleCriteria(3).FormatColor.Color = RGB(255, 51, 0)
        
        Set rg = Range("E" & comeco_int & ":E" & fim_int)
               
        rg.FormatConditions.Delete
        'colour scale will have three colours
        Set cs = rg.FormatConditions.AddColorScale(ColorScaleType:=3)
        
        cs.ColorScaleCriteria(1).FormatColor.Color = RGB(51, 255, 51)
        cs.ColorScaleCriteria(2).FormatColor.Color = RGB(255, 230, 153)
        cs.ColorScaleCriteria(3).FormatColor.Color = RGB(255, 51, 0)
        
        Set rg = Range("A" & comeco_int & ":F" & fim_int)
        rg.BorderAround ColorIndex:=1, Weight:=xlThick
               
        comeco_int = fim_int + 1
                            
    End If
    
    i1 = i1 + 1
    i2 = i2 + 1
    
    cod_log1 = Planilha2.Cells(i1, 1).Value
    cod_log2 = Planilha2.Cells(i2, 1).Value

Loop

End Sub


