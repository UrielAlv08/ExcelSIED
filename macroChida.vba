Sub CopiarDatosACreditosYAmortizacionesYPagos()
    Dim HojaEstructura As Worksheet
    Dim HojaCreditos As Worksheet
    Dim HojaAmortizaciones As Worksheet
    Dim HojaPagos As Worksheet
    Dim UltimaFilaEstructura As Long
    Dim UltimaFilaCreditos As Long
    Dim UltimaFilaAmortizaciones As Long
    Dim UltimaFilaPagos As Long
    Dim i As Long
    
    ' Establece la hoja de origen (Estructura) y destinos (Creditos, Amortizaciones y Pagos)
    Set HojaEstructura = ThisWorkbook.Sheets("Estructura")
    Set HojaCreditos = ThisWorkbook.Sheets("Creditos")
    Set HojaAmortizaciones = ThisWorkbook.Sheets("Amortizaciones")
    Set HojaPagos = ThisWorkbook.Sheets("Pagos")
    
    ' Encuentra la última fila en las hojas Estructura, Creditos, Amortizaciones y Pagos
    UltimaFilaEstructura = HojaEstructura.Cells(HojaEstructura.Rows.Count, "D").End(xlUp).Row
    UltimaFilaCreditos = HojaCreditos.Cells(HojaCreditos.Rows.Count, "A").End(xlUp).Row
    UltimaFilaAmortizaciones = HojaAmortizaciones.Cells(HojaAmortizaciones.Rows.Count, "B").End(xlUp).Row
    UltimaFilaPagos = HojaPagos.Cells(HojaPagos.Rows.Count, "B").End(xlUp).Row
    
    ' Itera a través de las filas de la hoja Estructura
    For i = 2 To UltimaFilaEstructura ' Empezamos desde la segunda fila para omitir encabezados
        ' Copia los datos a la hoja Creditos
        UltimaFilaCreditos = UltimaFilaCreditos + 1
        HojaCreditos.Cells(UltimaFilaCreditos, "A").Value = HojaEstructura.Cells(i, "A").Value ' Nombre
        HojaCreditos.Cells(UltimaFilaCreditos, "B").Value = HojaEstructura.Cells(i, "B").Value ' Creditos
        HojaCreditos.Cells(UltimaFilaCreditos, "C").Value = HojaEstructura.Cells(i, "G").Value ' FechaInicio
        HojaCreditos.Cells(UltimaFilaCreditos, "D").Value = HojaEstructura.Cells(i, "D").Value ' Prestamo
        HojaCreditos.Cells(UltimaFilaCreditos, "E").Value = HojaEstructura.Cells(i, "O").Value ' Interes
        HojaCreditos.Cells(UltimaFilaCreditos, "F").Value = HojaEstructura.Cells(i, "H").Value ' PlazosNum
        HojaCreditos.Cells(UltimaFilaCreditos, "L").Value = HojaEstructura.Cells(i, "I").Value ' PlazosNum

        ' Copia el número de Creditos a la columna "A" en la hoja Amortizaciones
        UltimaFilaAmortizaciones = UltimaFilaAmortizaciones + 1
        HojaAmortizaciones.Cells(UltimaFilaAmortizaciones, "A").Value = HojaEstructura.Cells(i, "B").Value
        HojaAmortizaciones.Cells(UltimaFilaAmortizaciones, "B").Value = HojaEstructura.Cells(i, "D").Value
        HojaAmortizaciones.Cells(UltimaFilaAmortizaciones, "C").Value = HojaEstructura.Cells(i, "E").Value
        HojaAmortizaciones.Cells(UltimaFilaAmortizaciones, "D").Value = HojaEstructura.Cells(i, "E").Value
        HojaAmortizaciones.Cells(UltimaFilaAmortizaciones, "E").Value = HojaEstructura.Cells(i, "F").Value
        HojaAmortizaciones.Cells(UltimaFilaAmortizaciones, "F").Value = HojaEstructura.Cells(i, "F").Value
        HojaAmortizaciones.Cells(UltimaFilaAmortizaciones, "G").Value = HojaEstructura.Cells(i, "G").Value
        HojaAmortizaciones.Cells(UltimaFilaAmortizaciones, "I").Value = HojaEstructura.Cells(i, "G").Value
        HojaAmortizaciones.Cells(UltimaFilaAmortizaciones, "H").Value = HojaEstructura.Cells(i, "H").Value
        HojaAmortizaciones.Cells(UltimaFilaAmortizaciones, "R").Value = HojaEstructura.Cells(i, "F").Value

        ' Copia datos a la hoja Pagos
        ' Copia el número de Creditos a la columna "A" en la hoja Pagos
        UltimaFilaPagos = UltimaFilaPagos + 1
        HojaPagos.Cells(UltimaFilaPagos, "A").Value = HojaEstructura.Cells(i, "B").Value
        HojaPagos.Cells(UltimaFilaPagos, "C").Value = HojaEstructura.Cells(i, "D").Value
        HojaPagos.Cells(UltimaFilaPagos, "C").Value = HojaEstructura.Cells(i, "I").Value
        HojaPagos.Cells(UltimaFilaPagos, "F").Value = HojaEstructura.Cells(i, "J").Value
    Next i
End Sub
