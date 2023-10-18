Sub CopiarDatosAContratoYNumerosYPagos()
    Dim HojaPrincipal As Worksheet
    Dim HojaContrato As Worksheet
    Dim HojaNumeros As Worksheet
    Dim HojaPagos As Worksheet
    Dim UltimaFilaPrincipal As Long
    Dim UltimaFilaContrato As Long
    Dim UltimaFilaNumeros As Long
    Dim UltimaFilaPagos As Long
    Dim i As Long
    
    ' Establece la hoja de origen (Principal) y destinos (Contrato, Numeros y Pagos)
    Set HojaPrincipal = ThisWorkbook.Sheets("Principal")
    Set HojaContrato = ThisWorkbook.Sheets("Contrato")
    Set HojaNumeros = ThisWorkbook.Sheets("Numeros")
    Set HojaPagos = ThisWorkbook.Sheets("Pagos")
    
    ' Encuentra la última fila en las hojas Principal, Contrato, Numeros y Pagos
    UltimaFilaPrincipal = HojaPrincipal.Cells(HojaPrincipal.Rows.Count, "D").End(xlUp).Row
    UltimaFilaContrato = HojaContrato.Cells(HojaContrato.Rows.Count, "A").End(xlUp).Row
    UltimaFilaNumeros = HojaNumeros.Cells(HojaNumeros.Rows.Count, "B").End(xlUp).Row
    UltimaFilaPagos = HojaPagos.Cells(HojaPagos.Rows.Count, "B").End(xlUp).Row
    
    ' Itera a través de las filas de la hoja Principal
    For i = 2 To UltimaFilaPrincipal ' Empezamos desde la segunda fila para omitir encabezados
        ' Copia los datos a la hoja Contrato
        UltimaFilaContrato = UltimaFilaContrato + 1
        HojaContrato.Cells(UltimaFilaContrato, "A").Value = HojaPrincipal.Cells(i, "A").Value ' Nombre
        HojaContrato.Cells(UltimaFilaContrato, "B").Value = HojaPrincipal.Cells(i, "B").Value ' Contrato
        HojaContrato.Cells(UltimaFilaContrato, "C").Value = HojaPrincipal.Cells(i, "L").Value ' FechaInicio
        HojaContrato.Cells(UltimaFilaContrato, "D").Value = HojaPrincipal.Cells(i, "M").Value ' Prestamo
        HojaContrato.Cells(UltimaFilaContrato, "E").Value = HojaPrincipal.Cells(i, "C").Value ' Interes
        HojaContrato.Cells(UltimaFilaContrato, "F").Value = HojaPrincipal.Cells(i, "N").Value ' PlazosNum

        ' Copia el número de contrato a la columna "A" en la hoja Numeros
        UltimaFilaNumeros = UltimaFilaNumeros + 1
        HojaNumeros.Cells(UltimaFilaNumeros, "A").Value = HojaPrincipal.Cells(i, "B").Value
        ' Copia los datos de la columna "Periodos" a la columna "Periodo" en la hoja Numeros
        HojaNumeros.Cells(UltimaFilaNumeros, "B").Value = HojaPrincipal.Cells(i, "D").Value
        ' Copia los datos de la columna "Fechas" en la hoja Principal a la columna "Fecha" en la hoja Numeros
        HojaNumeros.Cells(UltimaFilaNumeros, "C").Value = HojaPrincipal.Cells(i, "E").Value
        ' Copia los datos de la columna "Saldo" en la hoja Principal a la columna "Dinero" en la hoja Numeros
        HojaNumeros.Cells(UltimaFilaNumeros, "D").Value = HojaPrincipal.Cells(i, "F").Value
        ' Copia los datos de la columna "Moratorios" en la hoja Principal a la columna "Moratorios" en la hoja Numeros
        HojaNumeros.Cells(UltimaFilaNumeros, "E").Value = HojaPrincipal.Cells(i, "G").Value
        ' Copia los datos de la columna "IVA" en la hoja Principal a la columna "IVA" en la hoja Numeros
        HojaNumeros.Cells(UltimaFilaNumeros, "F").Value = HojaPrincipal.Cells(i, "H").Value
        ' Copia los datos de la columna "SaldoInsol" en la hoja Principal a la columna "SaldoInsol" en la hoja Numeros
        HojaNumeros.Cells(UltimaFilaNumeros, "G").Value = HojaPrincipal.Cells(i, "I").Value

        ' Copia datos a la hoja Pagos
        ' Copia el número de contrato a la columna "A" en la hoja Pagos
        UltimaFilaPagos = UltimaFilaPagos + 1
        HojaPagos.Cells(UltimaFilaPagos, "A").Value = HojaPrincipal.Cells(i, "B").Value
        ' Copia los datos de la columna "Periodos" en la hoja Principal a la columna "Periodo" en la hoja Pagos
        HojaPagos.Cells(UltimaFilaPagos, "B").Value = HojaPrincipal.Cells(i, "D").Value
        HojaPagos.Cells(UltimaFilaPagos, "C").Value = HojaPrincipal.Cells(i, "I").Value
        HojaPagos.Cells(UltimaFilaPagos, "D").Value = HojaPrincipal.Cells(i, "J").Value
    Next i
End Sub
