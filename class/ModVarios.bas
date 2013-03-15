Attribute VB_Name = "ModVarios"

Public Sub exportarExcel(RSSQL1 As ADODB.Recordset, titulo As String)
On Error GoTo ErrorExcel
Dim objExcel As Excel.Application
Dim HNom As Integer 'Horizontal
Dim VNom As Integer 'Vertical
Dim Hdatos As Integer 'Horizontal
Dim Vdatos As Integer 'Vertical
Dim cuentaNombres As Integer
Dim cuentadatos As Integer
Dim i As Integer
Dim n As Integer
Dim J As Integer

If RSSQL1.RecordCount <> 0 Then
   'Crear un objeto del tipo excel.application

   cuentaNombres = RSSQL1.Fields.Count
   cuentadatos = RSSQL1.RecordCount

   Set objExcel = New Excel.Application
   objExcel.Visible = True
   objExcel.SheetsInNewWorkbook = 1
   objExcel.Workbooks.Add

    'PONER UN TITULO
    objExcel.ActiveSheet.Cells(1, 1) = "EXPORTAR A EXCEL - " + titulo
    objExcel.ActiveSheet.Cells(2, 1) = cuentadatos
    objExcel.ActiveSheet.Cells(2, 2) = cuentaNombres
    With objExcel.ActiveSheet.Cells(1, 1).Font
      .Color = vbBlack
      .Size = 12
      .Bold = True
   End With

   'UTILIZAMOS LAS VARIABLES PARA LA UBICACION DE NUESTROS TEXTOS
   HNom = 1
   VNom = 4
   Vdatos = 5
   Hdatos = 1


   'AGREGAMOS LOS REGISTROS (RECUERDEN QUE NO IMPORTA CUANTAS COLUMNAS O REGISTROS TENGAMOS EL BUCLE_
   'FUNCIONA SEGUN EL NUMERO DE CABECERAS Y REGISTROS
  
    For i = 0 To (cuentaNombres - 1)
       objExcel.ActiveSheet.Cells(VNom, HNom) = RSSQL1.Fields(i).Name
       objExcel.ActiveSheet.Range(objExcel.ActiveSheet.Cells(VNom, HNom), objExcel.ActiveSheet.Cells(VNom, HNom)).HorizontalAlignment = xlHAlignCenterAcrossSelection
       With objExcel.ActiveSheet.Cells(VNom, HNom).Font
          .Size = 12
          .Color = vbRed
          .Bold = True
       End With
       RSSQL1.MoveFirst
       For n = 1 To RSSQL1.RecordCount
         objExcel.ActiveSheet.Cells(Vdatos, Hdatos) = RSSQL1.Fields(i).Value
         objExcel.ActiveSheet.Cells(Vdatos, Hdatos).Font.Size = 10
         Vdatos = Vdatos + 1
         RSSQL1.MoveNext
       Next
       HNom = HNom + 1
       Hdatos = Hdatos + 1
       Vdatos = 5
       RSSQL1.MoveFirst
   Next i
   'AHORA LE ASIGNAMOS UN TAMAÑO A CADA COLUMNA SEGUN NESECITEMOS
    objExcel.Columns("B").ColumnWidth = 15.43
    objExcel.Columns("C").ColumnWidth = 15.43
    objExcel.Columns("D").ColumnWidth = 25.86
    objExcel.Columns("E").ColumnWidth = 15.83
End If
Exit Sub
ErrorExcel:
MsgBox Err.Description
End Sub


