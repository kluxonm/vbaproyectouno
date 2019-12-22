VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public tCuota As String, eConstancia As String, eFojas As String, eDireccion As String, eGiro As String, eOficio As String, ePeriodo As String, eRazon As String, eRFC As String, eRP As String, fechaNotificadorIni As String, fechaNotificadorFinal As String, multa As String, multaPeriodo As String, fechaCredito As String, fechaMulta As String, fojasMulta As String, horaDiligenciaIni As String, horaDiligenciaFinal As String, fechaDiligenciaIni As String, fechaDiligenciaFinal As String, municipio As String
Public tCuotaRCV As String, eConstanciaRCV As String, eFojasRCV As String, eDireccionRCV As String, eGiroRCV As String, eOficioRCV As String, ePeriodoRCV As String, eRazonRCV As String, eRfcRCV As String, eRpRCV As String, fechaNotificadorIniRCV As String, fechaNotificadorFinalRCV As String, multaRCV As String, multaPeriodoRCV As String, fechaCreditoRCV As String, fechaMultaRCV As String, fojasMultaRCV As String, horaDiligenciaIniRCV As String, horaDiligenciaFinalRCV As String, fechaDiligenciaIniRCV As String, fechaDiligenciaFinalRCV As String, municipioRCV As String
Public tCuotaSCOP As String, eConstanciaSCOP As String, eFojasSCOP As String, eDireccionSCOP As String, eGiroSCOP As String, eOficioSCOP As String, ePeriodoSCOP As String, eRazonSCOP As String, eRfcSCOP As String, eRpSCOP As String, fechaNotificadorIniSCOP As String, fechaNotificadorFinalSCOP As String, multaSCOP As String, multaPeriodoSCOP As String, fechaCreditoSCOP As String, fechaMultaSCOP As String, fojasMultaSCOP As String, horaDiligenciaIniSCOP As String, horaDiligenciaFinalSCOP As String, fechaDiligenciaIniSCOP As String, fechaDiligenciaFinalSCOP As String, municipioSCOP As String
Public tCuotaSRCV As String, eConstanciaSRCV As String, eFojasSRCV As String, eDireccionSRCV As String, eGiroSRCV As String, eOficioSRCV As String, ePeriodoSRCV As String, eRazonSRCV As String, eRfcSRCV As String, eRpSRCV As String, fechaNotificadorIniSRCV As String, fechaNotificadorFinalSRCV As String, multaSRCV As String, multaPeriodoSRCV As String, fechaCreditoSRCV As String, fechaMultaSRCV As String, fojasMultaSRCV As String, horaDiligenciaIniSRCV As String, horaDiligenciaFinalSRCV As String, fechaDiligenciaIniSRCV As String, fechaDiligenciaFinalSRCV As String, municipioSRCV As String
Public razonGlobal As String, rpGlobal As String, rfcGlobal As String, direccionGlobal As String, municipioGlobal As String
Public llaveCOP As String, llaveRCV As String, llaveSCOP As String, llaveSRCV As String
Public oficioGlobal As String
Public redaccion As String


Sub RunSELECT(textoCuota As String)
    Dim cn As Object, rs As Object, output As String, sql As String, nombre As String
        
    tCuota = textoCuota
    
    '---Driver de la base datos---
    Set cn = CreateObject("ADODB.Connection")
    With cn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .ConnectionString = "Data Source='Aqui va la direccion de la hoja excel a conectar';" & "Extended Properties=""Excel 12.0 Xml;HDR=YES"";" '--Ruta del libro al que se accede para la consulta--
        .Open
    End With
    

      
    
      sql = "SELECT * FROM [Hoja2$] WHERE [CREDITO CUOTA]=" & tCuota
      Set rs = cn.Execute(sql)
      
    If Not rs.EOF Then
        '--Do--
         multa = rs(1)
         eRP = rs(2)
         eRazon = rs(3)
         eRFC = rs(4)
         eDireccion = rs(5)
         municipio = rs(6)
         fechaCredito = rs(7)
         ePeriodo = rs(8)
         eOficio = rs(9)
         oficioGlobal = eOficio
         razonGlobal = eRazon
         rfcGlobal = eRFC
         rpGlobal = eRP
         direccionGlobal = eDireccion
         municipioGlobal = municipio
         llaveCOP = eOficio & eRP
         

        rs.Close
        cn.Close
        Set cn = Nothing
        Set rs = Nothing
    Else
        MsgBox "Número de cuota incorrecto"
         multa = ""
         eRP = ""
         eRazon = ""
         eRFC = ""
         eDireccion = ""
         municipio = ""
         fechaCredito = ""
         ePeriodo = ""
         eOficio = ""
         llaveCOP = ""
        rs.Close
        cn.Close
        Set cn = Nothing
        Set rs = Nothing
     
    End If
    

    
End Sub

Sub consultaRCV(textoCuota As String)
    Dim cn As Object, rs As Object, output As String, sql As String, nombre As String
        
    tCuotaRCV = textoCuota
    
    '---Driver---
    Set cn = CreateObject("ADODB.Connection")
    With cn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .ConnectionString = "Data Source='';" & "Extended Properties=""Excel 12.0 Xml;HDR=YES"";" '--Ruta del libro al que se accede para la consulta--
        .Open
    End With
    
    
      sql = "SELECT * FROM [Hoja2$] WHERE [CREDITO CUOTA]=" & tCuotaRCV
      Set rs = cn.Execute(sql)
      
    If Not rs.EOF Then
      
         multaRCV = rs(1)
         eRpRCV = rs(2)
         eRazonRCV = rs(3)
         eRfcRCV = rs(4)
         eDireccionRCV = rs(5)
         municipioRCV = rs(6)
         fechaCreditoRCV = rs(7)
         ePeriodoRCV = rs(8)
         eOficioRCV = rs(9)
         oficioGlobal = eOficioRCV
         razonGlobal = eRazonRCV
         rfcGlobal = eRfcRCV
         rpGlobal = eRpRCV
         direccionGlobal = eDireccionRCV
         municipioGlobal = municipioRCV
         llaveRCV = eOficioRCV & eRpRCV
         
   
        rs.Close
        cn.Close
        Set cn = Nothing
        Set rs = Nothing
    Else
        MsgBox "Número de cuota incorrecto"
         multaRCV = ""
         eRpRCV = ""
         eRazonRCV = ""
         eRfcRCV = ""
         eDireccionRCV = ""
         municipioRCV = ""
         fechaCreditoRCV = ""
         ePeriodoRCV = ""
         eOficioRCV = ""
         llaveRCV = ""
        rs.Close
        cn.Close
        Set cn = Nothing
        Set rs = Nothing
     
    End If
    
  
    
End Sub


Sub consultaSCOP(textoCuota As String)
    Dim cn As Object, rs As Object, output As String, sql As String, nombre As String
        
    tCuotaSCOP = textoCuota
    
    '---Driver---
    Set cn = CreateObject("ADODB.Connection")
    With cn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .ConnectionString = "Data Source='';" & "Extended Properties=""Excel 12.0 Xml;HDR=YES"";" '--Ruta del libro al que se accede para la consulta--
        .Open
    End With
    
      sql = "SELECT * FROM [Hoja2$] WHERE [CREDITO CUOTA]=" & tCuotaSCOP
      Set rs = cn.Execute(sql)
      
    If Not rs.EOF Then
        '--Do--
         multaSCOP = rs(1)
         eRpSCOP = rs(2)
         eRazonSCOP = rs(3)
         eRfcSCOP = rs(4)
         eDireccionSCOP = rs(5)
         municipioSCOP = rs(6)
         fechaCreditoSCOP = rs(7)
         ePeriodoSCOP = rs(8)
         eOficioSCOP = rs(9)
         oficioGlobal = eOficioSCOP
         razonGlobal = eRazonSCOP
         rfcGlobal = eRfcSCOP
         rpGlobal = eRpSCOP
         direccionGlobal = eDireccionSCOP
         municipioGlobal = municipioSCOP
         llaveSCOP = eOficioSCOP & eRpSCOP
     
        rs.Close
        cn.Close
        Set cn = Nothing
        Set rs = Nothing
    Else
        MsgBox "Número de cuota incorrecto"
         multaSCOP = ""
         eRpSCOP = ""
         eRazonSCOP = ""
         eRfcSCOP = ""
         eDireccionSCOP = ""
         municipioSCOP = ""
         fechaCreditoSCOP = ""
         ePeriodoSCOP = ""
         eOficioSCOP = ""
         llaveSCOP = ""
        rs.Close
        cn.Close
        Set cn = Nothing
        Set rs = Nothing
     
    End If
    


    
End Sub

Sub consultaSRCV(textoCuota As String)
    Dim cn As Object, rs As Object, output As String, sql As String, nombre As String
        
    tCuotaSRCV = textoCuota
    
    '---Driver---
    Set cn = CreateObject("ADODB.Connection")
    With cn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .ConnectionString = "Data Source='';" & "Extended Properties=""Excel 12.0 Xml;HDR=YES"";" '--Ruta del libro al que se accede para la consulta--
        .Open
    End With
    
      sql = "SELECT * FROM [Hoja2$] WHERE [CREDITO CUOTA]=" & tCuotaSRCV
      Set rs = cn.Execute(sql)
      
    If Not rs.EOF Then
        '--Do--
         multaSRCV = rs(1)
         eRpSRCV = rs(2)
         eRazonSRCV = rs(3)
         eRfcSRCV = rs(4)
         eDireccionSRCV = rs(5)
         municipioSRCV = rs(6)
         fechaCreditoSRCV = rs(7)
         ePeriodoSRCV = rs(8)
         eOficioSRCV = rs(9)
         oficioGlobal = eOficioSRCV
         razonGlobal = eRazonSRCV
         rfcGlobal = eRfcSRCV
         rpGlobal = eRpSRCV
         direccionGlobal = eDireccionSRCV
         municipioGlobal = municipioSRCV
         llaveSRCV = eOficioSRCV & eRpSRCV
       
        rs.Close
        cn.Close
        Set cn = Nothing
        Set rs = Nothing
    Else
        MsgBox "Número de cuota incorrecto"
         multaSRCV = ""
         eRpSRCV = ""
         eRazonSRCV = ""
         eRfcSRCV = ""
         eDireccionSRCV = ""
         municipioSRCV = ""
         fechaCreditoSRCV = ""
         ePeriodoSRCV = ""
         eOficioSRCV = ""
         llaveSRCV = ""
        rs.Close
        cn.Close
        Set cn = Nothing
        Set rs = Nothing
     
    End If
    
    '---Cierre de resultset---
    

    
End Sub


Sub generarActaCircunstanciada(localizado As String, resultado As String, visita As String, horaInicial As String, horaFinal As String, inicioDiligencia As String, textoCreditoCOP As String, textoCreditoRCV As String, textoCreditoSCOP As String, textoCreditoSRCV As String)

    Dim ws As Worksheet
    Dim objWord As Object
    Dim i As Integer
   

    FileCopy "Aqui va la ruta del archivo base de Word", "Aca va la ruta donde se guardara una copia, para evitar sobreescribir el base" '--Crear nuevo copia para evitar sobreescriber el archivo base--

    Set objWord = CreateObject("Word.Application")

    objWord.Visible = True
    objWord.Documents.Open "Ruta de la copia creada anteriormente" '--Ruta del nuevo documento word--

    objWord.Activate
 

    With objWord.ActiveDocument.Content.Find
        .Text = "<<CreditoCOP>>"
        .Replacement.Text = textoCreditoCOP
        .Execute Replace:=2
    End With
  'Donde dice Text es la palabra que se va a reemplazar  con el metdo Replacement.Text, se pueden añadir tantos Content.Find como se necesite
    If Not Len(textoCreditoRCV) = 0 Then
    
        With objWord.ActiveDocument.Content.Find
            .Text = "<<Palabra>>"
            .Replacement.Text = Left(palabra, Round((Len(palabra) / 2) + 2, 0)) 'Operación para evitar la límitación de los 255 caracteres
            .Execute Replace:=2
        End With
    
     With objWord.ActiveDocument.Content.Find
            .Text = "<<Palabra>>"
            .Replacement.Text = Right(palabra, Round((Len(palabra) / 2) - 2, 0)) 'Operación para evitar la límitación de los 255 caracteres
            .Execute Replace:=2
        End With
    
    Else
        
        With objWord.ActiveDocument.Content.Find
            .Text = "<<palabra>>"
            .Replacement.Text = ""
            .Execute Replace:=2
        End With
    
     With objWord.ActiveDocument.Content.Find
            .Text = "<<palabra>>"
            .Replacement.Text = ""
            .Execute Replace:=2
        End With
    End If
        


End Sub

'--Este metodo hace lo mismo que el anterior, es para el boton 2--
Sub generarActaTestigos(notificado As String, resultado As String, redaccion As String, fechaDiligencia As String)
    Dim ws As Worksheet
    Dim objWord As Object
    Dim i As Integer
    Dim strValue As String



End Sub

