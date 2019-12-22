VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form 
   Caption         =   "Asistente de Actas"
   ClientHeight    =   8700.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   16668
   OleObjectBlob   =   "FormTwo.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bNotificado As String, bResultado As String, bVisita As String, creditoCOP As String, creditoRCV As String, creditoSCOP As String, creditoSRCV As String
Public evaluate As String, comparator As Boolean
Public redaccion As String


Private Sub botonConsultaRCV_Click()

Dim tc As Integer

tc = StrComp(Mid(textoCuotaRCV.Value, 1, 3), "197")

If Not Len(textoCuotaRCV.Value) >= 9 Then
    MsgBox "El número debe contener 9 caracteres"

ElseIf tc <> 0 Then

MsgBox "Número corresponde a otro tipo de cuota"

Else

    ThisWorkbook.consultaRCV textoCuotaRCV.Value
    
    etiquetaRazonRCV.Caption = ThisWorkbook.eRazonRCV
    etiquetaMultaRCV.Caption = ThisWorkbook.multaRCV
    etiquetaNotificadorRCV.Caption = ThisWorkbook.eOficioRCV
    '--etiquetaNumeroConstancia.Caption = ThisWorkbook.--
    
    textoCuotaRCV.Enabled = False

End If

End Sub

Private Sub botonConsultaSCOP_Click()


Dim tc As Integer

tc = StrComp(Mid(textoCuotaSCOP.Value, 1, 3), "193")

If Not Len(textoCuotaSCOP.Value) >= 9 Then
    MsgBox "El número debe contener 9 caracteres"
    
ElseIf tc <> 0 Then

MsgBox "Número corresponde a otro tipo de cuota"
    
Else

    ThisWorkbook.consultaSCOP textoCuotaSCOP.Value
    
    etiquetaRazonSCOP.Caption = ThisWorkbook.eRazonSCOP
    etiquetaMultaSCOP.Caption = ThisWorkbook.multaSCOP
    etiquetaNotificadorSCOP.Caption = ThisWorkbook.eOficioSCOP
    '--etiquetaNumeroConstancia.Caption = ThisWorkbook.--
    
    textoCuotaSCOP.Enabled = False

End If

End Sub

Private Sub botonConsultaSRCV_Click()

Dim tc As Integer

tc = StrComp(Mid(textoCuotaSRCV.Value, 1, 3), "193")

If Not Len(textoCuotaSRCV.Value) >= 9 Then
    MsgBox "El número debe contener 9 caracteres"
    
ElseIf tc <> 0 Then

MsgBox "Número corresponde a otro tipo de cuota"

    
Else

    ThisWorkbook.consultaSRCV textoCuotaSRCV.Value
    
    etiquetaRazonSRCV.Caption = ThisWorkbook.eRazonSRCV
    etiquetaMultaSRCV.Caption = ThisWorkbook.multaSRCV
    etiquetaNotificadorSRCV.Caption = ThisWorkbook.eOficioSRCV
    '--etiquetaNumeroConstancia.Caption = ThisWorkbook.--
    
    textoCuotaSRCV.Enabled = False

End If

End Sub

Private Sub botonNL_Click()

botonPrimeraVisita.Enabled = True

End Sub

Private Sub botonNN_Click()

botonPrimeraVisita.Enabled = False
botonSegundaVisita.Value = True

End Sub

Private Sub botonNuevaConsultaCOP_Click()

textoCuota.Value = ""

etiquetaOficio.Caption = ""
etiquetaRazon.Caption = ""
fojasTexto.Value = ""
etiquetaMulta.Caption = ""

textoCuota.Enabled = True

End Sub


Private Sub botonNuevaConsultaRCV_Click()

textoCuotaRCV.Value = ""

etiquetaNotificadorRCV.Caption = ""
etiquetaRazonRCV.Caption = ""
textoFojasRCV.Value = ""
etiquetaMultaRCV.Caption = ""

textoCuotaRCV.Enabled = True

End Sub

Private Sub botonNuevaConsultaSCOP_Click()

textoCuotaSCOP.Value = ""

etiquetaNotificadorSCOP.Caption = ""
etiquetaRazonSCOP.Caption = ""
textoFojasSCOP.Value = ""
etiquetaMultaSCOP.Caption = ""

textoCuotaSCOP.Enabled = True

End Sub

Private Sub botonNuevaConsultaSRCV_Click()

textoCuotaSRCV.Value = ""

etiquetaNotificadorSRCV.Caption = ""
etiquetaRazonSRCV.Caption = ""
textoFojasSRCV.Value = ""
etiquetaMultaSRCV.Caption = ""

textoCuotaSRCV.Enabled = True

End Sub

Private Sub CheckBox1_Click()

If CheckBox1 = True Then
    textoCuota.Enabled = True
    fojasTexto.Enabled = True
    botonNuevaConsultaCOP.Enabled = True
    botonConsultaCOP.Enabled = True
    evaluate = evaluate & "0"
Else
    textoCuota.Enabled = False
    fojasTexto.Enabled = False
    botonNuevaConsultaCOP.Enabled = False
    botonConsultaCOP.Enabled = False
    evaluate = Replace(evaluate, "0", "")
End If



End Sub

Sub botonConsultaCOP_Click()

Dim tc As Integer

tc = StrComp(Mid(textoCuota.Value, 1, 3), "192")

If Not Len(textoCuota.Value) >= 9 Then
    MsgBox "El número debe contener 9 caracteres"

ElseIf tc <> 0 Then
    
    MsgBox "Número corresponde a otro tipo de cuota"
    
Else

    ThisWorkbook.RunSELECT textoCuota.Value
    
    etiquetaRazon.Caption = ThisWorkbook.eRazon
    etiquetaMulta.Caption = ThisWorkbook.multa
    etiquetaOficio.Caption = ThisWorkbook.eOficio
    '--etiquetaNumeroConstancia.Caption = ThisWorkbook.--
    
    textoCuota.Enabled = False

End If

End Sub



Private Sub CheckBox2_Click()

If CheckBox2 = True Then
    textoCuotaRCV.Enabled = True
    textoFojasRCV.Enabled = True
    botonNuevaConsultaRCV.Enabled = True
    botonConsultaRCV.Enabled = True
    evaluate = evaluate & "1"
Else
     textoCuotaRCV.Enabled = False
     textoFojasRCV.Enabled = False
     botonNuevaConsultaRCV.Enabled = False
     botonConsultaRCV.Enabled = False
     evaluate = Replace(evaluate, "1", "")
End If


End Sub

Private Sub CheckBox3_Click()

If CheckBox3 = True Then
    textoCuotaSCOP.Enabled = True
    textoFojasSCOP.Enabled = True
    botonNuevaConsultaSCOP.Enabled = True
    botonConsultaSCOP.Enabled = True
    evaluate = evaluate & "2"
Else
     textoCuotaSCOP.Enabled = False
     textoFojasSCOP.Enabled = False
     botonNuevaConsultaSCOP.Enabled = False
     botonConsultaSCOP.Enabled = False
     evaluate = Replace(evaluate, "2", "")
End If

End Sub

Private Sub CheckBox4_Click()

If CheckBox4 = True Then
    textoCuotaSRCV.Enabled = True
    textoFojasSRCV.Enabled = True
    botonNuevaConsultaSRCV.Enabled = True
    botonConsultaSRCV.Enabled = True
    evaluate = evaluate & "3"
Else
     textoCuotaSRCV.Enabled = False
     textoFojasSRCV.Enabled = False
     botonNuevaConsultaSRCV.Enabled = False
     botonConsultaSRCV.Enabled = False
     evaluate = Replace(evaluate, "3", "")
End If

End Sub


Private Sub generarActaCircunstanciada_Click()

textoHoraInicial.Enabled = False
textoHoraFinal.Enabled = False
textoInicioDiligencia.Enabled = False
botonNN.Enabled = False
botonNL.Enabled = False
botonPrimeraVisita.Enabled = False
botonSegundaVisita.Enabled = False


capturaHechos.Show

comparator = False

Dim compararCOP As String, comparaRCV As String, compararSCOP As String, compararSRCV As String, arr(4) As String, x As Integer, y As Integer

If CheckBox1 = True Then
 creditoCOP = " Aqui iba un texto para el metodo generarActa "
 compararCOP = ThisWorkbook.llaveCOP
Else
 creditoCOP = ""
 compararCOP = ""
End If

If CheckBox2 = True Then
    creditoRCV = " "
    compararRCV = ThisWorkbook.llaveRCV
Else
 creditoRCV = ""
 compararRCV = ""
End If

If CheckBox3 = True Then
    creditoSCOP = "  "
    compararSCOP = ThisWorkbook.llaveSCOP
Else
 creditoSCOP = ""
 compararSCOP = ""
End If

If CheckBox4 = True Then
    creditoSRCV = "  "
    compararSRCV = ThisWorkbook.llaveSRCV
Else
 creditoSRCV = ""
 compararSRCV = ""
End If

If botonNL = True Then
    bNotificado = "Aqui tambien iba un texto"
    bResultado = ""
Else
    bNotificado = ""
    bResultado = ""
End If

If botonPrimeraVisita = True Then
    bVisita = ""
Else
    bVisita = ""
End If

'Verificacion

arr(0) = compararCOP
arr(1) = compararRCV
arr(2) = compararSCOP
arr(3) = compararSRCV

    If Len(evaluate) = 0 Then
        
        comparator = False
    
    ElseIf Len(evaluate) = 1 Then
        
        comparator = True
        
    ElseIf Len(evaluate) > 1 Then
          
          comparator = True
           
          For i = 1 To Len(evaluate)
          
            x = CInt(Mid(evaluate, i, 1))
            
            If Not StrComp(Mid(evaluate, i + 1, 1), "", vbBinaryCompare) = 0 Then
            
                y = CInt(Mid(evaluate, i + 1, 1))
            Else
                Exit For
            End If
          
            If StrComp(arr(x), arr(y), vbBinaryCompare) = 0 Then
                comparator = True
            Else
                comparator = False
                Exit For
            End If
          Next i
    End If

    
If comparator = True Then
   ThisWorkbook.generarActaCircunstanciada bNotificado, bResultado, bVisita, textoHoraInicial, textoHoraFinal, textoInicioDiligencia, creditoCOP, creditoRCV, creditoSCOP, creditoSRCV
   generarActaTestigos.Enabled = True
Else
    MsgBox "Notificador o registro patronal no concuerdan (Error 1)."
End If

End Sub

Private Sub generarActaTestigos_Click()

If botonNL.Value = True Then
    bNotificado = "Aqui iba un texto"
    bResultado = ""
Else
    bNotificado = ""
    bResultado = ""
End If


If botonPrimeraVisita.Value = True Then
    bVisita = ""
Else
    bVisita = ""
End If

ThisWorkbook.generarActaTestigos bNotificado, bResultado, redaccion, textoInicioDiligencia.Value

End Sub




Private Sub OptionButton2_Click()

End Sub

Private Sub OptionButton3_Click()

End Sub

Private Sub OptionButton4_Click()

End Sub

Private Sub reiniciar_Click()

capturaHechos.textoRedaccion.Value = "" 'Reiniciar el Frame
Unload Form
Form.Show


End Sub

Private Sub textoInicioDiligencia_Enter()

textoInicioDiligencia.Value = ""

End Sub

Private Sub UserForm_Click()

End Sub
