VERSION 5.00
Begin VB.Form frm2 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calculadora"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9930
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   9930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboOperacion 
      Height          =   315
      ItemData        =   "frm2.frx":0000
      Left            =   3000
      List            =   "frm2.frx":0010
      TabIndex        =   7
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox txtresultado 
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton cmdOperar 
      Caption         =   "operar"
      Height          =   735
      Left            =   3120
      TabIndex        =   5
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox txt2 
      Height          =   285
      Left            =   3240
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Left            =   3240
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Seleccione la operacion"
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Resultado"
      Height          =   195
      Left            =   1200
      TabIndex        =   4
      Top             =   1920
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Numero 2"
      Height          =   195
      Left            =   1200
      TabIndex        =   1
      Top             =   1080
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Numero 1"
      Height          =   195
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   690
   End
End
Attribute VB_Name = "frm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOperar_Click()
'ccur() otra opcion
Dim operacion As Integer
operacion = cboOperacion.ListIndex

If Len(txt1.Text) = 0 Then
    MsgBox "Debe ingresar numeros"
    txt1.SetFocus ' colocar el foco en el txt1
    Exit Sub
    
End If

If Not IsNumeric(txt1.Text) Then
    MsgBox "debe ingresar solo numeros"
    txt1.SetFocus ' colocar el foco en el txt1
    Exit Sub
End If




If Len(txt2.Text) = 0 Then
    MsgBox "Debe ingresar numeros"
    txt2.SetFocus ' colocar el foco en el txt2
    Exit Sub
    
End If

If Not IsNumeric(txt2.Text) Then
    MsgBox "debe ingresar solo numeros"
    txt2.SetFocus ' colocar el foco en el txt2
    Exit Sub
End If

'validar seleccion de operacion

If operacion = -1 Then
    MsgBox "Debe seleccionar una operacion"
    cboOperacion.SetFocus
    Exit Sub
End If


If operacion = 0 Then

'suma
txtresultado.Text = CStr(Val(txt1.Text) + Val(txt2.Text))

Exit Sub

End If

If operacion = 1 Then

'resta
txtresultado.Text = CStr(Val(txt1.Text) - Val(txt2.Text))

Exit Sub

End If

If operacion = 2 Then

'multip
txtresultado.Text = CStr(Val(txt1.Text) * Val(txt2.Text))

Exit Sub

End If

If operacion = 3 Then

'division
txtresultado.Text = CStr(Val(txt1.Text) / Val(txt2.Text))

Exit Sub

End If




End Sub

Private Sub Combo1_Change()

End Sub
