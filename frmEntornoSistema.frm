VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmEntornoSistema 
   Caption         =   "Entorno del Sistema"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtArquitectura 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   3120
      Width           =   2775
   End
   Begin VB.TextBox txtDireccionIP 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   240
      Width           =   2775
   End
   Begin VB.TextBox txtDominioDNS 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   4560
      Width           =   2775
   End
   Begin VB.TextBox txtDominio 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox txtSesionActiva 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox txtNombreUsuario 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox txtNombreComputadora 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   960
      Width           =   2775
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4680
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label7 
      Caption         =   "Dominio DNS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "Daminio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Arquitectura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Sesión Activa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Nombre Usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre Computadora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Dirección IP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmEntornoSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Me.Height = 5700
  Me.Width = 5895
  '    Call QuitarTitulo

  ' Declaración de variables
  Dim ws As Object
  Dim NombreComputadora As String
  Dim NombreUsuario As String
  Dim Sesion As String
  Dim Arquitectura As String
  Dim Dominio As String
  Dim DominioDNS As String

  'Obtener una variable de tipo "WScript.shell"
  Set ws = CreateObject("WScript.shell")

  'Obtener el valor de la variable de sistema que nos interese mediante la función ExpandEnvironmentStrings()
  NombreComputadora = ws.ExpandEnvironmentStrings("%COMPUTERNAME%")
  NombreUsuario = ws.ExpandEnvironmentStrings("%USERNAME%")
  Sesion = ws.ExpandEnvironmentStrings("%SESSIONNAME%")
  Arquitectura = ws.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
  Dominio = ws.ExpandEnvironmentStrings("%USERDOMAIN%")
  DominioDNS = ws.ExpandEnvironmentStrings("%USERDNSDOMAIN%")

  txtDireccionIP.Text = Winsock1.LocalIP
  txtNombreComputadora.Text = NombreComputadora
  txtNombreUsuario.Text = NombreUsuario
  txtSesionActiva.Text = Sesion
  txtArquitectura.Text = Arquitectura
  txtDominio.Text = Dominio
  txtDominioDNS.Text = DominioDNS

  '    ' Mostrar resultados
  '    MsgBox ("Dirección IP......: " & Winsock1.LocalIP)
  '    MsgBox ("Nombre de Equipo..: " & NombreComputadora)
  '    MsgBox ("Nombre de Usuario.: " & NombreUsuario)
  '    MsgBox ("Sesión............: " & Sesion)
  '    MsgBox ("Arquitectura......: " & Arquitectura)
  '    MsgBox ("Dominio...........: " & Dominio)
  '    MsgBox ("Dominio DNS.......: " & DominioDNS)
End Sub

