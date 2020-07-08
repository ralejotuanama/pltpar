VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Feriad_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   2325
   ClientLeft      =   8640
   ClientTop       =   5475
   ClientWidth     =   5715
   Icon            =   "PltPar_frm_042.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2355
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5745
      _Version        =   65536
      _ExtentX        =   10134
      _ExtentY        =   4154
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel SSPanel2 
         Height          =   825
         Left            =   30
         TabIndex        =   5
         Top             =   1470
         Width           =   5655
         _Version        =   65536
         _ExtentX        =   9975
         _ExtentY        =   1455
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin VB.TextBox txt_DesFer 
            Height          =   315
            Left            =   1350
            MaxLength       =   30
            TabIndex        =   2
            Top             =   420
            Width           =   4275
         End
         Begin EditLib.fpDateTime ipp_FecFer 
            Height          =   315
            Left            =   1350
            TabIndex        =   1
            Top             =   60
            Width           =   4275
            _Version        =   196608
            _ExtentX        =   7541
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   3
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "01/01/2008"
            DateCalcMethod  =   0
            DateTimeFormat  =   0
            UserDefinedFormat=   ""
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label2 
            Caption         =   "Descripción:"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   450
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Feriado:"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   825
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   5655
         _Version        =   65536
         _ExtentX        =   9975
         _ExtentY        =   1191
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin Threed.SSPanel pnl_TitPri 
            Height          =   315
            Left            =   660
            TabIndex        =   9
            Top             =   210
            Width           =   2775
            _Version        =   65536
            _ExtentX        =   4895
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Mantenimiento Feriados"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.76
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   90
            Picture         =   "PltPar_frm_042.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   735
         Left            =   30
         TabIndex        =   10
         Top             =   720
         Width           =   5655
         _Version        =   65536
         _ExtentX        =   9975
         _ExtentY        =   1296
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   30
            Picture         =   "PltPar_frm_042.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salir 
            Height          =   675
            Left            =   4950
            Picture         =   "PltPar_frm_042.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frm_Feriad_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Grabar_Click()
   
   'Validacion para que no sea NULL
   If Len(Trim(txt_DesFer.Text)) = 0 Then
      MsgBox "Ingrese una descripción.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_DesFer)
      Exit Sub
   End If
              
   'Obteniendo Información del Registro, se compara la BD con nuestros controles
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM OPE_DIAFER WHERE "
   g_str_Parame = g_str_Parame & "DIAFER_DIAFER = " & Format(CDate(ipp_FecFer.Text), "yyyymmdd")
   
   'Se hace la conexion a la base datos, se envia la cadena, ADO, Modalidad
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
      
   'Se evalua la data existente
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   
      'Se envia el mensaje si el codigo existe
      MsgBox "La Fecha que desea Ingresar ya se encuentra en el sistema.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
                            
   'Se cierra la conexion a la base de datos
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
      
   'Envia mensaje con confirmacion de grabado de datos
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Puntero con reloj de arena
   Screen.MousePointer = vbHourglass
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
                     
   'Se llama al procedure y se ejecuta el ingreso de la data en la base de datos
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_OPE_DIAFER ("
      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecFer.Text), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & "'" & txt_DesFer.Text & "') "
                              
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
                                 
      'Se genera el mensaje de error por la concurrencia que exista
      If moddat_g_int_CntErr = 5 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   'Puntero Normal
   Screen.MousePointer = vbDefault
   
   'Se muestra el mensaje con el grabado exitoso
   MsgBox "Se grabaron los datos correctamente.", vbInformation, modgen_g_str_NomPlt
   Unload Me
   
End Sub

Private Sub cmd_Salir_Click()

   'Cerrado de ventana
   Unload Me
   
End Sub

Private Sub Form_Load()

   'Centrar el Formulario
   Call gs_CentraForm(Me)
   
   Me.Caption = modgen_g_str_NomPlt
   ipp_FecFer.Text = (date)
      
End Sub

Private Sub ipp_FecFer_KeyPress(KeyAscii As Integer)

   'Se envia el curso al siguiente control
   Call gs_SetFocus(txt_DesFer)
   
End Sub
