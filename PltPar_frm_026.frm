VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Carter_2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8040
   ClientLeft      =   3105
   ClientTop       =   1740
   ClientWidth     =   8175
   Icon            =   "PltPar_frm_026.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8025
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      _Version        =   65536
      _ExtentX        =   14420
      _ExtentY        =   14155
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
      Begin Threed.SSPanel SSPanel9 
         Height          =   795
         Left            =   30
         TabIndex        =   1
         Top             =   780
         Width           =   8085
         _Version        =   65536
         _ExtentX        =   14261
         _ExtentY        =   1402
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
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   5940
            Picture         =   "PltPar_frm_026.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Buscar Datos"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   6660
            Picture         =   "PltPar_frm_026.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Limpiar Datos"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   7350
            Picture         =   "PltPar_frm_026.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir"
            Top             =   60
            Width           =   675
         End
         Begin VB.ComboBox cmb_Carter 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   180
            Width           =   4095
         End
         Begin VB.Label Label1 
            Caption         =   "Seleccione Cartera:"
            Height          =   285
            Left            =   90
            TabIndex        =   6
            Top             =   210
            Width           =   1815
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   765
         Left            =   30
         TabIndex        =   7
         Top             =   5580
         Width           =   8085
         _Version        =   65536
         _ExtentX        =   14261
         _ExtentY        =   1349
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
         Begin VB.CommandButton cmd_Editar 
            Height          =   675
            Left            =   6690
            Picture         =   "PltPar_frm_026.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   675
            Left            =   6000
            Picture         =   "PltPar_frm_026.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   675
            Left            =   7380
            Picture         =   "PltPar_frm_026.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   735
         Left            =   30
         TabIndex        =   11
         Top             =   7230
         Width           =   8085
         _Version        =   65536
         _ExtentX        =   14261
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
            Left            =   6690
            Picture         =   "PltPar_frm_026.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Cancel 
            Height          =   675
            Left            =   7380
            Picture         =   "PltPar_frm_026.frx":17C2
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   795
         Left            =   30
         TabIndex        =   14
         Top             =   6390
         Width           =   8085
         _Version        =   65536
         _ExtentX        =   14261
         _ExtentY        =   1402
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
         Begin VB.ComboBox cmb_SecEco 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   60
            Width           =   6405
         End
         Begin EditLib.fpDoubleSingle ipp_Porcen 
            Height          =   315
            Left            =   1620
            TabIndex        =   30
            Top             =   390
            Width           =   1515
            _Version        =   196608
            _ExtentX        =   2672
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
            ButtonStyle     =   0
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
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "100"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin Threed.SSPanel pnl_MetSec 
            Height          =   315
            Left            =   6600
            TabIndex        =   31
            Top             =   390
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "2,000,000.00 "
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin VB.Label Label6 
            Caption         =   "Meta x Sector:"
            Height          =   285
            Left            =   5220
            TabIndex        =   32
            Top             =   390
            Width           =   1275
         End
         Begin VB.Label Label3 
            Caption         =   "Sector Económico:"
            Height          =   285
            Left            =   90
            TabIndex        =   16
            Top             =   120
            Width           =   1485
         End
         Begin VB.Label Label2 
            Caption         =   "% Colocación:"
            Height          =   285
            Left            =   90
            TabIndex        =   15
            Top             =   450
            Width           =   1275
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   3915
         Left            =   30
         TabIndex        =   17
         Top             =   1620
         Width           =   8085
         _Version        =   65536
         _ExtentX        =   14261
         _ExtentY        =   6906
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   2685
            Left            =   60
            TabIndex        =   20
            Top             =   870
            Width           =   7905
            _ExtentX        =   13944
            _ExtentY        =   4736
            _Version        =   393216
            Rows            =   12
            Cols            =   4
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_TotPor 
            Height          =   285
            Left            =   4380
            TabIndex        =   26
            Top             =   3570
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Alignment       =   4
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   4380
            TabIndex        =   18
            Top             =   570
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "% Coloc."
            ForeColor       =   16777215
            BackColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   285
            Left            =   90
            TabIndex        =   19
            Top             =   570
            Width           =   4305
            _Version        =   65536
            _ExtentX        =   7594
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Sector Económico"
            ForeColor       =   16777215
            BackColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   285
            Left            =   5910
            TabIndex        =   21
            Top             =   570
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Importe Colocación"
            ForeColor       =   16777215
            BackColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_MetCol 
            Height          =   315
            Left            =   1620
            TabIndex        =   22
            Top             =   60
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "2,000,000.00 "
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_TipMon 
            Height          =   315
            Left            =   3090
            TabIndex        =   24
            Top             =   60
            Width           =   3105
            _Version        =   65536
            _ExtentX        =   5477
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "(DOLARES AMERICANOS)"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel SSPanel11 
            Height          =   60
            Left            =   30
            TabIndex        =   25
            Top             =   420
            Width           =   8025
            _Version        =   65536
            _ExtentX        =   14155
            _ExtentY        =   106
            _StockProps     =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelOuter      =   0
            BevelInner      =   1
         End
         Begin Threed.SSPanel pnl_TotImp 
            Height          =   285
            Left            =   5910
            TabIndex        =   27
            Top             =   3570
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Alignment       =   4
         End
         Begin VB.Label Label5 
            Caption         =   "Totales ==>"
            Height          =   315
            Left            =   3330
            TabIndex        =   28
            Top             =   3570
            Width           =   1005
         End
         Begin VB.Label Label4 
            Caption         =   "Meta Colocación:"
            Height          =   315
            Left            =   60
            TabIndex        =   23
            Top             =   60
            Width           =   1305
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   60
         TabIndex        =   33
         Top             =   60
         Width           =   8055
         _Version        =   65536
         _ExtentX        =   14208
         _ExtentY        =   1191
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
         Begin Threed.SSPanel SSPanel12 
            Height          =   480
            Left            =   630
            TabIndex        =   34
            Top             =   90
            Width           =   6525
            _Version        =   65536
            _ExtentX        =   11509
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Distribución de Cartera por Sectores Ecónomicos"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
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
            Picture         =   "PltPar_frm_026.frx":1ACC
            Top             =   90
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_Carter_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Carter()   As moddat_tpo_Genera

Private Sub cmb_Carter_Click()
   Call gs_SetFocus(cmd_Buscar)
End Sub

Private Sub cmb_Carter_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Carter_Click
   End If
End Sub

Private Sub cmb_SecEco_Click()
   Call gs_SetFocus(ipp_Porcen)
End Sub

Private Sub cmb_SecEco_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_SecEco_Click
   End If
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   
   Call fs_Activa_Editar(True)
   Call gs_SetFocus(cmb_SecEco)
End Sub

Private Sub cmd_Borrar_Click()
   grd_Listad.Col = 3
   moddat_g_str_Codigo = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)

   If MsgBox("¿Está seguro de eliminar la Empresa de Seguro?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   'Obteniendo Información del Registro
   g_str_Parame = "USP_BORRAR_MNT_DCASEC (" & "'" & moddat_g_str_CodPrd & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_Codigo & "') "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
       Exit Sub
   End If
   
   Call fs_Buscar
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_Carter.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Cartera de Colocación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Carter)
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   Call fs_Activa(False)
   Call fs_Buscar
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Cancel_Click()
   cmb_SecEco.ListIndex = -1
   ipp_Porcen.Value = 0
   pnl_MetSec.Caption = "0.00 "
   
   Call fs_Activa_Editar(False)
   Call gs_SetFocus(grd_Listad)
   
   If grd_Listad.Rows = 0 Then
      cmd_Agrega.Enabled = True
      cmd_Editar.Enabled = False
      cmd_Borrar.Enabled = False
      grd_Listad.Enabled = False
   End If
End Sub

Private Sub cmd_Editar_Click()
   grd_Listad.Col = 3
   moddat_g_str_Codigo = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgGrb = 2
   
   
   'Obteniendo Información del Registro
   g_str_Parame = ""
   
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_DCASEC WHERE "
   g_str_Parame = g_str_Parame & "DCASEC_CODCAR = '" & moddat_g_str_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "DCASEC_CODSEC = '" & moddat_g_str_Codigo & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   Call gs_BuscarCombo_Item(cmb_SecEco, CInt(g_rst_Princi!DCASEC_CODSEC))
   ipp_Porcen.Value = g_rst_Princi!DCASEC_PORCOL
   
   Call ipp_Porcen_Change
      
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call fs_Activa_Editar(True)
   
   cmb_SecEco.Enabled = False
   Call gs_SetFocus(ipp_Porcen)
End Sub

Private Sub cmd_Grabar_Click()
   If cmb_SecEco.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Sector Económico.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_SecEco)
      Exit Sub
   End If
   
   If moddat_g_int_FlgGrb = 1 Then
      'Validar que el registro no exista
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM MNT_DCASEC WHERE "
      g_str_Parame = g_str_Parame & "DCASEC_CODCAR = '" & moddat_g_str_CodPrd & "' AND "
      g_str_Parame = g_str_Parame & "DCASEC_CODSEC = '" & Format(cmb_SecEco.ItemData(cmb_SecEco.ListIndex), "00") & "' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
        
         MsgBox "El Sector ya ha sido registrado..", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      
      g_str_Parame = "USP_MNT_DCASEC ("
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodPrd & "', "
      g_str_Parame = g_str_Parame & "'" & Format(cmb_SecEco.ItemData(cmb_SecEco.ListIndex), "00") & "', "
      g_str_Parame = g_str_Parame & CStr(ipp_Porcen.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_MetSec.Caption)) & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ") "
            
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      Screen.MousePointer = 0
   Loop
   
   Screen.MousePointer = 11
   
   Call fs_Buscar
   Call cmd_Cancel_Click
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Limpia_Click()
   cmb_Carter.ListIndex = -1
   pnl_MetCol.Caption = "0.00 "
   pnl_TipMon.Caption = "0.00 "
   
   cmb_SecEco.ListIndex = -1
   ipp_Porcen.Value = 0
   pnl_MetSec.Caption = "0.00 "
   
   Call gs_LimpiaGrid(grd_Listad)
   
   Call fs_Activa_Editar(False)
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_SecEco)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt & " - Distribución de Cartera por Sectores Económicos"
   
   Call fs_Inicia
   
   Call cmd_Limpia_Click
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 4290
   grd_Listad.ColWidth(1) = 1530
   grd_Listad.ColWidth(2) = 1770
   grd_Listad.ColWidth(3) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignRightCenter
   grd_Listad.ColAlignment(2) = flexAlignRightCenter
   
   Call moddat_gs_Carga_Carter(cmb_Carter, l_arr_Carter())
   Call moddat_gs_Carga_LisIte_Combo(cmb_SecEco, 1, "103")
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_Carter.Enabled = p_Activa
   cmd_Buscar.Enabled = p_Activa
   
   grd_Listad.Enabled = Not p_Activa
   cmd_Agrega.Enabled = Not p_Activa
   cmd_Editar.Enabled = Not p_Activa
   cmd_Borrar.Enabled = Not p_Activa
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_Editar_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub fs_Activa_Editar(ByVal p_Activa As Integer)
   cmd_Grabar.Enabled = p_Activa
   cmd_Cancel.Enabled = p_Activa
   cmb_SecEco.Enabled = p_Activa
   ipp_Porcen.Enabled = p_Activa
   
   grd_Listad.Enabled = Not p_Activa
   cmd_Agrega.Enabled = Not p_Activa
   cmd_Editar.Enabled = Not p_Activa
   cmd_Borrar.Enabled = Not p_Activa
End Sub

Private Sub fs_Buscar()
   Dim r_dbl_TotPor  As Double
   Dim r_dbl_TotImp  As Double
   
   cmd_Agrega.Enabled = True
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   grd_Listad.Enabled = False
   
   moddat_g_str_CodPrd = l_arr_Carter(cmb_Carter.ListIndex + 1).Genera_Codigo
   pnl_MetCol.Caption = Format(l_arr_Carter(cmb_Carter.ListIndex + 1).Genera_Cantid, "###,###,###,###,##0.00") & " "
   pnl_TipMon.Caption = moddat_gf_Consulta_ParDes("204", l_arr_Carter(cmb_Carter.ListIndex + 1).Genera_TipPar)
   
   Call gs_LimpiaGrid(grd_Listad)
   
   pnl_TotPor.Caption = "0.00 "
   pnl_TotImp.Caption = "0.00 "
   
   r_dbl_TotPor = 0
   r_dbl_TotImp = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_DCASEC WHERE "
   g_str_Parame = g_str_Parame & "DCASEC_CODCAR = '" & moddat_g_str_CodPrd & "' "
   g_str_Parame = g_str_Parame & "ORDER BY DCASEC_CODSEC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
     Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = moddat_gf_Consulta_ParDes("103", g_rst_Princi!DCASEC_CODSEC)
      
      grd_Listad.Col = 1
      grd_Listad.Text = Format(g_rst_Princi!DCASEC_PORCOL, "##0.00")
      
      grd_Listad.Col = 2
      grd_Listad.Text = Format(g_rst_Princi!DCASEC_IMPORT, "###,###,###,##0.00")
      
      grd_Listad.Col = 3
      grd_Listad.Text = g_rst_Princi!DCASEC_CODSEC
      
      r_dbl_TotPor = r_dbl_TotPor + g_rst_Princi!DCASEC_PORCOL
      r_dbl_TotImp = r_dbl_TotImp + g_rst_Princi!DCASEC_IMPORT
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If grd_Listad.Rows > 0 Then
      cmd_Editar.Enabled = True
      cmd_Borrar.Enabled = True
      grd_Listad.Enabled = True
   End If
   
   pnl_TotPor.Caption = Format(r_dbl_TotPor, "##0.00") & " "
   pnl_TotImp.Caption = Format(r_dbl_TotImp, "###,###,###,##0.00") & " "
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub ipp_Porcen_Change()
   pnl_MetSec.Caption = Format(CDbl(pnl_MetCol.Caption) * CDbl(ipp_Porcen.Text) / 100, "###,###,###,##0.00") & " "
End Sub

Private Sub ipp_Porcen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub
