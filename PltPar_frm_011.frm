VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Produc_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Productos de Crédito"
   ClientHeight    =   8655
   ClientLeft      =   1890
   ClientTop       =   1545
   ClientWidth     =   8745
   Icon            =   "PltPar_frm_011.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8655
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   9015
      _Version        =   65536
      _ExtentX        =   15901
      _ExtentY        =   15266
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   735
         Left            =   60
         TabIndex        =   28
         Top             =   3960
         Width           =   8610
         _Version        =   65536
         _ExtentX        =   15187
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
         Begin VB.ComboBox cmb_VerSit 
            Height          =   315
            Left            =   750
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   240
            Width           =   1590
         End
         Begin VB.CommandButton cmd_SubPrd 
            Height          =   675
            Left            =   7230
            Picture         =   "PltPar_frm_011.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Sub-Productos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   675
            Left            =   6540
            Picture         =   "PltPar_frm_011.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   7920
            Picture         =   "PltPar_frm_011.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   675
            Left            =   5850
            Picture         =   "PltPar_frm_011.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   675
            Left            =   5160
            Picture         =   "PltPar_frm_011.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Mostrar:"
            Height          =   195
            Left            =   90
            TabIndex        =   37
            Top             =   300
            Width           =   570
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   735
         Left            =   60
         TabIndex        =   27
         Top             =   7860
         Width           =   8610
         _Version        =   65536
         _ExtentX        =   15187
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
            Left            =   7230
            Picture         =   "PltPar_frm_011.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Cancel 
            Height          =   675
            Left            =   7920
            Picture         =   "PltPar_frm_011.frx":15F0
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   3135
         Left            =   60
         TabIndex        =   18
         Top             =   780
         Width           =   8610
         _Version        =   65536
         _ExtentX        =   15187
         _ExtentY        =   5530
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
            Height          =   2745
            Left            =   60
            TabIndex        =   0
            Top             =   360
            Width           =   8490
            _ExtentX        =   14975
            _ExtentY        =   4842
            _Version        =   393216
            Rows            =   12
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_DesProd 
            Height          =   285
            Left            =   945
            TabIndex        =   19
            Top             =   60
            Width           =   6135
            _Version        =   65536
            _ExtentX        =   10821
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Descripción Producto"
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
         Begin Threed.SSPanel pnl_Codigo 
            Height          =   285
            Left            =   90
            TabIndex        =   20
            Top             =   60
            Width           =   870
            _Version        =   65536
            _ExtentX        =   1535
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Código"
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
         Begin Threed.SSPanel pnl_Situac 
            Height          =   285
            Left            =   7050
            TabIndex        =   35
            Top             =   60
            Width           =   1140
            _Version        =   65536
            _ExtentX        =   2011
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Situación"
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
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   3075
         Left            =   60
         TabIndex        =   21
         Top             =   4740
         Width           =   8610
         _Version        =   65536
         _ExtentX        =   15187
         _ExtentY        =   5424
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
         Begin VB.ComboBox cmb_SitCom 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1710
            Width           =   3225
         End
         Begin VB.ComboBox cmb_IndITF 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   2040
            Width           =   975
         End
         Begin VB.ComboBox cmb_Situac 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1380
            Width           =   3225
         End
         Begin VB.ComboBox cmb_TipCre 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1050
            Width           =   3225
         End
         Begin VB.ComboBox cmb_ClaCre 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   720
            Width           =   3225
         End
         Begin VB.TextBox txt_CodPrd 
            Height          =   315
            Left            =   2010
            MaxLength       =   3
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   60
            Width           =   975
         End
         Begin VB.TextBox txt_Descri 
            Height          =   315
            Left            =   2010
            MaxLength       =   80
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   390
            Width           =   6015
         End
         Begin EditLib.fpLongInteger ipp_VctCuo 
            Height          =   315
            Left            =   2010
            TabIndex        =   13
            Top             =   2370
            Width           =   975
            _Version        =   196608
            _ExtentX        =   1720
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
            ButtonStyle     =   1
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
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
            Text            =   "0"
            MaxValue        =   "9999"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpLongInteger ipp_VctCre 
            Height          =   315
            Left            =   2010
            TabIndex        =   14
            Top             =   2700
            Width           =   975
            _Version        =   196608
            _ExtentX        =   1720
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
            ButtonStyle     =   1
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
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
            Text            =   "0"
            MaxValue        =   "9999"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label9 
            Caption         =   "Situación Comercial:"
            Height          =   285
            Left            =   90
            TabIndex        =   34
            Top             =   1740
            Width           =   1605
         End
         Begin VB.Label Label8 
            Caption         =   "Vcto de Crédito:"
            Height          =   285
            Left            =   90
            TabIndex        =   31
            Top             =   2730
            Width           =   1605
         End
         Begin VB.Label Label7 
            Caption         =   "Vcto de Cuota:"
            Height          =   285
            Left            =   90
            TabIndex        =   30
            Top             =   2400
            Width           =   1605
         End
         Begin VB.Label Label6 
            Caption         =   "Indicador ITF:"
            Height          =   285
            Left            =   90
            TabIndex        =   29
            Top             =   2070
            Width           =   1605
         End
         Begin VB.Label Label5 
            Caption         =   "Estado:"
            Height          =   285
            Left            =   90
            TabIndex        =   26
            Top             =   1410
            Width           =   1605
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo de Crédito:"
            Height          =   285
            Left            =   90
            TabIndex        =   25
            Top             =   1080
            Width           =   1605
         End
         Begin VB.Label Label1 
            Caption         =   "Clase de Crédito:"
            Height          =   285
            Left            =   90
            TabIndex        =   24
            Top             =   750
            Width           =   1605
         End
         Begin VB.Label Label3 
            Caption         =   "Código Producto:"
            Height          =   285
            Left            =   90
            TabIndex        =   23
            Top             =   90
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Descripción Producto:"
            Height          =   285
            Left            =   90
            TabIndex        =   22
            Top             =   420
            Width           =   1815
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   32
         Top             =   60
         Width           =   8610
         _Version        =   65536
         _ExtentX        =   15187
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   480
            Left            =   630
            TabIndex        =   33
            Top             =   90
            Width           =   7185
            _Version        =   65536
            _ExtentX        =   12674
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Productos de Crédito Hipotecario"
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
            Picture         =   "PltPar_frm_011.frx":18FA
            Top             =   90
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_Produc_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_VerSit_Click()
   Call fs_Buscar
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   Call fs_Activa(False)
   Call gs_SetFocus(txt_CodPrd)
End Sub

Private Sub cmd_Borrar_Click()
   grd_Listad.Col = 0
   moddat_g_str_Codigo = grd_Listad.Text
   Call gs_RefrescaGrid(grd_Listad)

   If MsgBox("¿Está seguro de eliminar el Producto?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   'Obteniendo Información del Registro
   g_str_Parame = "USP_CRE_PRODUC_BORRAR (" & "'" & moddat_g_str_Codigo & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
   
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Cancel_Click()
Dim r_int_aux As Integer

   Call fs_Activa(True)
   r_int_aux = cmb_VerSit.ListIndex
   Call fs_Limpia
   cmb_VerSit.ListIndex = r_int_aux
   
   Call gs_SetFocus(grd_Listad)

   If grd_Listad.Rows = 0 Then
      cmd_Agrega.Enabled = True
      cmd_Editar.Enabled = False
      cmd_Borrar.Enabled = False
      cmd_SubPrd.Enabled = False
      grd_Listad.Enabled = False
   End If
End Sub

Private Sub cmd_Editar_Click()
   grd_Listad.Col = 0
   moddat_g_str_Codigo = grd_Listad.Text
   Call gs_RefrescaGrid(grd_Listad)
   moddat_g_int_FlgGrb = 2
   
   'Obteniendo Información del Registro
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PRODUC WHERE "
   g_str_Parame = g_str_Parame & "PRODUC_CODIGO = '" & moddat_g_str_Codigo & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   txt_CodPrd.Text = Trim(g_rst_Genera!PRODUC_CODIGO)
   txt_Descri.Text = Trim(g_rst_Genera!PRODUC_DESCRI)
   Call gs_BuscarCombo_Text(cmb_ClaCre, g_rst_Genera!PRODUC_CODCLA, 1)
   Call gs_BuscarCombo_Text(cmb_TipCre, g_rst_Genera!PRODUC_TIPCRE, 1)
   Call gs_BuscarCombo_Item(cmb_Situac, g_rst_Genera!PRODUC_SITUAC)
   Call gs_BuscarCombo_Item(cmb_SitCom, g_rst_Genera!PRODUC_SITCOM)
   Call gs_BuscarCombo_Item(cmb_IndITF, g_rst_Genera!PRODUC_INDITF)
   ipp_VctCuo.Value = g_rst_Genera!PRODUC_VCTCUO
   ipp_VctCre.Value = g_rst_Genera!PRODUC_VCTCRE
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   Call fs_Activa(False)
   txt_CodPrd.Enabled = False
   Call gs_SetFocus(txt_Descri)
End Sub

Private Sub cmd_Grabar_Click()
   txt_CodPrd.Text = Format(txt_CodPrd.Text, "000")
   
   If Len(Trim(txt_CodPrd.Text)) = 0 Then
      MsgBox "Debe ingresar el Código del Producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_CodPrd)
      Exit Sub
   End If
   If Len(Trim(txt_Descri.Text)) = 0 Then
      MsgBox "Debe ingresar la Descripción del Producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Descri)
      Exit Sub
   End If
   If cmb_ClaCre.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Clase de Crédito.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_ClaCre)
      Exit Sub
   End If
   If cmb_TipCre.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Crédito.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipCre)
      Exit Sub
   End If
   If cmb_Situac.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Estado del Producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Situac)
      Exit Sub
   End If
   If cmb_SitCom.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Situación Comercial del Producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_SitCom)
      Exit Sub
   End If
   If cmb_IndITF.ListIndex = -1 Then
      MsgBox "Debe seleccionar si se aplica ITF sobre el Producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_IndITF)
      Exit Sub
   End If
   If moddat_g_int_FlgGrb = 1 Then
      'Validar que el registro no exista
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PRODUC WHERE "
      g_str_Parame = g_str_Parame & "PRODUC_CODIGO = '" & txt_CodPrd.Text & "' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         MsgBox "El Código ya ha sido registrado. Por favor verifique el código e intente nuevamente.", vbExclamation, modgen_g_str_NomPlt
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
      
      g_str_Parame = "USP_CRE_PRODUC ("
      g_str_Parame = g_str_Parame & "'" & txt_CodPrd.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Descri.Text & "', "
      g_str_Parame = g_str_Parame & Left(cmb_ClaCre.Text, 1) & ", "
      g_str_Parame = g_str_Parame & Left(cmb_TipCre.Text, 1) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_Situac.ItemData(cmb_Situac.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_SitCom.ItemData(cmb_SitCom.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_IndITF.ItemData(cmb_IndITF.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_VctCuo.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_VctCre.Value) & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
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

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_SubPrd_Click()
   grd_Listad.Col = 0
   moddat_g_str_CodPrd = grd_Listad.Text
         
   grd_Listad.Col = 1
   moddat_g_str_NomPrd = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)

   frm_Produc_02.Show 1
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 0
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Activa(True)
   Call fs_Limpia
   cmb_VerSit.ListIndex = 0
   'Call fs_Buscar
   
   Call gs_CentraForm(Me)
   Call gs_SetFocus(grd_Listad)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 860
   grd_Listad.ColWidth(1) = 6130
   grd_Listad.ColWidth(2) = 1100
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_ClaCre, 1, "055")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipCre, 1, "056")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Situac, 1, "013")
   Call moddat_gs_Carga_LisIte_Combo(cmb_SitCom, 1, "013")
   Call moddat_gs_Carga_LisIte_Combo(cmb_VerSit, 1, "013")
   cmb_VerSit.AddItem Trim$("<<TODOS>>")
   cmb_VerSit.ItemData(cmb_VerSit.NewIndex) = 0
   Call moddat_gs_Carga_LisIte_Combo(cmb_IndITF, 1, "214")
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   grd_Listad.Enabled = p_Activa
   cmd_Agrega.Enabled = p_Activa
   cmd_Editar.Enabled = p_Activa
   cmd_Borrar.Enabled = p_Activa
   cmd_SubPrd.Enabled = p_Activa
   cmb_VerSit.Enabled = p_Activa
   
   txt_CodPrd.Enabled = Not p_Activa
   txt_Descri.Enabled = Not p_Activa
   cmb_ClaCre.Enabled = Not p_Activa
   cmb_TipCre.Enabled = Not p_Activa
   cmb_Situac.Enabled = Not p_Activa
   cmb_SitCom.Enabled = Not p_Activa
   cmb_IndITF.Enabled = Not p_Activa
   ipp_VctCuo.Enabled = Not p_Activa
   ipp_VctCre.Enabled = Not p_Activa
   cmd_Grabar.Enabled = Not p_Activa
   cmd_Cancel.Enabled = Not p_Activa
End Sub

Private Sub fs_Limpia()
   txt_CodPrd.Text = ""
   txt_Descri.Text = ""
   cmb_ClaCre.ListIndex = -1
   cmb_TipCre.ListIndex = -1
   cmb_Situac.ListIndex = -1
   cmb_SitCom.ListIndex = -1
   cmb_IndITF.ListIndex = -1
   cmb_VerSit.ListIndex = -1
   ipp_VctCuo.Value = 0
   ipp_VctCre.Value = 0
End Sub

Private Sub fs_Buscar()
   If cmb_VerSit.ListIndex = -1 Then
      Exit Sub
   End If
   
   cmd_Agrega.Enabled = True
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   cmd_SubPrd.Enabled = False
   grd_Listad.Enabled = False
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   'g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PRODUC WHERE PRODUC_SITUAC <> 0 "
   'g_str_Parame = g_str_Parame & "ORDER BY PRODUC_CODIGO ASC "
   g_str_Parame = g_str_Parame & " SELECT PRODUC_CODIGO, TRIM(PRODUC_DESCRI) PRODUC_DESCRI , PRODUC_CODCLA, PRODUC_TIPCRE,  "
   g_str_Parame = g_str_Parame & "        PRODUC_SITUAC, TRIM(B.PARDES_DESCRI) AS SITUACION, PRODUC_INDITF, PRODUC_VCTCUO,  "
   g_str_Parame = g_str_Parame & "        PRODUC_VCTCRE , PRODUC_SITCOM  "
   g_str_Parame = g_str_Parame & "   FROM CRE_PRODUC A  "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = 013 AND B.PARDES_CODITE = A.PRODUC_SITCOM  "
   g_str_Parame = g_str_Parame & "  WHERE PRODUC_SITUAC <> 0  "
   If cmb_VerSit.ItemData(cmb_VerSit.ListIndex) <> 0 Then
      g_str_Parame = g_str_Parame & "    AND PRODUC_SITCOM = " & cmb_VerSit.ItemData(cmb_VerSit.ListIndex)
   End If
   g_str_Parame = g_str_Parame & "  ORDER BY PRODUC_CODIGO ASC  "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
     Exit Sub
   End If
   
   grd_Listad.Redraw = False
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = Trim(g_rst_Genera!PRODUC_CODIGO)
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Genera!PRODUC_DESCRI)
      
      grd_Listad.Col = 2
      grd_Listad.Text = Trim(g_rst_Genera!SITUACION)
      
      g_rst_Genera.MoveNext
   Loop
   grd_Listad.Redraw = True
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   If grd_Listad.Rows > 0 Then
      cmd_Editar.Enabled = True
      cmd_Borrar.Enabled = True
      cmd_SubPrd.Enabled = True
      grd_Listad.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub pnl_Codigo_Click()
   If pnl_Codigo.Tag = "" Then
      pnl_Codigo.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Codigo.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub

Private Sub pnl_DesProd_Click()
   If pnl_Codigo.Tag = "" Then
      pnl_Codigo.Tag = "2"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Codigo.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub

Private Sub pnl_Situac_Click()
   If pnl_Codigo.Tag = "" Then
      pnl_Codigo.Tag = "2"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_Codigo.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub

Private Sub txt_CodPrd_GotFocus()
   Call gs_SelecTodo(txt_CodPrd)
End Sub

Private Sub txt_CodPrd_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Descri)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Descri_GotFocus()
   Call gs_SelecTodo(txt_Descri)
End Sub

Private Sub txt_Descri_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_ClaCre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & ",-_.( )$/")
   End If
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_Editar_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub cmb_ClaCre_Click()
   Call gs_SetFocus(cmb_TipCre)
End Sub

Private Sub cmb_ClaCre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_ClaCre_Click
   End If
End Sub

Private Sub cmb_TipCre_Click()
   Call gs_SetFocus(cmb_Situac)
End Sub

Private Sub cmb_TipCre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipCre_Click
   End If
End Sub

Private Sub cmb_Situac_Click()
   Call gs_SetFocus(cmb_SitCom)
End Sub

Private Sub cmb_Situac_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Situac_Click
   End If
End Sub

Private Sub cmb_SitCom_Click()
   Call gs_SetFocus(cmb_IndITF)
End Sub

Private Sub cmb_SitCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_SitCom_Click
   End If
End Sub

Private Sub cmb_IndITF_Click()
   Call gs_SetFocus(ipp_VctCuo)
End Sub

Private Sub cmb_IndITF_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_IndITF_Click
   End If
End Sub

Private Sub ipp_VctCuo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_VctCre)
   End If
End Sub

Private Sub ipp_VctCre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub
