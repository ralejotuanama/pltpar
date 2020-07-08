VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Comviv_2 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4020
   ClientLeft      =   2385
   ClientTop       =   5280
   ClientWidth     =   11175
   Icon            =   "PltPar_frm_040.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel3 
      Height          =   3945
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      _Version        =   65536
      _ExtentX        =   19711
      _ExtentY        =   6959
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
      Begin Threed.SSPanel SSPanel11 
         Height          =   1905
         Left            =   30
         TabIndex        =   8
         Top             =   2100
         Width           =   11085
         _Version        =   65536
         _ExtentX        =   19553
         _ExtentY        =   3360
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
         Begin VB.ComboBox cmb_TipCom 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   60
            Width           =   9015
         End
         Begin VB.ComboBox cmb_TipMon 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   420
            Width           =   9015
         End
         Begin EditLib.fpDoubleSingle ipp_PorCom 
            Height          =   345
            Left            =   2040
            TabIndex        =   5
            Top             =   1500
            Width           =   855
            _Version        =   196608
            _ExtentX        =   1508
            _ExtentY        =   609
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
            ThreeDInsideHighlightColor=   -2147483633
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
            ButtonMin       =   1
            ButtonMax       =   100
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
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
            Text            =   "0.0000"
            DecimalPlaces   =   4
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
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpLongInteger fpl_PlaIni 
            Height          =   315
            Left            =   2040
            TabIndex        =   3
            Top             =   780
            Width           =   615
            _Version        =   196608
            _ExtentX        =   1085
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
            ThreeDInsideHighlightColor=   -2147483633
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
            ButtonMax       =   30
            ButtonStyle     =   1
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
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
            MaxValue        =   "30"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpLongInteger fpl_PlaFin 
            Height          =   315
            Left            =   2040
            TabIndex        =   4
            Top             =   1140
            Width           =   615
            _Version        =   196608
            _ExtentX        =   1085
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
            ThreeDInsideHighlightColor=   -2147483633
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
            ButtonMax       =   30
            ButtonStyle     =   1
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
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
            MaxValue        =   "30"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   0   'False
            IncInt          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label6 
            Caption         =   "Porcentaje de Comisión:"
            Height          =   315
            Left            =   60
            TabIndex        =   13
            Top             =   1560
            Width           =   1785
         End
         Begin VB.Label Label5 
            Caption         =   "Plazo Final:"
            Height          =   315
            Left            =   60
            TabIndex        =   12
            Top             =   1200
            Width           =   1545
         End
         Begin VB.Label Label3 
            Caption         =   "Plazo Inicial:"
            Height          =   315
            Left            =   60
            TabIndex        =   11
            Top             =   840
            Width           =   1545
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo de Moneda:"
            Height          =   315
            Left            =   60
            TabIndex        =   10
            Top             =   480
            Width           =   1545
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo de Comisión:"
            Height          =   345
            Left            =   60
            TabIndex        =   9
            Top             =   120
            Width           =   1425
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   495
         Left            =   30
         TabIndex        =   14
         Top             =   1560
         Width           =   11085
         _Version        =   65536
         _ExtentX        =   19553
         _ExtentY        =   873
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
         Begin Threed.SSPanel pnl_CodPro 
            Height          =   315
            Left            =   2040
            TabIndex        =   15
            Top             =   90
            Width           =   8985
            _Version        =   65536
            _ExtentX        =   15849
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Label Label1 
            Caption         =   "Producto :"
            Height          =   315
            Left            =   60
            TabIndex        =   16
            Top             =   150
            Width           =   1515
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   735
         Left            =   30
         TabIndex        =   17
         Top             =   780
         Width           =   11085
         _Version        =   65536
         _ExtentX        =   19553
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10350
            Picture         =   "PltPar_frm_040.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   10
            Picture         =   "PltPar_frm_040.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   18
         Top             =   60
         Width           =   11085
         _Version        =   65536
         _ExtentX        =   19553
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
         Begin Threed.SSPanel pnl_TitPri 
            Height          =   315
            Left            =   660
            TabIndex        =   19
            Top             =   60
            Width           =   6615
            _Version        =   65536
            _ExtentX        =   11668
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Mantenimiento Comisiones Mivivienda"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
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
         Begin Threed.SSPanel SSPanel1 
            Height          =   315
            Left            =   660
            TabIndex        =   20
            Top             =   330
            Width           =   3975
            _Version        =   65536
            _ExtentX        =   7011
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "SSPanel1"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
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
            Picture         =   "PltPar_frm_040.frx":0890
            Top             =   90
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_Comviv_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Produc()      As moddat_tpo_Genera

Private Sub cmb_TipCom_Click()
   'Se pasa al siguiente control
   Call gs_SetFocus(cmb_TipMon)
End Sub

Private Sub cmb_TipMon_Click()
   'Se pasa al siguiente control
   Call gs_SetFocus(fpl_PlaIni)
End Sub


Private Sub Form_Load()
   Screen.MousePointer = vbHourglass
   'Se llama a los metodos fs_inica y gs_focus
   Call fs_Inicia
   Call gs_SetFocus(cmb_TipCom)
    
   'Se muestra el nombre de la plataforma y el nombre del producto
   Me.Caption = modgen_g_str_NomPlt
   pnl_CodPro.Caption = moddat_g_str_NomPrd
      
   'Condicion si el tipo de panel es 1 = nuevo ingreso
   If modvar_g_int_TipPan = 1 Then
      SSPanel1.Caption = "Nuevo Ingreso"
   End If
    
   'Condicion si el tipo de panel es 2 = modificar
   If modvar_g_int_TipPan = 2 Then
      SSPanel1.Caption = "Modificar Datos"
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM OPE_COMMVI WHERE "
      g_str_Parame = g_str_Parame & "COMMVI_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
      g_str_Parame = g_str_Parame & "COMMVI_TIPCOM = '" & modvar_g_int_TipCom & "' AND "
      g_str_Parame = g_str_Parame & "COMMVI_TIPMON = '" & modvar_g_int_TipMon & "' AND "
      g_str_Parame = g_str_Parame & "COMMVI_PLAINI = '" & modvar_g_int_PlaIni & "' "
         
      'Condicion si NO se ejecuto le sentencia SQL
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      g_rst_Princi.MoveFirst
      
      'Hacemos llamado a los combos
      Call gs_BuscarCombo_Item(cmb_TipCom, g_rst_Princi!COMMVI_TIPCOM)
      Call gs_BuscarCombo_Item(cmb_TipMon, g_rst_Princi!COMMVI_TIPMON)
      
      'Hacemos llamado y mostramos los campos correspondientes
      fpl_PlaIni.Value = g_rst_Princi!COMMVI_PLAINI
      fpl_PlaFin.Value = g_rst_Princi!COMMVI_PLAFIN
      ipp_PorCom.Value = g_rst_Princi!COMMVI_PORCEN
      
      'Cerramos la conexion
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      'Desabilitamos los controles para q la data sea de solo lectura
      cmb_TipCom.Enabled = False
      cmb_TipMon.Enabled = False
      fpl_PlaIni.Enabled = False
      fpl_PlaFin.Enabled = False
   End If
    
   'Centramos el Formulario
   Call gs_CentraForm(Me)
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Grabar_Click()
   'Validacion para escoger el Tipo de Comision
   If cmb_TipCom.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Comision.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipCom)
      Exit Sub
   End If
   'Validacion para escoger el tipo de moneda
   If cmb_TipMon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Moneda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipMon)
      Exit Sub
   End If
   'Validacion del plazo inicial
   If fpl_PlaIni.Value < 1 Then
      MsgBox "El Plazo Inicial debe ser mayor a cero.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(fpl_PlaIni)
      Exit Sub
   End If
   'Validacion del plazo final
   If fpl_PlaFin.Value < 1 Then
      MsgBox "El Plazo Final debe ser mayor a cero.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(fpl_PlaFin)
      Exit Sub
   End If
   'Validacion si el plazo inicial es menor al plazo final
   If CInt(fpl_PlaIni.Text) > CInt(fpl_PlaFin.Text) Then
      MsgBox "El plazo inicial tiene que ser menor a el plazo final.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   'Validacion para ingresar el porcentaje de comision
   If ipp_PorCom.Value < 0.01 Then
      MsgBox "Debe ingresar el Porcentaje de Comisión.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PorCom)
      Exit Sub
   End If
                                                            
   If modvar_g_int_TipPan = 1 Then
      'Obteniendo Información del Registro y ingresando la condion para que no se registre plazos ya existente en la BD
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM OPE_COMMVI WHERE "
      g_str_Parame = g_str_Parame & "COMMVI_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
      g_str_Parame = g_str_Parame & "COMMVI_TIPCOM = " & CStr(cmb_TipCom.ItemData(cmb_TipCom.ListIndex)) & " AND "
      g_str_Parame = g_str_Parame & "COMMVI_TIPMON = " & CStr(cmb_TipMon.ItemData(cmb_TipMon.ListIndex)) & " AND "
      g_str_Parame = g_str_Parame & "( (COMMVI_PLAINI <= " & fpl_PlaIni.Text & " AND COMMVI_PLAFIN >= " & fpl_PlaIni.Text & " AND COMMVI_PLAINI <= " & fpl_PlaFin.Text & " AND COMMVI_PLAFIN >= " & fpl_PlaFin.Text & ") OR  "
      g_str_Parame = g_str_Parame & "(COMMVI_PLAINI >= " & fpl_PlaIni.Text & " AND COMMVI_PLAINI <= " & fpl_PlaFin.Text & " AND COMMVI_PLAFIN >= " & fpl_PlaIni.Text & " AND COMMVI_PLAFIN <= " & fpl_PlaFin.Text & ") OR  "
      g_str_Parame = g_str_Parame & "(COMMVI_PLAINI >= " & fpl_PlaIni.Text & " AND COMMVI_PLAINI <= " & fpl_PlaFin.Text & " AND COMMVI_PLAFIN >= " & fpl_PlaIni.Text & " AND COMMVI_PLAFIN >= " & fpl_PlaFin.Text & ") OR  "
      g_str_Parame = g_str_Parame & "(COMMVI_PLAINI <= " & fpl_PlaIni.Text & " AND COMMVI_PLAINI <= " & fpl_PlaFin.Text & " AND COMMVI_PLAFIN >= " & fpl_PlaIni.Text & " AND COMMVI_PLAFIN <= " & fpl_PlaFin.Text & ") ) "
      g_str_Parame = g_str_Parame & "ORDER BY COMMVI_PLAINI ASC "
      
      'Condicion si No se ejecuta la sentencia SQL
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      'Condicion si No se encuentra al comienzo o al final del archivo y lo evalua
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         'Cerramos la conexion a la BD
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         
         'Mensaje mostranto que ya se encuentra el rango en uso
         MsgBox "El Rango que desea Ingresar ya se encuentra en el sistema.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      'Cerramos la conexion a la BD
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   'Mensaje de confirmacion de grabado de los datos correspondientes
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Reloj de arena
   Screen.MousePointer = vbHourglass
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
                     
   'Mientras la variable sea falsa se procede a ejecutar el procedure
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_MNT_COMMVI ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodPrd & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_TipCom.ItemData(cmb_TipCom.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_TipMon.ItemData(cmb_TipMon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(fpl_PlaIni.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(fpl_PlaFin.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_PorCom.Value) & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & CStr(modvar_g_int_TipPan) & ") "
                              
      'Condicion si No se ejecuto la sentencia SQL, se aumenta en uno la variable con el error
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
                                 
      'Si pregunta si la variable de error llega a 6, de darse el caso se muestra un mensaje si se desea seguir intentando
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
                     
   'Mouse normal
   Screen.MousePointer = vbDefault
   
   'Se envia mensaje mostrando el grabado de los datos
   MsgBox "Se grabaron los datos correctamente.", vbInformation, modgen_g_str_NomPlt
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   'Se cierra el formulario
   Unload Me
End Sub

Private Sub fs_Inicia()
   'Hacemos la llamada a los combos
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipCom, 1, "029")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipMon, 1, "204")
End Sub

Private Sub fpl_PlaFin_KeyPress(KeyAscii As Integer)
   'Se pasa al siguiente control
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PorCom)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & ",-_.( )$/")
   End If
End Sub

Private Sub fpl_PlaIni_KeyPress(KeyAscii As Integer)
   'Se pasa al siguiente control
   If KeyAscii = 13 Then
      Call gs_SetFocus(fpl_PlaFin)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & ",-_.( )$/")
   End If
End Sub

Private Sub ipp_PorCom_KeyPress(KeyAscii As Integer)
   'Se pasa al siguiente control
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & ",-_.( )$/")
   End If
End Sub
