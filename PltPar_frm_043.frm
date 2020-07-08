VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Produc_10 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   9990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9855
   Icon            =   "PltPar_frm_043.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9990
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   9975
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   9855
      _Version        =   65536
      _ExtentX        =   17383
      _ExtentY        =   17595
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
         TabIndex        =   27
         Top             =   4260
         Width           =   9750
         _Version        =   65536
         _ExtentX        =   17198
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
            Left            =   9060
            Picture         =   "PltPar_frm_043.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   675
            Left            =   8400
            Picture         =   "PltPar_frm_043.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   675
            Left            =   7740
            Picture         =   "PltPar_frm_043.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   675
            Left            =   7080
            Picture         =   "PltPar_frm_043.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.ComboBox cmb_VerSit 
            Height          =   315
            Left            =   990
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   240
            Width           =   1590
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Mostrar:"
            Height          =   195
            Left            =   90
            TabIndex        =   28
            Top             =   300
            Width           =   570
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   765
         Left            =   60
         TabIndex        =   29
         Top             =   9150
         Width           =   9750
         _Version        =   65536
         _ExtentX        =   17198
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
         Begin VB.CommandButton cmd_Cancel 
            Height          =   675
            Left            =   9060
            Picture         =   "PltPar_frm_043.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   8400
            Picture         =   "PltPar_frm_043.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   3435
         Left            =   60
         TabIndex        =   30
         Top             =   780
         Width           =   9750
         _Version        =   65536
         _ExtentX        =   17198
         _ExtentY        =   6059
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
            Height          =   3015
            Left            =   60
            TabIndex        =   0
            Top             =   360
            Width           =   9630
            _ExtentX        =   16986
            _ExtentY        =   5318
            _Version        =   393216
            Rows            =   13
            Cols            =   5
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
            TabIndex        =   31
            Top             =   60
            Width           =   5025
            _Version        =   65536
            _ExtentX        =   8864
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Descripción"
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
            TabIndex        =   32
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
            Left            =   8160
            TabIndex        =   33
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
         Begin Threed.SSPanel pnl_TasInt 
            Height          =   285
            Left            =   5940
            TabIndex        =   53
            Top             =   60
            Width           =   1140
            _Version        =   65536
            _ExtentX        =   2011
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tasa Activa"
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
            Left            =   7050
            TabIndex        =   56
            Top             =   60
            Width           =   1140
            _Version        =   65536
            _ExtentX        =   2011
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tasa Pasiva"
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
         Height          =   4065
         Left            =   60
         TabIndex        =   34
         Top             =   5040
         Width           =   9750
         _Version        =   65536
         _ExtentX        =   17198
         _ExtentY        =   7170
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
         Begin VB.ComboBox cmb_TipBon 
            Height          =   315
            Left            =   2100
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1380
            Width           =   1995
         End
         Begin VB.TextBox txt_Descri 
            Height          =   315
            Left            =   2100
            MaxLength       =   80
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   60
            Width           =   7305
         End
         Begin VB.ComboBox cmb_TipPro 
            Height          =   315
            Left            =   2100
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   390
            Width           =   7305
         End
         Begin VB.ComboBox cmb_TipEva 
            Height          =   315
            Left            =   2100
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1050
            Width           =   7305
         End
         Begin VB.ComboBox cmb_Situac 
            Height          =   315
            Left            =   2100
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   3690
            Width           =   1995
         End
         Begin VB.ComboBox cmb_CodPry 
            Height          =   315
            Left            =   2100
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   720
            Width           =   7305
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   2100
            TabIndex        =   11
            Top             =   1710
            Width           =   1995
            _Version        =   196608
            _ExtentX        =   3519
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
            Text            =   "28/09/2004"
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
         Begin EditLib.fpDateTime ipp_FecFin 
            Height          =   345
            Left            =   7410
            TabIndex        =   12
            Top             =   1710
            Width           =   1995
            _Version        =   196608
            _ExtentX        =   3519
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
            Text            =   "28/09/2004"
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
         Begin EditLib.fpLongInteger ipp_PlaMin 
            Height          =   315
            Left            =   2100
            TabIndex        =   17
            Top             =   2700
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
            MaxValue        =   "300"
            MinValue        =   "60"
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
         Begin EditLib.fpLongInteger ipp_PlaMax 
            Height          =   315
            Left            =   7410
            TabIndex        =   18
            Top             =   2730
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
            MaxValue        =   "300"
            MinValue        =   "60"
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
         Begin EditLib.fpDoubleSingle ipp_TasIntAct 
            Height          =   315
            Left            =   2100
            TabIndex        =   21
            Top             =   3360
            Width           =   1995
            _Version        =   196608
            _ExtentX        =   3519
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
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
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
         Begin EditLib.fpDoubleSingle ipp_PorIniMin 
            Height          =   315
            Left            =   2100
            TabIndex        =   19
            Top             =   3030
            Width           =   1995
            _Version        =   196608
            _ExtentX        =   3519
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
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
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
         Begin EditLib.fpDoubleSingle ipp_PorIniMax 
            Height          =   315
            Left            =   7410
            TabIndex        =   20
            Top             =   3060
            Width           =   1995
            _Version        =   196608
            _ExtentX        =   3519
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
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
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
         Begin EditLib.fpDoubleSingle ipp_ValPreMin 
            Height          =   315
            Left            =   2100
            TabIndex        =   13
            Top             =   2040
            Width           =   1995
            _Version        =   196608
            _ExtentX        =   3519
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
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
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
         Begin EditLib.fpDoubleSingle ipp_ValPreMax 
            Height          =   315
            Left            =   7410
            TabIndex        =   14
            Top             =   2070
            Width           =   1995
            _Version        =   196608
            _ExtentX        =   3519
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
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
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
         Begin EditLib.fpDoubleSingle ipp_ValInmMin 
            Height          =   315
            Left            =   2100
            TabIndex        =   15
            Top             =   2370
            Width           =   1995
            _Version        =   196608
            _ExtentX        =   3519
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
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
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
         Begin EditLib.fpDoubleSingle ipp_ValInmMax 
            Height          =   315
            Left            =   7410
            TabIndex        =   16
            Top             =   2400
            Width           =   1995
            _Version        =   196608
            _ExtentX        =   3519
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
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
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
         Begin EditLib.fpDoubleSingle ipp_TasIntPas 
            Height          =   315
            Left            =   7410
            TabIndex        =   22
            Top             =   3390
            Width           =   1995
            _Version        =   196608
            _ExtentX        =   3519
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
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
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
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Bono:"
            Height          =   195
            Left            =   150
            TabIndex        =   55
            Top             =   1440
            Width           =   420
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Tasa Pasiva (%):"
            Height          =   195
            Left            =   5430
            TabIndex        =   54
            Top             =   3450
            Width           =   1185
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Porc. Inicial (% Máx):"
            Height          =   195
            Left            =   5430
            TabIndex        =   52
            Top             =   3120
            Width           =   1470
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Porc. Inicial (% Mín):"
            Height          =   195
            Left            =   150
            TabIndex        =   51
            Top             =   3090
            Width           =   1455
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Tasa Activa (%):"
            Height          =   195
            Left            =   150
            TabIndex        =   50
            Top             =   3420
            Width           =   1155
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Plazo Máximo (Meses):"
            Height          =   195
            Left            =   5430
            TabIndex        =   49
            Top             =   2790
            Width           =   1620
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Plazo Mínimo (Meses):"
            Height          =   195
            Left            =   150
            TabIndex        =   48
            Top             =   2760
            Width           =   1605
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Valor Inmueble (Máx):"
            Height          =   195
            Left            =   5430
            TabIndex        =   47
            Top             =   2460
            Width           =   1530
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Valor Inmueble (Min):"
            Height          =   195
            Left            =   150
            TabIndex        =   46
            Top             =   2430
            Width           =   1485
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Valor Préstamo (Máx):"
            Height          =   195
            Left            =   5430
            TabIndex        =   45
            Top             =   2130
            Width           =   1545
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Valor Préstamo (Min):"
            Height          =   195
            Left            =   150
            TabIndex        =   44
            Top             =   2100
            Width           =   1500
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Inicio:"
            Height          =   195
            Left            =   150
            TabIndex        =   43
            Top             =   1770
            Width           =   915
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Fin:"
            Height          =   195
            Left            =   5430
            TabIndex        =   42
            Top             =   1785
            Width           =   750
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Left            =   150
            TabIndex        =   39
            Top             =   120
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Producto:"
            Height          =   195
            Left            =   150
            TabIndex        =   38
            Top             =   450
            Width           =   690
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Evaluación"
            Height          =   195
            Left            =   150
            TabIndex        =   37
            Top             =   1110
            Width           =   1380
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   150
            TabIndex        =   36
            Top             =   3750
            Width           =   540
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Proyecto:"
            Height          =   195
            Left            =   180
            TabIndex        =   35
            Top             =   780
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   40
         Top             =   60
         Width           =   9750
         _Version        =   65536
         _ExtentX        =   17198
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
         Begin Threed.SSPanel frm_Produc_10 
            Height          =   480
            Left            =   630
            TabIndex        =   41
            Top             =   90
            Width           =   7185
            _Version        =   65536
            _ExtentX        =   12674
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Tasas de Créditos Hipotecario"
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
            Picture         =   "PltPar_frm_043.frx":14B8
            Top             =   90
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_Produc_10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_Produc()      As moddat_tpo_Genera
Dim l_arr_Proyec()      As moddat_tpo_Genera


Private Sub cmb_CodPry_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipBon)  'ipp_FecIni
   End If
End Sub

Private Sub cmb_Situac_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub


Private Sub cmb_TipBon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecIni)
   End If
End Sub

Private Sub cmb_TipEva_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipBon) 'cmb_CodPry
   End If
End Sub

Private Sub cmb_TipPro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CodPry)  'cmb_TipEva
   End If
End Sub

Private Sub cmb_VerSit_Click()
   Call fs_Buscar
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   Call fs_Activa(False)
   cmb_TipEva.Text = Trim$("<< TODOS >>>")
   Call gs_SetFocus(txt_Descri)
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 860
   grd_Listad.ColWidth(1) = 5030
   grd_Listad.ColWidth(2) = 1100
   grd_Listad.ColWidth(3) = 1100
   grd_Listad.ColWidth(4) = 1100
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   
   Call moddat_gs_Carga_Produc_Comerc(cmb_TipPro, l_arr_Produc, 4)
   cmb_TipPro.AddItem Trim$("<< TODOS >>>")
   ReDim Preserve l_arr_Produc(UBound(l_arr_Produc) + 1)
   l_arr_Produc(UBound(l_arr_Produc)).Genera_Codigo = "000"
   l_arr_Produc(UBound(l_arr_Produc)).Genera_Nombre = "<< TODOS >>>"
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipEva, 1, "038")
   cmb_TipEva.AddItem Trim$("<< TODOS >>>")
   cmb_TipEva.ItemData(cmb_TipEva.NewIndex) = "000"
      
   Call moddat_gs_Carga_Proyec(cmb_CodPry, l_arr_Proyec)
   cmb_CodPry.AddItem "<< TODOS >>>"
   ReDim Preserve l_arr_Proyec(UBound(l_arr_Proyec) + 1)
   l_arr_Proyec(UBound(l_arr_Proyec)).Genera_Codigo = "000000"
   l_arr_Proyec(UBound(l_arr_Proyec)).Genera_Nombre = "<< TODOS >>>"
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipBon, 1, "532")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_VerSit, 1, "013")
   cmb_VerSit.AddItem Trim$("<<TODOS>>")
   cmb_VerSit.ItemData(cmb_VerSit.NewIndex) = 0
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_Situac, 1, "013")
   ipp_FecFin.Text = Format(Now, "dd/mm/yyyy")
   ipp_FecIni.Text = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   grd_Listad.Enabled = p_Activa
   cmd_Agrega.Enabled = p_Activa
   cmd_Editar.Enabled = p_Activa
   cmd_Borrar.Enabled = p_Activa
   cmb_VerSit.Enabled = p_Activa
   
   txt_Descri.Enabled = Not p_Activa
   cmb_TipPro.Enabled = Not p_Activa
   cmb_TipEva.Enabled = False ' not p_Activa
   cmb_CodPry.Enabled = Not p_Activa
   cmb_TipBon.Enabled = Not p_Activa
   cmb_Situac.Enabled = Not p_Activa
   ipp_ValPreMin.Enabled = Not p_Activa
   ipp_ValPreMax.Enabled = Not p_Activa
   ipp_ValInmMin.Enabled = Not p_Activa
   ipp_ValInmMax.Enabled = Not p_Activa
   ipp_PlaMin.Enabled = Not p_Activa
   ipp_PlaMax.Enabled = Not p_Activa
   ipp_PorIniMin.Enabled = Not p_Activa
   ipp_PorIniMax.Enabled = Not p_Activa
   ipp_TasIntAct.Enabled = Not p_Activa
   ipp_TasIntPas.Enabled = Not p_Activa
   ipp_FecIni.Enabled = Not p_Activa
   ipp_FecFin.Enabled = Not p_Activa
   cmd_Grabar.Enabled = Not p_Activa
   cmd_Cancel.Enabled = Not p_Activa
End Sub

Private Sub fs_Limpia()
   txt_Descri.Text = ""
   cmb_TipPro.ListIndex = -1
   cmb_TipEva.ListIndex = -1
   cmb_CodPry.ListIndex = -1
   cmb_TipBon.ListIndex = -1
   cmb_Situac.ListIndex = -1
   cmb_VerSit.ListIndex = -1
   ipp_ValPreMin.Value = 0
   ipp_ValPreMax.Value = 0
   ipp_ValInmMin.Value = 0
   ipp_ValInmMax.Value = 0
   ipp_PlaMin.Value = 0
   ipp_PlaMax.Value = 0
   ipp_PorIniMin.Value = 0
   ipp_PorIniMax.Value = 0
   ipp_TasIntAct.Value = 0
   ipp_TasIntPas.Value = 0
   ipp_FecIni.Value = 0
   ipp_FecFin.Value = 0
End Sub

Private Sub cmd_Borrar_Click()
   If grd_Listad.Rows = 0 Then Exit Sub
   
   grd_Listad.Col = 0
   moddat_g_str_Codigo = grd_Listad.Text
   Call gs_RefrescaGrid(grd_Listad)

   If MsgBox("¿Está seguro de eliminar tasa seleccionada?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   g_str_Parame = "USP_CRE_TASPRD_BORRAR (" & "'" & moddat_g_str_Codigo & "', "
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

Private Sub fs_Buscar()
   If cmb_VerSit.ListIndex = -1 Then
      Exit Sub
   End If
   
   cmd_Agrega.Enabled = True
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   grd_Listad.Enabled = False
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT TASPRD_CODITE  , TRIM(TASPRD_DESCRI) TASPRD_DESCRI  , TASPRD_TASINT_ACT, TASPRD_TASINT_PAS, "
   g_str_Parame = g_str_Parame & "        TASPRD_SITUAC  , TRIM(B.PARDES_DESCRI) AS SITUACION " ', TRIM(C.PARDES_DESCRI) AS TIPO_BONO "
   g_str_Parame = g_str_Parame & "   FROM CRE_TASPRD A  "
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = 013 AND B.PARDES_CODITE = A.TASPRD_SITUAC  "
  ' g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 532 AND C.PARDES_CODITE = A.TASPRD_SITUAC  "
   g_str_Parame = g_str_Parame & "  WHERE TASPRD_SITUAC <> 0  "
   If cmb_VerSit.ItemData(cmb_VerSit.ListIndex) <> 0 Then
      g_str_Parame = g_str_Parame & "    AND TASPRD_SITUAC = " & cmb_VerSit.ItemData(cmb_VerSit.ListIndex)
   End If
   g_str_Parame = g_str_Parame & "  ORDER BY TASPRD_CODITE ASC  "

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
      grd_Listad.Text = Trim(g_rst_Genera!TASPRD_CODITE)
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Genera!TASPRD_DESCRI)
      
      grd_Listad.Col = 2
      grd_Listad.Text = Format(g_rst_Genera!TASPRD_TASINT_ACT, "0.00")
           
      grd_Listad.Col = 3
      If Not IsNull(g_rst_Genera!TASPRD_TASINT_PAS) Then
         grd_Listad.Text = Format(g_rst_Genera!TASPRD_TASINT_PAS, "0.00")
      Else
         grd_Listad.Text = Format(0, "0.00")
      End If
      
      grd_Listad.Col = 4
      grd_Listad.Text = Trim(g_rst_Genera!SITUACION)
      
      g_rst_Genera.MoveNext
   Loop
   grd_Listad.Redraw = True
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   If grd_Listad.Rows > 0 Then
      cmd_Editar.Enabled = True
      cmd_Borrar.Enabled = True
      grd_Listad.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
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
      grd_Listad.Enabled = False
   End If
End Sub

Private Sub cmd_Editar_Click()
   If grd_Listad.Rows = 0 Then Exit Sub
   grd_Listad.Col = 0
   moddat_g_str_Codigo = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)
   moddat_g_int_FlgGrb = 2
   
   'Obteniendo Información del Registro
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_TASPRD WHERE "
   g_str_Parame = g_str_Parame & "TASPRD_CODITE = '" & moddat_g_str_Codigo & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   txt_Descri.Text = Trim(g_rst_Genera!TASPRD_DESCRI)
   'Call moddat_gs_Carga_Produc(cmb_TipPro, l_arr_Produc, 4)
   'Call gs_BuscarCombo_Text(cmb_TipPro, g_rst_Genera!TASPRD_CODPRD, 1)
   'Call gs_BuscarCombo_Text(cmb_TipEva, g_rst_Genera!TASPRD_TIPEVA, 1)
   'Call gs_BuscarCombo_Item(cmb_CodPry, g_rst_Genera!TASPRD_TIPPRY)
   
   cmb_TipPro.ListIndex = gf_Busca_Arregl(l_arr_Produc, g_rst_Genera!TASPRD_CODPRD) - 1
   Call gs_BuscarCombo_Item(cmb_TipEva, g_rst_Genera!TASPRD_TIPEVA)
   cmb_CodPry.ListIndex = gf_Busca_Arregl(l_arr_Proyec, Trim(g_rst_Genera!TASPRD_TIPPRY)) - 1
   Call gs_BuscarCombo_Item(cmb_Situac, g_rst_Genera!TASPRD_SITUAC)
   
   If Not IsNull(g_rst_Genera!TASPRD_TIPBON) Then
      Call gs_BuscarCombo_Item(cmb_TipBon, g_rst_Genera!TASPRD_TIPBON)
   End If
   ipp_FecIni.Text = gf_FormatoFecha(CStr(g_rst_Genera!TASPRD_FECINI))
   ipp_FecFin.Text = gf_FormatoFecha(CStr(g_rst_Genera!TASPRD_FECFIN))
   ipp_ValPreMin.Value = g_rst_Genera!TASPRD_VALPRE_MIN
   ipp_ValPreMax.Value = g_rst_Genera!TASPRD_VALPRE_MAX
   ipp_ValInmMin.Value = g_rst_Genera!TASPRD_VALINM_MIN
   ipp_ValInmMax.Value = g_rst_Genera!TASPRD_VALINM_MAX
   ipp_PlaMin.Value = g_rst_Genera!TASPRD_PLZPRE_MIN
   ipp_PlaMax.Value = g_rst_Genera!TASPRD_PLZPRE_MAX
'  ipp_PorIniMin.Value = g_rst_Genera!TASPRD_EDACLI_MIN
'  ipp_PorIniMax.Value = g_rst_Genera!TASPRD_EDACLI_MAX
   ipp_PorIniMin.Value = g_rst_Genera!TASPRD_PORINI_MIN
   ipp_PorIniMax.Value = g_rst_Genera!TASPRD_PORINI_MAX
   ipp_TasIntAct.Value = g_rst_Genera!TASPRD_TASINT_ACT
   If Not IsNull(g_rst_Genera!TASPRD_TASINT_PAS) Then
      ipp_TasIntPas.Value = g_rst_Genera!TASPRD_TASINT_PAS
   Else
      ipp_TasIntPas.Value = 0
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   Call fs_Activa(False)
   txt_Descri.Enabled = False
   Call gs_SetFocus(txt_Descri)
End Sub

Private Sub cmd_Grabar_Click()
Dim r_int_CodIte  As Integer

   If Len(Trim(txt_Descri.Text)) = 0 Then
      MsgBox "Debe ingresar la Descripción del Parámetro.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Descri)
      Exit Sub
   End If
   If cmb_TipPro.ListIndex = -1 Then
      MsgBox "Debe seleccionar Producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipPro)
      Exit Sub
   End If
   If cmb_TipEva.ListIndex = -1 Then
      MsgBox "Debe seleccionar Tipo de Evaluación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipEva)
      Exit Sub
   End If
   If cmb_CodPry.ListIndex = -1 Then
      MsgBox "Debe seleccionar Proyecto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodPry)
      Exit Sub
   End If
'   If ipp_FecIni.Value = ipp_FecFin.Value Then
'      MsgBox "Debe ingresar una rango de fechas válido.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(ipp_FecIni)
'      Exit Sub
'   End If
   If cmb_TipBon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Bono.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipBon)
      Exit Sub
   End If
   
   If Format(ipp_FecIni.Text, "yyyymmdd") > Format(ipp_FecFin.Text, "yyyymmdd") Then
      MsgBox "Debe ingresar una Fecha de Inicio válida.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecIni)
      Exit Sub
   End If
   
   'VALOR DEL PRESTAMO
   If CDbl(ipp_ValPreMin.Value) = 0 Then
      MsgBox "Debe ingresar el Valor Mínimo del Préstamo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ValPreMin)
      Exit Sub
   End If
   
   If CDbl(ipp_ValPreMax.Value) = 0 Then
      MsgBox "Debe ingresar el Valor Máximo del Préstamo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ValPreMax)
      Exit Sub
   End If
   
   If CDbl(ipp_ValPreMin.Value) > CDbl(ipp_ValPreMax.Value) Then
      MsgBox "El Valor Mínimo del Préstamo no puede ser mayor al Valor Máximo del Préstamo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ValPreMin)
      Exit Sub
   End If
   
   If CDbl(ipp_ValPreMin.Value) = 0 And CDbl(ipp_ValPreMax.Value) > 0 Then
      MsgBox "Debe ingresar el Valor Mínimo del Préstamo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ValPreMin)
      Exit Sub
   End If
   
   If CDbl(ipp_ValPreMin.Value) > 0 And CDbl(ipp_ValPreMax.Value) = 0 Then
      MsgBox "Debe ingresar el Valor Máximo del Préstamo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ValPreMax)
      Exit Sub
   End If
   
   'VALOR DEL INMUEBLE
   If CDbl(ipp_ValInmMin.Value) = 0 Then
      MsgBox "Debe ingresar el Valor Mínimo del Inmueble.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ValInmMin)
      Exit Sub
   End If
   
   If CDbl(ipp_ValInmMax.Value) = 0 Then
      MsgBox "Debe ingresar el Valor Máximo del Inmueble.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ValInmMax)
      Exit Sub
   End If
   
   If CDbl(ipp_ValInmMin.Value) > CDbl(ipp_ValInmMax.Value) Then
      MsgBox "El Valor Mínimo del Inmueble no puede ser mayor al Valor Máximo del Inmueble.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ValInmMin)
      Exit Sub
   End If
   
   If CDbl(ipp_ValInmMin.Value) = 0 And CDbl(ipp_ValInmMax.Value) > 0 Then
      MsgBox "Debe ingresar el Valor Mínimo del Inmueble.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ValInmMin)
      Exit Sub
   End If
   
   If CDbl(ipp_ValInmMin.Value) > 0 And CDbl(ipp_ValInmMax.Value) = 0 Then
      MsgBox "Debe ingresar el Valor Máximo del Inmueble.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ValInmMax)
      Exit Sub
   End If
   
   'PLAZO
   If CDbl(ipp_PlaMin.Value) = 0 Then
      MsgBox "Debe ingresar el Plazo Mínimo(Meses).", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaMin)
      Exit Sub
   End If
   
   If CDbl(ipp_PlaMax.Value) = 0 Then
      MsgBox "Debe ingresar el Plazo Máximo(Meses).", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaMax)
      Exit Sub
   End If
   
   If CInt(ipp_PlaMin.Value) > CInt(ipp_PlaMax.Value) Then
      MsgBox "El Plazo Mínimo no puede ser mayor al Plazo Máximo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaMin)
      Exit Sub
   End If
   
   If CDbl(ipp_PlaMin.Value) = 0 And CDbl(ipp_PlaMax.Value) > 0 Then
      MsgBox "Debe ingresar el Plazo Mínimo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaMin)
      Exit Sub
   End If
   
   If CDbl(ipp_PlaMin.Value) > 0 And CDbl(ipp_PlaMax.Value) = 0 Then
      MsgBox "Debe ingresar el Plazo Máximo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaMax)
      Exit Sub
   End If
   
   'PORCENTAJE INICIAL
   If CDbl(ipp_PorIniMin.Value) = 0 Then
      MsgBox "Debe ingresar el Porcentaje de Inicial Mínimo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PorIniMin)
      Exit Sub
   End If
   
'   If CDbl(ipp_PorIniMin.Value) > CDbl(ipp_PorIniMax.Value) Then
'      MsgBox "El Porcentaje Inicial Mínimo no puede ser mayor al Porcentaje Inicial Máximo.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(ipp_PorIniMin)
'      Exit Sub
'   End If
   
   If CDbl(ipp_PorIniMin.Value) = 0 And CDbl(ipp_PorIniMax.Value) > 0 Then
      MsgBox "Debe ingresar el Porcentaje Inicial Mínimo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PorIniMin)
      Exit Sub
   End If
   
'   If CDbl(ipp_PorIniMin.Value) > 0 And CDbl(ipp_PorIniMax.Value) = 0 Then
'      MsgBox "Debe ingresar el Porcentaje Inicial Máximo.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(ipp_PorIniMax)
'      Exit Sub
'   End If
   
   If ipp_TasIntAct.Value = 0# Then
      MsgBox "Debe ingresar valor de Tasa Activa.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_TasIntAct)
      Exit Sub
   End If
   
'   If ipp_TasIntPas.Value = 0# Then
'      MsgBox "Debe ingresar valor de Tasa Pasiva.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(ipp_TasIntPas)
'      Exit Sub
'   End If
   
   If cmb_Situac.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Estado del Parámetro.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Situac)
      Exit Sub
   End If
   
   If moddat_g_int_FlgGrb = 1 Then
      'Validar que el registro no exista
      
      If fs_Validar_Iteracion(l_arr_Produc(cmb_TipPro.ListIndex + 1).Genera_Codigo, cmb_TipEva.ItemData(cmb_TipEva.ListIndex), l_arr_Proyec(cmb_CodPry.ListIndex + 1).Genera_Codigo, cmb_TipBon.ItemData(cmb_TipBon.ListIndex), Format(ipp_FecIni.Text, "yyyymmdd"), Format(ipp_FecFin.Text, "yyyymmdd"), CDbl(ipp_ValPreMin.Value), CDbl(ipp_ValPreMax.Value), CDbl(ipp_ValInmMin.Value), CDbl(ipp_ValInmMax.Value), CInt(ipp_PlaMin.Value), CInt(ipp_PlaMax.Value), CDbl(ipp_PorIniMin.Value), CDbl(ipp_PorIniMax.Value)) = True Then
         MsgBox "El Parámetro ya ha sido registrado. Por favor verifique e intente nuevamente.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   End If
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      
      If moddat_g_int_FlgGrb = 1 Then
         r_int_CodIte = fs_GeneraCodIte
      Else
         r_int_CodIte = moddat_g_str_Codigo
      End If
      
      g_str_Parame = "USP_CRE_TASPRD ("
      g_str_Parame = g_str_Parame & CStr(r_int_CodIte) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_Descri.Text & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_Produc(cmb_TipPro.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_TipEva.ItemData(cmb_TipEva.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & l_arr_Proyec(cmb_CodPry.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_TipBon.ItemData(cmb_TipBon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValPreMin.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValPreMax.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValInmMin.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValInmMax.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_PlaMin.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_PlaMax.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_PorIniMin.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_PorIniMax.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_Situac.ItemData(cmb_Situac.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_TasIntAct.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_TasIntPas.Value) & ", "

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

Private Function fs_Validar_Iteracion(p_CodPrd As String, p_TipEva As Integer, p_CodPry As String, p_CodBon As Integer, p_FecIni As Long, p_FecFin As Long, p_ValPre_Min As Double, p_ValPre_Max As Double, p_ValInm_Min As Double, p_ValInm_Max As Double, p_PlzPre_Min As Integer, p_PlzPre_Max As Integer, p_PorIni_Min As Double, p_PorIni_Max As Double) As Boolean
   fs_Validar_Iteracion = False
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT TASPRD_TASINT_ACT FROM CRE_TASPRD "
   g_str_Parame = g_str_Parame & " WHERE TASPRD_CODPRD = '" & p_CodPrd & "' "
   g_str_Parame = g_str_Parame & "   AND TASPRD_TIPEVA = " & p_TipEva & " "
   g_str_Parame = g_str_Parame & "   AND TASPRD_TIPPRY = '" & p_CodPry & "' "
   g_str_Parame = g_str_Parame & "   AND TASPRD_TIPBON = '" & p_CodBon & "' "
   g_str_Parame = g_str_Parame & "   AND TASPRD_FECINI = " & p_FecIni & " "
   g_str_Parame = g_str_Parame & "   AND TASPRD_FECFIN = " & p_FecFin & " "
   g_str_Parame = g_str_Parame & "   AND TASPRD_VALPRE_MIN = " & p_ValPre_Min & " "
   g_str_Parame = g_str_Parame & "   AND TASPRD_VALPRE_MAX = " & p_ValPre_Max & " "
   g_str_Parame = g_str_Parame & "   AND TASPRD_VALINM_MIN = " & p_ValInm_Min & " "
   g_str_Parame = g_str_Parame & "   AND TASPRD_VALINM_MAX = " & p_ValInm_Max & " "
   g_str_Parame = g_str_Parame & "   AND TASPRD_PLZPRE_MIN = " & p_PlzPre_Min & " "
   g_str_Parame = g_str_Parame & "   AND TASPRD_PLZPRE_MAX = " & p_PlzPre_Max & " "
   g_str_Parame = g_str_Parame & "   AND TASPRD_PORINI_MIN = " & p_PorIni_Min & " "
   g_str_Parame = g_str_Parame & "   AND TASPRD_PORINI_MAX = " & p_PorIni_Max & " "
   g_str_Parame = g_str_Parame & "   AND TASPRD_SITUAC = '" & cmb_Situac.ItemData(cmb_Situac.ListIndex) & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Function
   End If
                   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      fs_Validar_Iteracion = True
   End If
                                 
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Private Function fs_GeneraCodIte() As Integer
Dim r_str_Parame     As String

   fs_GeneraCodIte = 0
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT NVL(MAX(TASPRD_CODITE),0) CODITE FROM CRE_TASPRD "
   
   If Not gf_EjecutaSQL(r_str_Parame, g_rst_GenAux, 3) Then
       Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Function
   End If
     
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      fs_GeneraCodIte = g_rst_GenAux!CODITE + 1
   End If
End Function

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 0
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Activa(True)
   Call fs_Limpia
   cmb_VerSit.ListIndex = 0
   
   Call gs_CentraForm(Me)
   Call gs_SetFocus(grd_Listad)
   Screen.MousePointer = 0
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_Editar_Click
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValPreMin)
   End If
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFin)
   End If
End Sub

Private Sub ipp_PlaMax_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PorIniMin)
   End If
End Sub

Private Sub ipp_PlaMin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PlaMax)
   End If
End Sub

Private Sub ipp_PorIniMax_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_TasIntAct)
   End If
End Sub

Private Sub ipp_PorIniMin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PorIniMax)
   End If
End Sub

Private Sub ipp_TasInt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub ipp_TasIntAct_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_TasIntPas)
   End If
End Sub
Private Sub ipp_TasIntPas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Situac)
   End If
End Sub

Private Sub ipp_ValInmMax_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PlaMin)
   End If
End Sub

Private Sub ipp_ValInmMin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValInmMax)
   End If
End Sub

Private Sub ipp_ValPreMax_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValInmMin)
   End If
End Sub

Private Sub ipp_ValPreMin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValPreMax)
   End If
End Sub

Private Sub txt_Descri_GotFocus()
   Call gs_SelecTodo(txt_Descri)
End Sub

Private Sub txt_Descri_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipPro)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & ",-_.( )$/")
   End If
End Sub
