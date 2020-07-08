VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Produc_07 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   8205
   ClientLeft      =   3750
   ClientTop       =   960
   ClientWidth     =   8205
   Icon            =   "PltPar_frm_029.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8205
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8265
      _Version        =   65536
      _ExtentX        =   14579
      _ExtentY        =   14473
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
         Height          =   735
         Left            =   30
         TabIndex        =   25
         Top             =   7410
         Width           =   8145
         _Version        =   65536
         _ExtentX        =   14367
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
         Begin EditLib.fpLongInteger ipp_DiaIni 
            Height          =   315
            Left            =   1200
            TabIndex        =   16
            Top             =   240
            Width           =   1005
            _Version        =   196608
            _ExtentX        =   1773
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
         Begin EditLib.fpLongInteger ipp_DiaFin 
            Height          =   315
            Left            =   3810
            TabIndex        =   17
            Top             =   240
            Width           =   1005
            _Version        =   196608
            _ExtentX        =   1773
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
         Begin EditLib.fpDoubleSingle ipp_Import 
            Height          =   315
            Left            =   6210
            TabIndex        =   18
            Top             =   240
            Width           =   1005
            _Version        =   196608
            _ExtentX        =   1773
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
         Begin VB.Label Label6 
            Caption         =   "Rango Inicio:"
            Height          =   285
            Left            =   90
            TabIndex        =   28
            Top             =   270
            Width           =   1005
         End
         Begin VB.Label Label3 
            Caption         =   "Rango Fin:"
            Height          =   285
            Left            =   2865
            TabIndex        =   27
            Top             =   270
            Width           =   825
         End
         Begin VB.Label Label41 
            Caption         =   "Importe:"
            Height          =   285
            Left            =   5490
            TabIndex        =   26
            Top             =   270
            Width           =   795
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   735
         Left            =   30
         TabIndex        =   10
         Top             =   6630
         Width           =   8145
         _Version        =   65536
         _ExtentX        =   14367
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
            Left            =   7440
            Picture         =   "PltPar_frm_029.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   6060
            Picture         =   "PltPar_frm_029.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Cancel 
            Height          =   675
            Left            =   6750
            Picture         =   "PltPar_frm_029.frx":0890
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   675
            Left            =   5370
            Picture         =   "PltPar_frm_029.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   675
            Left            =   4680
            Picture         =   "PltPar_frm_029.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   675
            Left            =   3990
            Picture         =   "PltPar_frm_029.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   2025
         Left            =   30
         TabIndex        =   15
         Top             =   4560
         Width           =   8145
         _Version        =   65536
         _ExtentX        =   14367
         _ExtentY        =   3572
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
            Height          =   1665
            Left            =   60
            TabIndex        =   21
            Top             =   330
            Width           =   8025
            _ExtentX        =   14155
            _ExtentY        =   2937
            _Version        =   393216
            Rows            =   1
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   1860
            TabIndex        =   22
            Top             =   60
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Rango Final"
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
            TabIndex        =   23
            Top             =   60
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Rango Inicial"
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   285
            Left            =   3630
            TabIndex        =   24
            Top             =   60
            Width           =   2415
            _Version        =   65536
            _ExtentX        =   4260
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Importe"
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
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   30
         TabIndex        =   29
         Top             =   30
         Width           =   8145
         _Version        =   65536
         _ExtentX        =   14367
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   480
            Left            =   630
            TabIndex        =   30
            Top             =   90
            Width           =   3105
            _Version        =   65536
            _ExtentX        =   5477
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Gastos de Cobranzas"
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
            Picture         =   "PltPar_frm_029.frx":14B8
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   765
         Left            =   30
         TabIndex        =   31
         Top             =   750
         Width           =   8145
         _Version        =   65536
         _ExtentX        =   14367
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
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   1230
            TabIndex        =   32
            Top             =   60
            Width           =   6855
            _Version        =   65536
            _ExtentX        =   12091
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "SSPanel3"
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
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_SubPrd 
            Height          =   315
            Left            =   1230
            TabIndex        =   33
            Top             =   390
            Width           =   6855
            _Version        =   65536
            _ExtentX        =   12091
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "SSPanel3"
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
            Alignment       =   1
         End
         Begin VB.Label Label10 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   120
            TabIndex        =   35
            Top             =   90
            Width           =   1605
         End
         Begin VB.Label Label1 
            Caption         =   "Sub-Producto:"
            Height          =   315
            Left            =   120
            TabIndex        =   34
            Top             =   420
            Width           =   1605
         End
      End
      Begin Threed.SSPanel SSPanel12 
         Height          =   1395
         Left            =   30
         TabIndex        =   5
         Top             =   1560
         Width           =   8145
         _Version        =   65536
         _ExtentX        =   14367
         _ExtentY        =   2461
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
         Begin MSFlexGridLib.MSFlexGrid grd_ListaCab 
            Height          =   1035
            Left            =   60
            TabIndex        =   36
            Top             =   330
            Width           =   8025
            _ExtentX        =   14155
            _ExtentY        =   1826
            _Version        =   393216
            Rows            =   1
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel13 
            Height          =   285
            Left            =   1860
            TabIndex        =   37
            Top             =   60
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Rango Final"
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
         Begin Threed.SSPanel SSPanel14 
            Height          =   285
            Left            =   90
            TabIndex        =   38
            Top             =   60
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Rango Inicial"
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
      Begin Threed.SSPanel SSPanel15 
         Height          =   735
         Left            =   30
         TabIndex        =   39
         Top             =   3780
         Width           =   8145
         _Version        =   65536
         _ExtentX        =   14367
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
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1140
            TabIndex        =   6
            Top             =   225
            Width           =   1305
            _Version        =   196608
            _ExtentX        =   2302
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
            Height          =   315
            Left            =   3600
            TabIndex        =   7
            Top             =   225
            Width           =   1305
            _Version        =   196608
            _ExtentX        =   2302
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
         Begin VB.Label Label4 
            Caption         =   "Rango Fin:"
            Height          =   195
            Left            =   2730
            TabIndex        =   41
            Top             =   285
            Width           =   780
         End
         Begin VB.Label Label2 
            Caption         =   "Rango Inicio:"
            Height          =   195
            Left            =   90
            TabIndex        =   40
            Top             =   285
            Width           =   945
         End
      End
      Begin Threed.SSPanel SSPanel16 
         Height          =   735
         Left            =   30
         TabIndex        =   1
         Top             =   3000
         Width           =   8145
         _Version        =   65536
         _ExtentX        =   14367
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
         Begin VB.CommandButton cmd_CancelCab 
            Height          =   675
            Left            =   7440
            Picture         =   "PltPar_frm_029.frx":17C2
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_GrabarCab 
            Height          =   675
            Left            =   6750
            Picture         =   "PltPar_frm_029.frx":1ACC
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_AgregaCab 
            Height          =   675
            Left            =   4680
            Picture         =   "PltPar_frm_029.frx":1F0E
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_EditarCab 
            Height          =   675
            Left            =   5370
            Picture         =   "PltPar_frm_029.frx":2218
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_BorrarCab 
            Height          =   675
            Left            =   6060
            Picture         =   "PltPar_frm_029.frx":2522
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frm_Produc_07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   
   Call fs_Limpia
   Call fs_Activa_Editar(True)
   Call gs_SetFocus(ipp_DiaIni)
End Sub

Private Sub cmd_AgregaCab_Click()
   moddat_g_int_FlgGrb = 1
   
   Call fs_LimpiaCab
   Call fs_Activa_Editar_Cab(True)
   Call gs_SetFocus(ipp_FecIni)
End Sub

Private Sub cmd_Borrar_Click()
   Dim r_int_CodRan As Integer
   
   grd_ListaCab.Col = 2
   r_int_CodRan = grd_ListaCab.Text
   
   grd_Listad.Col = 0
   moddat_g_str_CodIte = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)
   
   If MsgBox("¿Está seguro de eliminar el Item?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11

   g_str_Parame = "USP_BORRAR_OPE_GASCOB ("
   g_str_Parame = g_str_Parame & CInt(r_int_CodRan) & ", "
   g_str_Parame = g_str_Parame & CStr(CInt(moddat_g_str_CodIte)) & ", "
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

Private Sub cmd_BorrarCab_Click()
   
   Dim r_int_CodRan As Integer

   moddat_g_str_CodIte = 0
   grd_ListaCab.Col = 2
   moddat_g_str_CodIte = grd_ListaCab.Text
         
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT MAX(GASCOBCAB_CODRAN) AS CODRAN FROM OPE_GASCOB_CAB "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If

   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      r_int_CodRan = IIf(IsNull(g_rst_Genera!CODRAN), 0, g_rst_Genera!CODRAN)
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   If r_int_CodRan > CInt(moddat_g_str_CodIte) Then
      MsgBox "El Rango no puede eliminarse. Por favor verifique e intente nuevamente.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(grd_ListaCab)
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT COUNT(*) AS TOTAL FROM OPE_GASCOB WHERE GASCOB_CODRAN = " & moddat_g_str_CodIte & ""

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      If g_rst_Genera!TOTAL > 0 Then
         MsgBox "El Rango no puede eliminarse, existen datos en OPE_GASCOB. Elimine previamente éstos datos.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(grd_ListaCab)
         Screen.MousePointer = 0
         Exit Sub
      End If
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   Call gs_RefrescaGrid(grd_ListaCab)
   
   If MsgBox("¿Está seguro de eliminar el Item?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11

   g_str_Parame = "USP_BORRAR_OPE_GASCOB_CAB ("
   g_str_Parame = g_str_Parame & CInt(moddat_g_str_CodIte) & ", "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodPrd & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodSub & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
       Exit Sub
   End If
  
   Call fs_Buscar_Cab
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Cancel_Click()
   ipp_DiaIni.Value = 0
   ipp_DiaFin.Value = 0
   ipp_Import.Value = 0

   Call fs_Activa_Editar(False)
   Call gs_SetFocus(grd_Listad)

   If grd_Listad.Rows = 0 Then
      cmd_Agrega.Enabled = True
      cmd_Editar.Enabled = True
      cmd_Borrar.Enabled = True
      cmd_Salida.Enabled = True
   End If
End Sub

Private Sub cmd_CancelCab_Click()
   ipp_FecIni.Value = 0
   ipp_FecFin.Value = 0
   
   Call fs_Activa_Editar_Cab(False)
   Call gs_SetFocus(grd_ListaCab)
   
   If grd_ListaCab.Rows = 0 Then
      cmd_AgregaCab.Enabled = True
      cmd_EditarCab.Enabled = True
      cmd_BorrarCab.Enabled = True
   End If
End Sub

Private Sub cmd_Editar_Click()
Dim r_int_CodRan As Integer
   
   grd_Listad.Col = 2
   r_int_CodRan = grd_ListaCab.Text
   
   grd_Listad.Col = 0
   moddat_g_str_CodIte = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgGrb = 2
   
   'Obteniendo Información del Registro
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM OPE_GASCOB WHERE "
   g_str_Parame = g_str_Parame & "GASCOB_CODRAN = " & r_int_CodRan & " AND "
   g_str_Parame = g_str_Parame & "GASCOB_DIAINI = " & CStr(CInt(moddat_g_str_CodIte)) & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   
   ipp_DiaIni.Value = g_rst_Genera!GASCOB_DIAINI
   ipp_DiaFin.Value = g_rst_Genera!GASCOB_DIAFIN
   ipp_Import.Value = g_rst_Genera!GASCOB_IMPORT
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   Call fs_Activa_Editar(True)
   
   cmd_Agrega.Enabled = False
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   
   cmd_Salida.Enabled = False
   
   ipp_DiaIni.Enabled = False
   ipp_DiaFin.Enabled = False
   
   Call gs_SetFocus(ipp_Import)
End Sub

Private Sub cmd_EditarCab_Click()
   moddat_g_str_CodIte = 0
   grd_ListaCab.Col = 2
   moddat_g_str_CodIte = grd_ListaCab.Text
         
   Call gs_RefrescaGrid(grd_ListaCab)
   
   moddat_g_int_FlgGrb = 2
   
   'Obteniendo Información del Registro
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT TO_DATE(GASCOBCAB_FECINI,'yyyymmdd') AS FECINI, TO_DATE(GASCOBCAB_FECFIN,'yyyymmdd') AS FECFIN"
   g_str_Parame = g_str_Parame & "  FROM OPE_GASCOB_CAB  "
   g_str_Parame = g_str_Parame & " WHERE GASCOBCAB_CODPRD = '" & moddat_g_str_CodPrd & "' "
   g_str_Parame = g_str_Parame & "   AND GASCOBCAB_CODSUB = '" & moddat_g_str_CodSub & "' "
   g_str_Parame = g_str_Parame & "   AND GASCOBCAB_CODRAN = " & CStr(CInt(moddat_g_str_CodIte)) & " "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   
   ipp_FecIni.Text = g_rst_Genera!FECINI
   ipp_FecFin.Text = g_rst_Genera!FECFIN
     
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   Call fs_Activa_Editar_Cab(True)
   ipp_FecIni.Enabled = False
End Sub

Private Sub cmd_Grabar_Click()
   
   moddat_g_str_CodIte = 0
   grd_ListaCab.Col = 2
   moddat_g_str_CodIte = grd_ListaCab.Text
   
   If CInt(ipp_DiaIni.Text) > CInt(ipp_DiaFin.Text) Then
      MsgBox "El Día Inicial no puede ser mayor al Día Final.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_DiaIni)
      Exit Sub
   End If
   
   If moddat_g_int_FlgGrb = 1 Then
      'Validar que el registro no exista
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM OPE_GASCOB WHERE "
      g_str_Parame = g_str_Parame & "GASCOB_CODRAN = " & CInt(moddat_g_str_CodIte) & " AND "
      g_str_Parame = g_str_Parame & "GASCOB_DIAINI = " & CStr(CInt(ipp_DiaIni.Text)) & " "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
        
         MsgBox "El Rango ya ha sido registrado. Por favor verifique e intente nuevamente.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_OPE_GASCOB ("
         
      g_str_Parame = g_str_Parame & CInt(moddat_g_str_CodIte) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_DiaIni.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_DiaFin.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_Import.Value) & ", "
      
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
   Loop
   
   Call fs_Buscar
   Call cmd_Cancel_Click
   Call gs_SetFocus(cmd_Agrega)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Limpia()
   ipp_DiaIni.Value = 0
   ipp_DiaFin.Value = 0
   ipp_Import.Value = 0
End Sub
Private Sub fs_LimpiaCab()
   ipp_FecIni.Text = Format(Now, "dd/mm/yyyy")
   ipp_FecFin.Text = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub cmd_GrabarCab_Click()
Dim r_int_CodRan As Integer

   If ipp_FecIni.Value > ipp_FecFin.Value Then
      MsgBox "La Fecha Inicial no puede ser mayor a la Fecha Final", vbInformation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecFin)
      Exit Sub
   End If
  
   'Validar que el registro no exista
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM OPE_GASCOB_CAB WHERE "
   g_str_Parame = g_str_Parame & "GASCOBCAB_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "GASCOBCAB_CODSUB = '" & moddat_g_str_CodSub & "' AND "
   g_str_Parame = g_str_Parame & "GASCOBCAB_FECINI = " & Format(ipp_FecIni.Text, "yyyymmdd") & "  "
   If moddat_g_int_FlgGrb <> 1 Then
      g_str_Parame = g_str_Parame & "AND GASCOBCAB_FECFIN = " & Format(ipp_FecFin.Text, "yyyymmdd") & " "
   End If
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If

   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
     
      MsgBox "El Rango ya ha sido registrado. Por favor verifique e intente nuevamente.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecFin)
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
      
   If moddat_g_int_FlgGrb = 1 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT MAX(GASCOBCAB_CODRAN) AS CODRAN FROM OPE_GASCOB_CAB "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         r_int_CodRan = IIf(IsNull(g_rst_Genera!CODRAN), 0, g_rst_Genera!CODRAN) + 1
      Else
         r_int_CodRan = 1
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      If r_int_CodRan = 0 Then
         grd_ListaCab.Col = 2
         r_int_CodRan = grd_ListaCab.Text
      End If
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
  
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_OPE_GASCOB_CAB ("
      g_str_Parame = g_str_Parame & "" & r_int_CodRan & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodPrd & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodSub & "', "
      g_str_Parame = g_str_Parame & Format(ipp_FecIni.Text, "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & Format(ipp_FecFin.Text, "yyyymmdd") & ", "
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
   Loop
   
   Call fs_Buscar_Cab
   Call cmd_CancelCab_Click
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Call fs_Activa_Editar_Cab(False)
   Call fs_Activa_Editar(False)
   Call fs_LimpiaCab
   grd_Listad.Enabled = False
   cmd_Agrega.Enabled = False
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   cmd_Salida.Enabled = False
   grd_Listad.Clear
   grd_Listad.Rows = 0
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   pnl_Produc.Caption = moddat_g_str_CodPrd & " - " & moddat_g_str_NomPrd
   pnl_SubPrd.Caption = moddat_g_str_CodSub & " - " & moddat_g_str_DesSub
   
   Call fs_Inicia
   Call fs_Limpia
   Call fs_LimpiaCab
   Call fs_Activa_Editar_Cab(False)
   Call fs_Activa_Editar(False)
   
   grd_Listad.Enabled = False
   cmd_Agrega.Enabled = False
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   cmd_Salida.Enabled = False
   
   Call fs_Buscar_Cab
      
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla Detalle
   grd_Listad.ColWidth(0) = 1770
   grd_Listad.ColWidth(1) = 1770
   grd_Listad.ColWidth(2) = 2400
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignRightCenter
   
   'Inicializando Rejilla Cabecera
   grd_ListaCab.ColWidth(0) = 1770
   grd_ListaCab.ColWidth(1) = 1770
   grd_ListaCab.ColWidth(2) = 0
   
   grd_ListaCab.ColAlignment(0) = flexAlignCenterCenter
   grd_ListaCab.ColAlignment(1) = flexAlignCenterCenter

End Sub
Private Sub grd_ListaCab_DblClick()
   Call fs_Buscar
End Sub

Private Sub ipp_DiaFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Import)
   End If
End Sub

Private Sub ipp_DiaIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_DiaFin)
   End If
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_GrabarCab)
   End If
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFin)
   End If
End Sub

Private Sub ipp_Import_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
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

Private Sub fs_Activa_Editar(ByVal p_Activa As Integer)
   cmd_Grabar.Enabled = p_Activa
   cmd_Cancel.Enabled = p_Activa
   cmd_Salida.Enabled = p_Activa
   ipp_DiaIni.Enabled = p_Activa
   ipp_DiaFin.Enabled = p_Activa
   ipp_Import.Enabled = p_Activa
   
   grd_Listad.Enabled = Not p_Activa
   cmd_Agrega.Enabled = Not p_Activa
   cmd_Editar.Enabled = Not p_Activa
   cmd_Borrar.Enabled = Not p_Activa
   cmd_Salida.Enabled = Not p_Activa
End Sub
Private Sub fs_Activa_Editar_Cab(ByVal p_Activa As Integer)
   cmd_GrabarCab.Enabled = p_Activa
   cmd_CancelCab.Enabled = p_Activa
   ipp_FecIni.Enabled = p_Activa
   ipp_FecFin.Enabled = p_Activa
  
   grd_ListaCab.Enabled = Not p_Activa
   cmd_AgregaCab.Enabled = Not p_Activa
   cmd_EditarCab.Enabled = Not p_Activa
   cmd_BorrarCab.Enabled = Not p_Activa
End Sub
Private Sub fs_Buscar()
   
   grd_ListaCab.Enabled = False
   cmd_AgregaCab.Enabled = False
   cmd_EditarCab.Enabled = False
   cmd_BorrarCab.Enabled = False
   
   cmd_Agrega.Enabled = True
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   grd_Listad.Enabled = False
   cmd_Salida.Enabled = True
   
   Call gs_LimpiaGrid(grd_Listad)
   
   moddat_g_str_CodIte = 0
   grd_ListaCab.Col = 2
   moddat_g_str_CodIte = grd_ListaCab.Text
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT GASCOB_DIAINI, GASCOB_DIAFIN, GASCOB_IMPORT  "
   g_str_Parame = g_str_Parame & "  FROM OPE_GASCOB "
   g_str_Parame = g_str_Parame & " WHERE GASCOB_CODRAN = " & moddat_g_str_CodIte & ""
   g_str_Parame = g_str_Parame & " ORDER BY GASCOB_DIAINI ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
     Exit Sub
   End If
   
   
   ipp_FecIni.Text = CDate(grd_ListaCab.TextMatrix(grd_ListaCab.Row, 0))
   ipp_FecFin.Text = CDate(grd_ListaCab.TextMatrix(grd_ListaCab.Row, 1))
   
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = CStr(g_rst_Princi!GASCOB_DIAINI)
      
      grd_Listad.Col = 1
      grd_Listad.Text = CStr(g_rst_Princi!GASCOB_DIAFIN)
      
      grd_Listad.Col = 2
      grd_Listad.Text = Format(g_rst_Princi!GASCOB_IMPORT, "###,###,##0.00")
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If grd_Listad.Rows > 0 Then
      cmd_Editar.Enabled = True
      cmd_Borrar.Enabled = True
      grd_Listad.Enabled = True
      cmd_Salida.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub
Private Sub fs_Buscar_Cab()
   
   cmd_AgregaCab.Enabled = True
   cmd_EditarCab.Enabled = False
   cmd_BorrarCab.Enabled = False
   grd_ListaCab.Enabled = False
   
   Call gs_LimpiaGrid(grd_ListaCab)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT GASCOBCAB_CODRAN, TO_DATE (GASCOBCAB_FECINI,'yyyymmdd') AS FECINI, TO_DATE(GASCOBCAB_FECFIN,'yyyymmdd') AS FECFIN"
   g_str_Parame = g_str_Parame & "   FROM OPE_GASCOB_CAB  "
   g_str_Parame = g_str_Parame & "  WHERE GASCOBCAB_CODPRD = '" & moddat_g_str_CodPrd & "'  "
   g_str_Parame = g_str_Parame & "    AND GASCOBCAB_CODSUB = '" & moddat_g_str_CodSub & "' "
   g_str_Parame = g_str_Parame & "  ORDER BY GASCOBCAB_FECINI DESC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
     Exit Sub
   End If
   
   grd_ListaCab.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_ListaCab.Rows = grd_ListaCab.Rows + 1
      grd_ListaCab.Row = grd_ListaCab.Rows - 1
      
      grd_ListaCab.Col = 0
      grd_ListaCab.Text = g_rst_Princi!FECINI
      
      grd_ListaCab.Col = 1
      grd_ListaCab.Text = g_rst_Princi!FECFIN
      
      grd_ListaCab.Col = 2
      grd_ListaCab.Text = g_rst_Princi!GASCOBCAB_CODRAN
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_ListaCab.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If grd_ListaCab.Rows > 0 Then
      cmd_EditarCab.Enabled = True
      cmd_BorrarCab.Enabled = True
      grd_ListaCab.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_ListaCab)
   Call gs_SetFocus(grd_ListaCab)
End Sub
