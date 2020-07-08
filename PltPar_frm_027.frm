VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7950
   ClientLeft      =   765
   ClientTop       =   1845
   ClientWidth     =   14040
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   14040
   Begin Threed.SSPanel SSPanel1 
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14025
      _Version        =   65536
      _ExtentX        =   24739
      _ExtentY        =   13996
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
         Height          =   765
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   13935
         _Version        =   65536
         _ExtentX        =   24580
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
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   7500
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   210
            Width           =   4095
         End
         Begin VB.ComboBox cmb_Carter 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   210
            Width           =   4095
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   13200
            Picture         =   "PltPar_frm_027.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   12510
            Picture         =   "PltPar_frm_027.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   11790
            Picture         =   "PltPar_frm_027.frx":074C
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Buscar Modalidades"
            Top             =   30
            Width           =   675
         End
         Begin VB.Label Label7 
            Caption         =   "Seleccione Sector:"
            Height          =   285
            Left            =   5940
            TabIndex        =   34
            Top             =   210
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Seleccione Cartera:"
            Height          =   285
            Left            =   60
            TabIndex        =   6
            Top             =   210
            Width           =   1455
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   765
         Left            =   30
         TabIndex        =   7
         Top             =   5130
         Width           =   13935
         _Version        =   65536
         _ExtentX        =   24580
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
         Begin VB.CommandButton cmd_Borrar 
            Height          =   675
            Left            =   13200
            Picture         =   "PltPar_frm_027.frx":0A56
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   675
            Left            =   11820
            Picture         =   "PltPar_frm_027.frx":0D60
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   675
            Left            =   12510
            Picture         =   "PltPar_frm_027.frx":106A
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   765
         Left            =   30
         TabIndex        =   11
         Top             =   7110
         Width           =   13935
         _Version        =   65536
         _ExtentX        =   24580
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
            Left            =   13230
            Picture         =   "PltPar_frm_027.frx":1374
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   12540
            Picture         =   "PltPar_frm_027.frx":167E
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   1125
         Left            =   30
         TabIndex        =   14
         Top             =   5940
         Width           =   13935
         _Version        =   65536
         _ExtentX        =   24580
         _ExtentY        =   1984
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
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   60
            Width           =   12255
         End
         Begin VB.ComboBox cmb_SecEco 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   390
            Width           =   12255
         End
         Begin EditLib.fpDoubleSingle ipp_Porcen 
            Height          =   315
            Left            =   1590
            TabIndex        =   16
            Top             =   720
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
            ButtonMax       =   100
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
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
            ThreeDFrameColor=   -2147483633
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin Threed.SSPanel pnl_MetSec 
            Height          =   315
            Left            =   12390
            TabIndex        =   17
            Top             =   720
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "2,000,000.00 "
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin VB.Label Label9 
            Caption         =   "Giro (Grupos):"
            Height          =   315
            Left            =   60
            TabIndex        =   38
            Top             =   60
            Width           =   1485
         End
         Begin VB.Label Label2 
            Caption         =   "% Colocación:"
            Height          =   315
            Left            =   60
            TabIndex        =   20
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label Label3 
            Caption         =   "Giro (Detalles):"
            Height          =   315
            Left            =   60
            TabIndex        =   19
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label6 
            Caption         =   "Meta x Giro:"
            Height          =   285
            Left            =   10980
            TabIndex        =   18
            Top             =   720
            Width           =   1275
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   4245
         Left            =   30
         TabIndex        =   21
         Top             =   840
         Width           =   13935
         _Version        =   65536
         _ExtentX        =   24580
         _ExtentY        =   7488
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
            TabIndex        =   22
            Top             =   1200
            Width           =   13815
            _ExtentX        =   24368
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
            Left            =   10470
            TabIndex        =   23
            Top             =   3900
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
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
            Left            =   10470
            TabIndex        =   24
            Top             =   900
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
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
            TabIndex        =   25
            Top             =   900
            Width           =   8445
            _Version        =   65536
            _ExtentX        =   14896
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Giro Comercial"
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
            Left            =   11820
            TabIndex        =   26
            Top             =   900
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
            TabIndex        =   27
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
            TabIndex        =   28
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
            Height          =   90
            Left            =   30
            TabIndex        =   29
            Top             =   750
            Width           =   13875
            _Version        =   65536
            _ExtentX        =   24474
            _ExtentY        =   159
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
            Left            =   11820
            TabIndex        =   30
            Top             =   3900
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   315
            Left            =   1620
            TabIndex        =   35
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
         Begin VB.Label Label8 
            Caption         =   "Meta Coloc. Sector:"
            Height          =   315
            Left            =   60
            TabIndex        =   36
            Top             =   390
            Width           =   1425
         End
         Begin VB.Label Label4 
            Caption         =   "Meta Coloc. Cart.:"
            Height          =   315
            Left            =   60
            TabIndex        =   32
            Top             =   60
            Width           =   1305
         End
         Begin VB.Label Label5 
            Caption         =   "Totales ==>"
            Height          =   315
            Left            =   9450
            TabIndex        =   31
            Top             =   3900
            Width           =   915
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

