VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Produc_05 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parámetros x Actividad Económica"
   ClientHeight    =   9240
   ClientLeft      =   3105
   ClientTop       =   570
   ClientWidth     =   8220
   Icon            =   "PltPar_frm_015.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9255
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8235
      _Version        =   65536
      _ExtentX        =   14526
      _ExtentY        =   16325
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
         Height          =   4305
         Left            =   60
         TabIndex        =   20
         Top             =   2370
         Width           =   8115
         _Version        =   65536
         _ExtentX        =   14314
         _ExtentY        =   7594
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   1500
            TabIndex        =   21
            Top             =   60
            Width           =   6315
            _Version        =   65536
            _ExtentX        =   11139
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Descripción Grupo"
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
            Left            =   60
            TabIndex        =   22
            Top             =   60
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Código Grupo"
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   3915
            Left            =   30
            TabIndex        =   23
            Top             =   360
            Width           =   8025
            _ExtentX        =   14155
            _ExtentY        =   6906
            _Version        =   393216
            Rows            =   12
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   765
         Left            =   60
         TabIndex        =   15
         Top             =   6720
         Width           =   8115
         _Version        =   65536
         _ExtentX        =   14314
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
         Begin VB.CommandButton cmd_Imprim 
            Height          =   675
            Left            =   6720
            Picture         =   "PltPar_frm_015.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Imprimir"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_LisIte 
            Height          =   675
            Left            =   7410
            Picture         =   "PltPar_frm_015.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Detalle de Grupo"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   675
            Left            =   6030
            Picture         =   "PltPar_frm_015.frx":0890
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   675
            Left            =   5340
            Picture         =   "PltPar_frm_015.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   675
            Left            =   4650
            Picture         =   "PltPar_frm_015.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   675
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   0
            Top             =   0
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Presentación Preliminar"
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   735
         Left            =   60
         TabIndex        =   12
         Top             =   8430
         Width           =   8115
         _Version        =   65536
         _ExtentX        =   14314
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
            Left            =   6720
            Picture         =   "PltPar_frm_015.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Cancel 
            Height          =   675
            Left            =   7410
            Picture         =   "PltPar_frm_015.frx":15F0
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   735
         Left            =   60
         TabIndex        =   7
         Top             =   1590
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
         Begin VB.ComboBox cmb_ActEco 
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   210
            Width           =   3345
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   7380
            Picture         =   "PltPar_frm_015.frx":18FA
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   6690
            Picture         =   "PltPar_frm_015.frx":1D3C
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   5970
            Picture         =   "PltPar_frm_015.frx":2046
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Buscar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.Label Label2 
            Caption         =   "Seleccione Actividad:"
            Height          =   285
            Left            =   90
            TabIndex        =   11
            Top             =   270
            Width           =   1605
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   855
         Left            =   60
         TabIndex        =   8
         Top             =   7530
         Width           =   8115
         _Version        =   65536
         _ExtentX        =   14314
         _ExtentY        =   1508
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
         Begin VB.TextBox txt_NomGrp 
            Height          =   315
            Left            =   1710
            MaxLength       =   80
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   420
            Width           =   6345
         End
         Begin VB.TextBox txt_CodGrp 
            Height          =   315
            Left            =   1710
            MaxLength       =   3
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   90
            Width           =   825
         End
         Begin VB.Label Label4 
            Caption         =   "Descripción Grupo:"
            Height          =   285
            Left            =   90
            TabIndex        =   10
            Top             =   450
            Width           =   1425
         End
         Begin VB.Label Label3 
            Caption         =   "Código Grupo:"
            Height          =   285
            Left            =   90
            TabIndex        =   9
            Top             =   120
            Width           =   1185
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   675
         Left            =   60
         TabIndex        =   24
         Top             =   60
         Width           =   8115
         _Version        =   65536
         _ExtentX        =   14314
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
            TabIndex        =   25
            Top             =   90
            Width           =   5355
            _Version        =   65536
            _ExtentX        =   9446
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Parámetros x Actividad Económica"
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
            Picture         =   "PltPar_frm_015.frx":2350
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   765
         Left            =   60
         TabIndex        =   26
         Top             =   780
         Width           =   8115
         _Version        =   65536
         _ExtentX        =   14314
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
            Left            =   1200
            TabIndex        =   27
            Top             =   60
            Width           =   6885
            _Version        =   65536
            _ExtentX        =   12144
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
            Left            =   1200
            TabIndex        =   28
            Top             =   390
            Width           =   6885
            _Version        =   65536
            _ExtentX        =   12144
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
            Left            =   90
            TabIndex        =   30
            Top             =   90
            Width           =   1605
         End
         Begin VB.Label Label1 
            Caption         =   "Sub-Producto:"
            Height          =   315
            Left            =   90
            TabIndex        =   29
            Top             =   420
            Width           =   1605
         End
      End
   End
End
Attribute VB_Name = "frm_Produc_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Produc()      As moddat_tpo_Genera

Private Sub cmb_Produc_Click()
   Call gs_SetFocus(cmb_ActEco)
End Sub

Private Sub cmb_Produc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Produc_Click
   End If
End Sub

Private Sub cmb_ActEco_Click()
   Call gs_SetFocus(cmd_Buscar)
End Sub

Private Sub cmb_ActEco_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_ActEco_Click
   End If
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   
   Call fs_Activa_Editar(True)
   Call gs_SetFocus(txt_CodGrp)
End Sub

Private Sub cmd_Borrar_Click()
   grd_Listad.Col = 0
   moddat_g_str_CodGrp = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)

   If MsgBox("¿Está seguro de eliminar el grupo?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   'Obteniendo Información del Registro
   g_str_Parame = "USP_CRE_PARACT_BORRAR ("
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodPrd & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodSub & "', "
   g_str_Parame = g_str_Parame & moddat_g_str_CodMod & ", "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodGrp & "', "
   g_str_Parame = g_str_Parame & "'000', "
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

Private Sub cmd_Buscar_Click()
   If cmb_ActEco.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Actividad Económica.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_ActEco)
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Activa(False)
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Cancel_Click()
   txt_CodGrp.Text = ""
   txt_NomGrp.Text = ""
   
   Call fs_Activa_Editar(False)
   Call gs_SetFocus(grd_Listad)
   
   If grd_Listad.Rows = 0 Then
      cmd_Agrega.Enabled = True
      cmd_Editar.Enabled = False
      cmd_LisIte.Enabled = False
      cmd_Borrar.Enabled = False
      cmd_Imprim.Enabled = False
      grd_Listad.Enabled = False
   End If
End Sub

Private Sub cmd_Editar_Click()
   grd_Listad.Col = 0
   moddat_g_str_CodGrp = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgGrb = 2
   
   
   'Obteniendo Información del Registro
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARACT WHERE "
   g_str_Parame = g_str_Parame & "PARACT_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "PARACT_CODSUB = '" & moddat_g_str_CodSub & "' AND "
   g_str_Parame = g_str_Parame & "PARACT_CODACT = " & moddat_g_str_CodMod & " AND "
   g_str_Parame = g_str_Parame & "PARACT_CODGRP = '" & moddat_g_str_CodGrp & "' AND "
   g_str_Parame = g_str_Parame & "PARACT_CODITE = '000' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   
   txt_CodGrp.Text = Trim(g_rst_Genera!PARACT_CODGRP)
   txt_NomGrp.Text = Trim(g_rst_Genera!PARACT_DESCRI)
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   Call fs_Activa_Editar(True)
   
   txt_CodGrp.Enabled = False
   Call gs_SetFocus(txt_NomGrp)
End Sub

Private Sub cmd_Grabar_Click()
   If moddat_g_int_FlgGrb = 1 Then
      txt_CodGrp.Text = Format(txt_CodGrp.Text, "000")
      
      If Len(Trim(txt_CodGrp.Text)) < 3 Then
         MsgBox "El Código de Grupo es de 3 dígitos.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_CodGrp)
         Exit Sub
      End If
   End If
   
   If Len(Trim(txt_NomGrp.Text)) = 0 Then
      MsgBox "Debe ingresar la Descripción del Grupo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomGrp)
      Exit Sub
   End If

   If moddat_g_int_FlgGrb = 1 Then
      'Validar que el registro no exista
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARACT WHERE "
      g_str_Parame = g_str_Parame & "PARACT_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
      g_str_Parame = g_str_Parame & "PARACT_CODSUB = '" & moddat_g_str_CodSub & "' AND "
      g_str_Parame = g_str_Parame & "PARACT_CODACT = " & moddat_g_str_CodMod & " AND "
      g_str_Parame = g_str_Parame & "PARACT_CODGRP = '" & txt_CodGrp.Text & "' AND "
      g_str_Parame = g_str_Parame & "PARACT_CODITE = '000' "
      
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
      
      g_str_Parame = "USP_CRE_PARACT ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodPrd & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodSub & "', "
      g_str_Parame = g_str_Parame & moddat_g_str_CodMod & ", "
      g_str_Parame = g_str_Parame & "'" & txt_CodGrp.Text & "', "
      g_str_Parame = g_str_Parame & "'000', "
      g_str_Parame = g_str_Parame & "'" & txt_NomGrp.Text & "', "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "1, "
      g_str_Parame = g_str_Parame & "0, "
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
   
   If moddat_g_int_FlgGrb = 2 Then
      Call fs_Buscar
      Call cmd_Cancel_Click
   Else
      Call fs_Buscar
      
      Call fs_Activa(False)
      Call fs_Activa_Editar(True)
      
      txt_CodGrp.Text = ""
      txt_NomGrp.Text = ""
      
      Call gs_SetFocus(txt_CodGrp)
   End If
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Imprim_Click()
   If MsgBox("¿Está seguro de imprimir el Reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   'Borrando Spool Local
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_PARACB"
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_PARADT"

   'Grabando en DAO (Cabecera)
   moddat_g_str_CadDAO = "SELECT * FROM RPT_PARACB WHERE PARACB_PRODUC = ' '"
   Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
   
   moddat_g_rst_RecDAO.AddNew
                        
   moddat_g_rst_RecDAO("PARACB_PRODUC") = pnl_Produc.Caption
   moddat_g_rst_RecDAO("PARACB_SUBPRD") = pnl_SubPrd.Caption
   moddat_g_rst_RecDAO("PARACB_ACTECO") = moddat_g_str_CodMod & " - " & moddat_g_str_DesMod
                        
   moddat_g_rst_RecDAO.Update
   moddat_g_rst_RecDAO.Close

   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARACT WHERE "
   g_str_Parame = g_str_Parame & "PARACT_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "PARACT_CODSUB = '" & moddat_g_str_CodSub & "' AND "
   g_str_Parame = g_str_Parame & "PARACT_CODACT = " & moddat_g_str_CodMod & " "
   g_str_Parame = g_str_Parame & "ORDER BY PARACT_CODGRP ASC, PARACT_CODITE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         'Grabando en DAO
         moddat_g_str_CadDAO = "SELECT * FROM RPT_PARADT WHERE PARADT_CODGRP = '" & g_rst_Princi!PARACT_CODGRP & "'"
         Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
         
         moddat_g_rst_RecDAO.AddNew
                              
         moddat_g_rst_RecDAO("PARADT_CODGRP") = Trim(g_rst_Princi!PARACT_CODGRP & "")
         moddat_g_rst_RecDAO("PARADT_CODITE") = IIf(Trim(g_rst_Princi!PARACT_CODITE & "") = "000", " ", Trim(g_rst_Princi!PARACT_CODITE & ""))
         moddat_g_rst_RecDAO("PARADT_DESCRI") = Trim(g_rst_Princi!PARACT_DESCRI & "")
         moddat_g_rst_RecDAO("PARADT_TIPPAR") = IIf(Trim(g_rst_Princi!PARACT_CODITE & "") = "000", " ", moddat_gf_Consulta_ParDes("036", CStr(g_rst_Princi!PARACT_TIPPAR)))
         
         If g_rst_Princi!PARACT_TIPPAR <> 3 Then
            moddat_g_rst_RecDAO("PARADT_TIPVAL") = IIf(Trim(g_rst_Princi!PARACT_CODITE & "") = "000", " ", moddat_gf_Consulta_ParDes("037", CStr(g_rst_Princi!PARACT_TIPVAL)))
         Else
            moddat_g_rst_RecDAO("PARADT_TIPVAL") = ""
         End If
         
         moddat_g_rst_RecDAO("PARADT_CANTID") = g_rst_Princi!PARACT_CANTID
         moddat_g_rst_RecDAO("PARADT_VALMIN") = g_rst_Princi!PARACT_VALMIN
         moddat_g_rst_RecDAO("PARADT_VALMAX") = g_rst_Princi!PARACT_VALMAX
         moddat_g_rst_RecDAO("PARADT_SITUAC") = ""
                              
         moddat_g_rst_RecDAO.Update
         moddat_g_rst_RecDAO.Close
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   Screen.MousePointer = 0

   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "PAR_PARACT_01.RPT"
   crp_Imprim.Action = 1
End Sub

Private Sub cmd_Limpia_Click()
   cmb_ActEco.ListIndex = -1
   
   txt_CodGrp.Text = ""
   txt_NomGrp.Text = ""
   
   Call gs_LimpiaGrid(grd_Listad)
   
   Call fs_Activa_Editar(False)
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_ActEco)
End Sub

Private Sub cmd_LisIte_Click()
   grd_Listad.Col = 0
   moddat_g_str_CodGrp = grd_Listad.Text
         
   grd_Listad.Col = 1
   moddat_g_str_DesGrp = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)
   
   frm_Produc_06.Show 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   pnl_Produc.Caption = moddat_g_str_CodPrd & " - " & moddat_g_str_NomPrd
   pnl_SubPrd.Caption = moddat_g_str_CodSub & " - " & moddat_g_str_DesSub
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call moddat_gs_Carga_LisIte_Combo(cmb_ActEco, 1, "008")
   Call cmd_Limpia_Click
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 1425
   grd_Listad.ColWidth(1) = 6130
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_ActEco.Enabled = p_Activa
   
   cmd_Buscar.Enabled = p_Activa
   
   grd_Listad.Enabled = Not p_Activa
   cmd_Agrega.Enabled = Not p_Activa
   cmd_Editar.Enabled = Not p_Activa
   cmd_Borrar.Enabled = Not p_Activa
   cmd_Imprim.Enabled = Not p_Activa
   cmd_LisIte.Enabled = Not p_Activa
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
   txt_CodGrp.Enabled = p_Activa
   txt_NomGrp.Enabled = p_Activa
   
   grd_Listad.Enabled = Not p_Activa
   cmd_Agrega.Enabled = Not p_Activa
   cmd_Editar.Enabled = Not p_Activa
   cmd_LisIte.Enabled = Not p_Activa
   cmd_Imprim.Enabled = Not p_Activa
   cmd_Borrar.Enabled = Not p_Activa
End Sub

Private Sub fs_Buscar()
   cmd_Agrega.Enabled = True
   cmd_Editar.Enabled = False
   cmd_LisIte.Enabled = False
   cmd_Borrar.Enabled = False
   cmd_Imprim.Enabled = False
   grd_Listad.Enabled = False
   
   moddat_g_str_CodMod = CStr(cmb_ActEco.ItemData(cmb_ActEco.ListIndex))
   moddat_g_str_DesMod = cmb_ActEco.Text
   
   Call gs_LimpiaGrid(grd_Listad)

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARACT WHERE "
   g_str_Parame = g_str_Parame & "PARACT_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "PARACT_CODSUB = '" & moddat_g_str_CodSub & "' AND "
   g_str_Parame = g_str_Parame & "PARACT_CODACT = " & moddat_g_str_CodMod & " AND "
   g_str_Parame = g_str_Parame & "PARACT_CODITE = '000' "
   g_str_Parame = g_str_Parame & "ORDER BY PARACT_CODGRP ASC "

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
      grd_Listad.Text = Trim(g_rst_Genera!PARACT_CODGRP)
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Genera!PARACT_DESCRI)
      
      g_rst_Genera.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   If grd_Listad.Rows > 0 Then
      cmd_Editar.Enabled = True
      cmd_LisIte.Enabled = True
      cmd_Borrar.Enabled = True
      cmd_Imprim.Enabled = True
      grd_Listad.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub SSPanel7_Click()
   Call gs_SorteaGrid(grd_Listad, 0, "C")
End Sub

Private Sub SSPanel8_Click()
   Call gs_SorteaGrid(grd_Listad, 1, "C")
End Sub

Private Sub txt_CodGrp_GotFocus()
   Call gs_SelecTodo(txt_CodGrp)
End Sub

Private Sub txt_CodGrp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NomGrp)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_NomGrp_GotFocus()
   Call gs_SelecTodo(txt_NomGrp)
End Sub

Private Sub txt_NomGrp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "()-_=/&%$#@ ?¿*")
   End If
End Sub


