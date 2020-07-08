VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_ParDes_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9945
   ClientLeft      =   6945
   ClientTop       =   2160
   ClientWidth     =   8205
   Icon            =   "PltPar_frm_002.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9945
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9945
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   8205
      _Version        =   65536
      _ExtentX        =   14473
      _ExtentY        =   17542
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
         Height          =   5895
         Left            =   30
         TabIndex        =   21
         Top             =   1560
         Width           =   8115
         _Version        =   65536
         _ExtentX        =   14314
         _ExtentY        =   10398
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
            TabIndex        =   22
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
            TabIndex        =   23
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
            Height          =   5475
            Left            =   30
            TabIndex        =   26
            Top             =   360
            Width           =   8025
            _ExtentX        =   14155
            _ExtentY        =   9657
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
         Left            =   30
         TabIndex        =   20
         Top             =   7500
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
            Picture         =   "PltPar_frm_002.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_LisIte 
            Height          =   675
            Left            =   7410
            Picture         =   "PltPar_frm_002.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Lista de Items"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   675
            Left            =   4650
            Picture         =   "PltPar_frm_002.frx":0890
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   675
            Left            =   5340
            Picture         =   "PltPar_frm_002.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   675
            Left            =   6030
            Picture         =   "PltPar_frm_002.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   675
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   60
            Top             =   60
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
         Height          =   765
         Left            =   30
         TabIndex        =   19
         Top             =   9120
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
         Begin VB.CommandButton cmd_Cancel 
            Height          =   675
            Left            =   7410
            Picture         =   "PltPar_frm_002.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   6720
            Picture         =   "PltPar_frm_002.frx":14B8
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   765
         Left            =   30
         TabIndex        =   17
         Top             =   750
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
         Begin VB.ComboBox cmb_TipGrp 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   210
            Width           =   3255
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   7350
            Picture         =   "PltPar_frm_002.frx":18FA
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   6660
            Picture         =   "PltPar_frm_002.frx":1D3C
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   5970
            Picture         =   "PltPar_frm_002.frx":2046
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Buscar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.Label Label1 
            Caption         =   "Seleccione Tipo de Grupo a mostrar:"
            Height          =   465
            Left            =   90
            TabIndex        =   18
            Top             =   180
            Width           =   1665
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   765
         Left            =   30
         TabIndex        =   14
         Top             =   8310
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
         Begin VB.TextBox txt_CodGrp 
            Height          =   315
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   60
            Width           =   825
         End
         Begin VB.TextBox txt_NomGrp 
            Height          =   315
            Left            =   1650
            MaxLength       =   80
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   390
            Width           =   6405
         End
         Begin VB.Label Label3 
            Caption         =   "Código Grupo:"
            Height          =   285
            Left            =   90
            TabIndex        =   16
            Top             =   90
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Descripción Grupo:"
            Height          =   285
            Left            =   90
            TabIndex        =   15
            Top             =   420
            Width           =   1455
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   675
         Left            =   30
         TabIndex        =   24
         Top             =   30
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
            Width           =   3105
            _Version        =   65536
            _ExtentX        =   5477
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Parámetros Descriptivos"
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
            Picture         =   "PltPar_frm_002.frx":2350
            Top             =   90
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_ParDes_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_TipGrp_Click()
   Call gs_SetFocus(cmd_Buscar)
End Sub

Private Sub cmb_TipGrp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipGrp_Click
   End If
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   
   Call fs_Activa_Editar(True)
   Call gs_SetFocus(txt_CodGrp)
End Sub

Private Sub cmd_Borrar_Click()
   grd_Listad.Col = 0
   moddat_g_str_Codigo = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)

   If MsgBox("¿Está seguro de eliminar el grupo?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   'Obteniendo Información del Registro
   g_str_Parame = "USP_BORRAR_MNT_PARDES_GRUPO (" & "'" & moddat_g_str_Codigo & "', "
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
   If cmb_TipGrp.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Grupo a mostar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipGrp)
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
   moddat_g_str_Codigo = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgGrb = 2
   
   
   'Obteniendo Información del Registro
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_PARDES WHERE "
   g_str_Parame = g_str_Parame & "PARDES_CODGRP = '" & moddat_g_str_Codigo & "' AND "
   g_str_Parame = g_str_Parame & "PARDES_CODITE = '000000' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If

   g_rst_Genera.MoveFirst
   
   txt_CodGrp.Text = Trim(g_rst_Genera!PARDES_CODGRP)
   txt_NomGrp.Text = Trim(g_rst_Genera!PARDES_DESCRI)
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   Call fs_Activa_Editar(True)
   
   txt_CodGrp.Enabled = False
   Call gs_SetFocus(txt_NomGrp)
End Sub

Private Sub cmd_Grabar_Click()
MsgBox ("entramos")

   If moddat_g_int_FlgGrb = 1 Then
      txt_CodGrp.Text = Format(txt_CodGrp.Text, "000")
      
      If Len(Trim(txt_CodGrp.Text)) < 3 Then
         MsgBox "El Código de Grupo es de 3 dígitos.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_CodGrp)
         Exit Sub
      End If
   
      If CInt(moddat_g_str_TipPar) = 1 Then
         If Not (CInt(txt_CodGrp.Text) >= 1 And CInt(txt_CodGrp.Text) <= 499) Then
            MsgBox "El Código de Grupo debe estar entre 001 y 499.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_CodGrp)
            Exit Sub
         End If
      Else
         If Not (CInt(txt_CodGrp.Text) >= 500 And CInt(txt_CodGrp.Text) <= 999) Then
            MsgBox "El Código de Grupo debe estar entre 500 y 999.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_CodGrp)
            Exit Sub
         End If
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
      g_str_Parame = g_str_Parame & "SELECT * FROM MNT_PARDES WHERE "
      g_str_Parame = g_str_Parame & "PARDES_CODGRP = '" & txt_CodGrp.Text & "' AND "
      g_str_Parame = g_str_Parame & "PARDES_CODITE = '000000' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
        g_rst_Genera.Close
        Set g_rst_Genera = Nothing
        
        MsgBox "El Código ya ha sido registrado. Por favor verifique el código e intente nuevamente.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(txt_CodGrp)
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
      
      If moddat_g_int_FlgGrb = 1 Then
         g_str_Parame = "USP_INSERTA_MNT_PARDES ("
         
         g_str_Parame = g_str_Parame & "'" & txt_CodGrp.Text & "', "
         g_str_Parame = g_str_Parame & "'000000', "
         g_str_Parame = g_str_Parame & "'" & txt_NomGrp.Text & "', "
         g_str_Parame = g_str_Parame & "1, "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      Else
         g_str_Parame = "USP_MODIFICA_MNT_PARDES ("
         
         g_str_Parame = g_str_Parame & "'" & txt_CodGrp.Text & "', "
         g_str_Parame = g_str_Parame & "'000000', "
         g_str_Parame = g_str_Parame & "'" & txt_NomGrp.Text & "', "
         g_str_Parame = g_str_Parame & "1, "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      End If
      
      MsgBox (g_str_Parame)
      
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
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_PARADT"
                        
   If CInt(moddat_g_str_TipPar) = 1 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM MNT_PARDES WHERE "
      g_str_Parame = g_str_Parame & "PARDES_SITUAC = '1' AND "
      g_str_Parame = g_str_Parame & "PARDES_CODGRP >= '001' AND "
      g_str_Parame = g_str_Parame & "PARDES_CODGRP <= '499' "
      g_str_Parame = g_str_Parame & "ORDER BY PARDES_CODGRP ASC, PARDES_CODITE ASC "
   Else
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM MNT_PARDES WHERE "
      g_str_Parame = g_str_Parame & "PARDES_SITUAC = '1' AND "
      g_str_Parame = g_str_Parame & "PARDES_CODGRP >= '500' AND "
      g_str_Parame = g_str_Parame & "PARDES_CODGRP <= '999' "
      g_str_Parame = g_str_Parame & "ORDER BY PARDES_CODGRP ASC, PARDES_CODITE ASC "
   End If

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         'Grabando en DAO
         moddat_g_str_CadDAO = "SELECT * FROM RPT_PARADT WHERE PARADT_CODGRP = '" & g_rst_Princi!PARDES_CODGRP & "'"
         Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
         
         moddat_g_rst_RecDAO.AddNew
                              
         moddat_g_rst_RecDAO("PARADT_CODGRP") = Trim(g_rst_Princi!PARDES_CODGRP & "")
         moddat_g_rst_RecDAO("PARADT_CODITE") = IIf(Trim(g_rst_Princi!PARDES_CODITE & "") = "000000", " ", Trim(g_rst_Princi!PARDES_CODITE & ""))
         moddat_g_rst_RecDAO("PARADT_DESCRI") = Trim(g_rst_Princi!PARDES_DESCRI & "")
         moddat_g_rst_RecDAO("PARADT_TIPPAR") = ""
         moddat_g_rst_RecDAO("PARADT_TIPVAL") = ""
         moddat_g_rst_RecDAO("PARADT_CANTID") = 0
         moddat_g_rst_RecDAO("PARADT_VALMIN") = 0
         moddat_g_rst_RecDAO("PARADT_VALMAX") = 0
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

   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "PAR_PARDES_01.RPT"
   crp_Imprim.Action = 1
End Sub

Private Sub cmd_Limpia_Click()
   cmb_TipGrp.ListIndex = -1
   
   txt_CodGrp.Text = ""
   txt_NomGrp.Text = ""
   
   Call gs_LimpiaGrid(grd_Listad)
   
   Call fs_Activa_Editar(False)
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_TipGrp)
End Sub

Private Sub cmd_LisIte_Click()
   grd_Listad.Col = 0
   moddat_g_str_Codigo = grd_Listad.Text
         
   grd_Listad.Col = 1
   moddat_g_str_Descri = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)
   
   frm_ParDes_02.Show 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call modsis_gs_Carga_TipGrp(cmb_TipGrp)
   
   Call cmd_Limpia_Click
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 1440
   grd_Listad.ColWidth(1) = 6200
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_TipGrp.Enabled = p_Activa
   
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
   cmd_Imprim.Enabled = Not p_Activa
   cmd_LisIte.Enabled = Not p_Activa
   cmd_Borrar.Enabled = Not p_Activa
End Sub

Private Sub fs_Buscar()
   cmd_Agrega.Enabled = True
   cmd_Editar.Enabled = False
   cmd_LisIte.Enabled = False
   cmd_Borrar.Enabled = False
   cmd_Imprim.Enabled = False
   grd_Listad.Enabled = False
   
   Call gs_LimpiaGrid(grd_Listad)
   
   moddat_g_str_TipPar = cmb_TipGrp.ItemData(cmb_TipGrp.ListIndex)
   
   If CInt(moddat_g_str_TipPar) = 1 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM MNT_PARDES WHERE "
      g_str_Parame = g_str_Parame & "PARDES_CODITE = '000000' AND "
      g_str_Parame = g_str_Parame & "PARDES_SITUAC = '1' AND "
      g_str_Parame = g_str_Parame & "PARDES_CODGRP >= '001' AND "
      g_str_Parame = g_str_Parame & "PARDES_CODGRP <= '499' "
      g_str_Parame = g_str_Parame & "ORDER BY PARDES_CODGRP ASC "
   Else
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM MNT_PARDES WHERE "
      g_str_Parame = g_str_Parame & "PARDES_CODITE = '000000' AND "
      g_str_Parame = g_str_Parame & "PARDES_SITUAC = '1' AND "
      g_str_Parame = g_str_Parame & "PARDES_CODGRP >= '500' AND "
      g_str_Parame = g_str_Parame & "PARDES_CODGRP <= '999' "
      g_str_Parame = g_str_Parame & "ORDER BY PARDES_CODGRP ASC "
   End If

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   'If g_rst_Genera.MaxRecords = 0 Then
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
      grd_Listad.Text = Trim(g_rst_Genera!PARDES_CODGRP)
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Genera!PARDES_DESCRI)
      
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

