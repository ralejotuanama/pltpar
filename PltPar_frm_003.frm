VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_ParDes_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10080
   ClientLeft      =   9300
   ClientTop       =   615
   ClientWidth     =   8235
   Icon            =   "PltPar_frm_003.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10065
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8235
      _Version        =   65536
      _ExtentX        =   14526
      _ExtentY        =   17754
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
         Height          =   6255
         Left            =   30
         TabIndex        =   19
         Top             =   1320
         Width           =   8115
         _Version        =   65536
         _ExtentX        =   14314
         _ExtentY        =   11033
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
            TabIndex        =   20
            Top             =   60
            Width           =   6315
            _Version        =   65536
            _ExtentX        =   11139
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Descripción Item"
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
            TabIndex        =   21
            Top             =   60
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Código Item"
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
            Height          =   5865
            Left            =   30
            TabIndex        =   0
            Top             =   360
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   10345
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
         TabIndex        =   18
         Top             =   7620
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
            Picture         =   "PltPar_frm_003.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   675
            Left            =   5340
            Picture         =   "PltPar_frm_003.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   675
            Left            =   4650
            Picture         =   "PltPar_frm_003.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   675
            Left            =   6030
            Picture         =   "PltPar_frm_003.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   7410
            Picture         =   "PltPar_frm_003.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   0
            Top             =   30
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Presentación Preliminar"
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowNavigationCtls=   -1  'True
            WindowShowCancelBtn=   -1  'True
            WindowShowPrintBtn=   -1  'True
            WindowShowExportBtn=   -1  'True
            WindowShowZoomCtl=   -1  'True
            WindowShowProgressCtls=   -1  'True
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   765
         Left            =   30
         TabIndex        =   17
         Top             =   9240
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   6720
            Picture         =   "PltPar_frm_003.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Cancel 
            Height          =   675
            Left            =   7410
            Picture         =   "PltPar_frm_003.frx":15F0
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   525
         Left            =   30
         TabIndex        =   11
         Top             =   750
         Width           =   8115
         _Version        =   65536
         _ExtentX        =   14314
         _ExtentY        =   926
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
         Begin Threed.SSPanel pnl_Grupo 
            Height          =   405
            Left            =   990
            TabIndex        =   13
            Top             =   60
            Width           =   7065
            _Version        =   65536
            _ExtentX        =   12462
            _ExtentY        =   714
            _StockProps     =   15
            Caption         =   "SSPanel3"
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
            Alignment       =   1
         End
         Begin VB.Label Label1 
            Caption         =   "Grupo:"
            Height          =   315
            Left            =   90
            TabIndex        =   12
            Top             =   120
            Width           =   915
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   765
         Left            =   30
         TabIndex        =   14
         Top             =   8430
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
         Begin VB.TextBox txt_Nombre 
            Height          =   315
            Left            =   1470
            MaxLength       =   80
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   390
            Width           =   6585
         End
         Begin VB.TextBox txt_Codigo 
            Height          =   315
            Left            =   1470
            MaxLength       =   6
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   60
            Width           =   845
         End
         Begin VB.Label Label4 
            Caption         =   "Descripción Item:"
            Height          =   285
            Left            =   90
            TabIndex        =   16
            Top             =   420
            Width           =   1305
         End
         Begin VB.Label Label3 
            Caption         =   "Código Item:"
            Height          =   285
            Left            =   90
            TabIndex        =   15
            Top             =   90
            Width           =   1095
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   675
         Left            =   30
         TabIndex        =   22
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
            TabIndex        =   23
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
            Picture         =   "PltPar_frm_003.frx":18FA
            Top             =   90
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_ParDes_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   
   Call fs_Activa(True)
   Call gs_SetFocus(txt_Codigo)
End Sub

Private Sub cmd_Borrar_Click()
   grd_Listad.Col = 0
   moddat_g_str_CodIte = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)
   
   If MsgBox("¿Está seguro de eliminar el Item?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11

   g_str_Parame = "USP_BORRAR_MNT_PARDES_ITEM (" & "'" & moddat_g_str_Codigo & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodIte & "', "
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
   txt_Codigo.Text = ""
   txt_Nombre.Text = ""
   
   Call fs_Activa(False)
   Call gs_SetFocus(grd_Listad)
   
   If grd_Listad.Rows = 0 Then
      cmd_Agrega.Enabled = True
      cmd_Editar.Enabled = False
      cmd_Borrar.Enabled = False
      cmd_Imprim.Enabled = False
      grd_Listad.Enabled = False
   End If
End Sub

Private Sub cmd_Editar_Click()
   grd_Listad.Col = 0
   moddat_g_str_CodIte = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgGrb = 2
   
   'Obteniendo Información del Registro
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_PARDES WHERE "
   g_str_Parame = g_str_Parame & "PARDES_CODGRP = '" & moddat_g_str_Codigo & "' AND "
   g_str_Parame = g_str_Parame & "PARDES_CODITE = '" & moddat_g_str_CodIte & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   
   txt_Codigo.Text = Trim(g_rst_Genera!PARDES_CODITE)
   txt_Nombre.Text = Trim(g_rst_Genera!PARDES_DESCRI)
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   Call fs_Activa(True)
   
   txt_Codigo.Enabled = False
   Call gs_SetFocus(txt_Nombre)
End Sub

Private Sub cmd_Grabar_Click()
   If moddat_g_int_FlgGrb = 1 Then
      txt_Codigo.Text = Format(txt_Codigo.Text, "000000")
      
      If Len(Trim(txt_Codigo.Text)) = 0 Then
         MsgBox "Debe ingresar el Código de Item.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Codigo)
         Exit Sub
      End If
   End If
   
   If Len(Trim(txt_Nombre.Text)) = 0 Then
      MsgBox "Debe ingresar la Descripción del Item.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Nombre)
      Exit Sub
   End If

   If moddat_g_int_FlgGrb = 1 Then
      'Validar que el registro no exista
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM MNT_PARDES WHERE "
      g_str_Parame = g_str_Parame & "PARDES_CODGRP = '" & moddat_g_str_Codigo & "' AND "
      g_str_Parame = g_str_Parame & "PARDES_CODITE = '" & txt_Codigo.Text & "'"
   
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
      
      If moddat_g_int_FlgGrb = 1 Then
         g_str_Parame = "USP_INSERTA_MNT_PARDES ("
         
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_Codigo & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Codigo.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Nombre.Text & "', "
         g_str_Parame = g_str_Parame & "1, "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      Else
         g_str_Parame = "USP_MODIFICA_MNT_PARDES ("
         
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_Codigo & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Codigo.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Nombre.Text & "', "
         g_str_Parame = g_str_Parame & "1, "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      End If
      
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
      
      Call fs_Activa(True)
      
      txt_Codigo.Text = ""
      txt_Nombre.Text = ""
      Call gs_SetFocus(txt_Codigo)
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
                        
   moddat_g_rst_RecDAO("PARACB_CODGRP") = moddat_g_str_Codigo
   moddat_g_rst_RecDAO("PARACB_DESGRP") = moddat_g_str_Descri
                        
   moddat_g_rst_RecDAO.Update
   moddat_g_rst_RecDAO.Close
                        

   'Generando Detalle
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_PARDES WHERE "
   g_str_Parame = g_str_Parame & "PARDES_CODGRP = '" & moddat_g_str_Codigo & "' AND "
   g_str_Parame = g_str_Parame & "PARDES_CODITE <> '000000' AND "
   g_str_Parame = g_str_Parame & "PARDES_SITUAC = '1' "
   g_str_Parame = g_str_Parame & "ORDER BY PARDES_CODITE ASC "

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

   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "PAR_PARDES_02.RPT"
   crp_Imprim.Action = 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_Grupo.Caption = moddat_g_str_Codigo & " - " & moddat_g_str_Descri
   
   Call fs_Inicia
   Call fs_Activa(False)
   
   txt_Codigo.Text = ""
   txt_Nombre.Text = ""
   
   Call fs_Buscar
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 1440
   grd_Listad.ColWidth(1) = 6300
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
End Sub

Private Sub fs_Buscar()
   cmd_Agrega.Enabled = True
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   cmd_Imprim.Enabled = False
   
   grd_Listad.Enabled = False
   
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_PARDES WHERE "
   g_str_Parame = g_str_Parame & "PARDES_CODGRP = '" & moddat_g_str_Codigo & "' AND "
   g_str_Parame = g_str_Parame & "PARDES_CODITE <> '000000' AND "
   g_str_Parame = g_str_Parame & "PARDES_SITUAC = '1' "
   g_str_Parame = g_str_Parame & "ORDER BY PARDES_CODITE ASC "

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
      grd_Listad.Text = g_rst_Genera!PARDES_CODITE
      
      grd_Listad.Col = 1
      grd_Listad.Text = g_rst_Genera!PARDES_DESCRI
      
      g_rst_Genera.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   If grd_Listad.Rows > 0 Then
      cmd_Editar.Enabled = True
      cmd_Borrar.Enabled = True
      cmd_Imprim.Enabled = True
      
      grd_Listad.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub fs_Activa(p_Activa As Integer)
   cmd_Grabar.Enabled = p_Activa
   cmd_Cancel.Enabled = p_Activa
   txt_Codigo.Enabled = p_Activa
   txt_Nombre.Enabled = p_Activa
   
   grd_Listad.Enabled = Not p_Activa
   cmd_Agrega.Enabled = Not p_Activa
   cmd_Editar.Enabled = Not p_Activa
   cmd_Borrar.Enabled = Not p_Activa
   cmd_Imprim.Enabled = Not p_Activa
End Sub

Private Sub SSPanel7_Click()
   Call gs_SorteaGrid(grd_Listad, 0, "C")
End Sub

Private Sub SSPanel8_Click()
   Call gs_SorteaGrid(grd_Listad, 1, "C")
End Sub

Private Sub txt_Codigo_GotFocus()
   Call gs_SelecTodo(txt_Codigo)
End Sub

Private Sub txt_Codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Nombre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Nombre_GotFocus()
   Call gs_SelecTodo(txt_Nombre)
End Sub

Private Sub txt_Nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "<>()-_=/&%$,.;:#@ ?¿*")
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

