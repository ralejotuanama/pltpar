VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Comviv_1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frm_ComViv_1"
   ClientHeight    =   4830
   ClientLeft      =   4695
   ClientTop       =   4650
   ClientWidth     =   9285
   Icon            =   "PltPar_frm_039.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel4 
      Height          =   4785
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9285
      _Version        =   65536
      _ExtentX        =   16378
      _ExtentY        =   8440
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
         Height          =   675
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   9195
         _Version        =   65536
         _ExtentX        =   16219
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
         Begin Threed.SSPanel SSPanel3 
            Height          =   480
            Left            =   600
            TabIndex        =   10
            Top             =   90
            Width           =   4785
            _Version        =   65536
            _ExtentX        =   8440
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Mantenimiento Comisiones Mivivienda "
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
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   8760
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
         Begin VB.Image Image1 
            Height          =   480
            Left            =   90
            Picture         =   "PltPar_frm_039.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   735
         Left            =   30
         TabIndex        =   11
         Top             =   720
         Width           =   9195
         _Version        =   65536
         _ExtentX        =   16219
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
         Begin VB.CommandButton cmd_Imprim 
            Height          =   675
            Left            =   3480
            Picture         =   "PltPar_frm_039.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Imprimir"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   675
            Left            =   2790
            Picture         =   "PltPar_frm_039.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   675
            Left            =   2100
            Picture         =   "PltPar_frm_039.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   675
            Left            =   1410
            Picture         =   "PltPar_frm_039.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   30
            Picture         =   "PltPar_frm_039.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Buscar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   720
            Picture         =   "PltPar_frm_039.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salir 
            Height          =   675
            Left            =   8460
            Picture         =   "PltPar_frm_039.frx":168A
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   2715
         Left            =   30
         TabIndex        =   12
         Top             =   2040
         Width           =   9195
         _Version        =   65536
         _ExtentX        =   16219
         _ExtentY        =   4789
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   285
            Left            =   4680
            TabIndex        =   13
            Top             =   60
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Plazo Inicial"
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   285
            Left            =   2670
            TabIndex        =   14
            Top             =   60
            Width           =   2025
            _Version        =   65536
            _ExtentX        =   3572
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo de Moneda"
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
            Height          =   2325
            Left            =   60
            TabIndex        =   15
            Top             =   360
            Width           =   9165
            _ExtentX        =   16166
            _ExtentY        =   4101
            _Version        =   393216
            Rows            =   12
            Cols            =   7
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   5820
            TabIndex        =   16
            Top             =   60
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Plazo Final"
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
         Begin Threed.SSPanel SSPanel12 
            Height          =   285
            Left            =   6960
            TabIndex        =   17
            Top             =   60
            Width           =   1875
            _Version        =   65536
            _ExtentX        =   3307
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Porcentaje de Comisión"
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
         Begin Threed.SSPanel SSPanel13 
            Height          =   285
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   2625
            _Version        =   65536
            _ExtentX        =   4630
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo de Comisión"
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
      Begin Threed.SSPanel SSPanel1 
         Height          =   555
         Left            =   30
         TabIndex        =   19
         Top             =   1470
         Width           =   9195
         _Version        =   65536
         _ExtentX        =   16219
         _ExtentY        =   979
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
         Begin VB.ComboBox cmb_CodPro 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   120
            Width           =   7305
         End
         Begin VB.Label Label8 
            Caption         =   "Código de Producto:"
            Height          =   285
            Left            =   120
            TabIndex        =   20
            Top             =   150
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "frm_Comviv_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Produc()      As moddat_tpo_Genera

Private Sub cmb_CodPro_KeyPress(KeyAscii As Integer)

   'Se llama al siguiente control, luego de presionar Enter
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If

End Sub

Private Sub cmd_Agrega_Click()

   modvar_g_int_TipPan = 1
   
   'Se llama al formulario2
   frm_Comviv_2.Show 1
   Call fs_Buscar
   
End Sub

Private Sub cmd_Borrar_Click()
    
   'Tipo de Comisión
   grd_Listad.Col = 5
   modvar_g_int_TipCom = grd_Listad.Text

   'Tipo de Moneda
   grd_Listad.Col = 6
   modvar_g_int_TipMon = grd_Listad.Text

   'Plazo Inicial
   grd_Listad.Col = 2
   modvar_g_int_PlaIni = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)
   
   'Se envia mensaje con pregunta de eliminacion del campo correspondiente
   If MsgBox("¿Está seguro de eliminar la Comision?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Puntero reloj de arena
   Screen.MousePointer = 11
   
   'Obteniendo Información del Registro
   g_str_Parame = "USP_MNT_COMMVI_BORRAR (" & "'" & moddat_g_str_CodPrd & "', "
   g_str_Parame = g_str_Parame & "'" & modvar_g_int_TipCom & "', "
   g_str_Parame = g_str_Parame & "'" & modvar_g_int_TipMon & "', "
   g_str_Parame = g_str_Parame & "'" & modvar_g_int_PlaIni & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "

   'Pregunta si no se ejecuta la sentencia y sale del metodo
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
       Exit Sub
   End If
   
   'Se llama al metodo fs_Buscar
   Call fs_Buscar
   
   'Reloj Normal
   Screen.MousePointer = 0
   
End Sub

Private Sub cmd_Buscar_Click()

   'Validacion de combo si es que no se eligio ningun producto
   If cmb_CodPro.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodPro)
      Exit Sub
   End If
      
   'Se llama las variables globales y se envia los controles correspondientes
   moddat_g_str_NomPrd = cmb_CodPro.Text
   moddat_g_str_CodPrd = l_arr_Produc(cmb_CodPro.ListIndex + 1).Genera_Codigo
   
   Call fs_Buscar
   
   'Se desabilita el combo codigo y el boton buscar
   cmd_Buscar.Enabled = False
   cmb_CodPro.Enabled = False
   
End Sub

Private Sub cmd_Editar_Click()

   'Tipo de Comision
   grd_Listad.Col = 5
   modvar_g_int_TipCom = grd_Listad.Text

   'Tipo de Moneda
   grd_Listad.Col = 6
   modvar_g_int_TipMon = grd_Listad.Text

   'Plazo Inicial
   grd_Listad.Col = 2
   modvar_g_int_PlaIni = grd_Listad.Text
   
   'Plazo Final
   grd_Listad.Col = 3
   modvar_g_int_PlaFin = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)

   modvar_g_int_TipPan = 2
   frm_Comviv_2.Show 1
   
   Call fs_Buscar
   
End Sub

Private Sub cmd_Imprim_Click()

   'Se envia pregunta si se desea imprimir el reporte
   If MsgBox("¿Está seguro de Imprimir el reporte?.", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Se conecta al crystal report
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   'Se envia las tablas correspondientes en el orden que fueron utilizadas
   crp_Imprim.DataFiles(0) = "OPE_COMMVI"
   crp_Imprim.DataFiles(1) = "CRE_PRODUC"
   
   'Se selecciona la formula con el tipo de producto
   crp_Imprim.SelectionFormula = "{OPE_COMMVI.COMMVI_CODPRD} = '" & l_arr_Produc(cmb_CodPro.ListIndex + 1).Genera_Codigo & "'"
   
   'Se pone la llamada del nombre del reporte y se escoge donde se destinara el reporte
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_COMMIV_01.RPT"
   crp_Imprim.Destination = crptToWindow
   
   crp_Imprim.Action = 1
   
End Sub

Private Sub cmd_Limpia_Click()
   
   'Se hace llamado a los metodos fs_habilitarbotones, gs_limpiaGrid, gs_setFocus
   Call fs_HabilitarBotones(False)
   Call gs_LimpiaGrid(grd_Listad)
   Call gs_SetFocus(cmb_CodPro)
   
   cmb_CodPro.ListIndex = -1
   
   'Se habilitan el combo del producto y el boton buscar
   cmd_Buscar.Enabled = True
   cmb_CodPro.Enabled = True
   
End Sub

Private Sub cmd_Salir_Click()

   'Se cierra el formulario
   Unload Me

End Sub

Private Sub Form_Load()

   Call fs_Inicia
   Me.Caption = modgen_g_str_NomPlt
   
   'Se Habilitan los botones
   Call fs_HabilitarBotones(False)
   
   'Se escoge la seleccion y se centra el Formulario
   Call gs_SetFocus(cmb_CodPro)
   Call gs_CentraForm(Me)
      
End Sub
Private Sub fs_Inicia()

   'Llenando Lista de Productos
   Call moddat_gs_Carga_Produc(cmb_CodPro, l_arr_Produc, 4)
   
   'Inicializando Columnas de Grid
   'Se le da el Ancho a las Columnas
   grd_Listad.ColWidth(0) = 2600
   grd_Listad.ColWidth(1) = 2000
   grd_Listad.ColWidth(2) = 1135
   grd_Listad.ColWidth(3) = 1135
   grd_Listad.ColWidth(4) = 1855
   'Estos dos campos son para poder coger el codigo
   grd_Listad.ColWidth(5) = 0
   grd_Listad.ColWidth(6) = 0
      
   'Se da el Alineamiento a las columnas
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignRightCenter
   
   Call gs_LimpiaGrid(grd_Listad)
   
End Sub

Private Sub fs_Buscar()
   
   'Se realiza la llamada del combo en la BD
   g_str_Parame = "SELECT * FROM OPE_COMMVI "
   g_str_Parame = g_str_Parame & "WHERE COMMVI_CODPRD = '" & l_arr_Produc(cmb_CodPro.ListIndex + 1).Genera_Codigo & "' "
      
   Call gs_LimpiaGrid(grd_Listad)
   
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   'Se avalua si existe data en la Grilla
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se han encontrado Solicitudes para esa selección.", vbExclamation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      'Se desabilitan y habilitan los botones
      cmd_Borrar.Enabled = False
      cmd_Editar.Enabled = False
      cmd_Agrega.Enabled = True
      cmd_Imprim.Enabled = False
      Exit Sub
   End If
   
   'Reloj de arena
   Screen.MousePointer = 11
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   
   'Mientras no sea el final del archivo se van agregando los campos a la grilla
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = moddat_gf_Consulta_ParDes("029", CStr(g_rst_Princi!COMMVI_TIPCOM))
      
      grd_Listad.Col = 1
      grd_Listad.Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!COMMVI_TIPMON))
               
      grd_Listad.Col = 2
      grd_Listad.Text = g_rst_Princi!COMMVI_PLAINI

      grd_Listad.Col = 3
      grd_Listad.Text = g_rst_Princi!COMMVI_PLAFIN
      
      grd_Listad.Col = 4
      grd_Listad.Text = g_rst_Princi!COMMVI_PORCEN
      
      grd_Listad.Col = 5
      grd_Listad.Text = CStr(g_rst_Princi!COMMVI_TIPCOM)
               
      grd_Listad.Col = 6
      grd_Listad.Text = CStr(g_rst_Princi!COMMVI_TIPMON)
               
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   'Cerramos la conexion a la BD
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
   Call fs_HabilitarBotones(True)
   
   'Reloj Normal
   Screen.MousePointer = 0
   
End Sub

Private Sub grd_Listad_SelChange()

   'Seleccion de Filas
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
   
End Sub

Private Sub SSPanel10_Click()

   'Se ordena de forma ascendente y descendente
   If Len(Trim(SSPanel10.Tag)) = 0 Or SSPanel10.Tag = "D" Then
      SSPanel10.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      SSPanel10.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
   
End Sub

Private Sub SSPanel11_Click()
   
   'Se ordena de forma ascendente y descendente
   If Len(Trim(SSPanel11.Tag)) = 0 Or SSPanel11.Tag = "D" Then
      SSPanel11.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 3, "N")
   Else
      SSPanel11.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 3, "N-")
   End If
   
End Sub

Private Sub SSPanel12_Click()

   'Se ordena de forma ascendente y descendente
   If Len(Trim(SSPanel12.Tag)) = 0 Or SSPanel12.Tag = "D" Then
      SSPanel12.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 4, "N")
   Else
      SSPanel12.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 4, "N-")
   End If
   
End Sub

Private Sub SSPanel13_Click()

   'Se ordena de forma ascendente y descendente
   If Len(Trim(SSPanel13.Tag)) = 0 Or SSPanel13.Tag = "D" Then
      SSPanel13.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      SSPanel13.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If

End Sub

Private Sub SSPanel7_Click()

   'Se ordena de forma ascendente y descendente
   If Len(Trim(SSPanel7.Tag)) = 0 Or SSPanel7.Tag = "D" Then
      SSPanel7.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "N")
   Else
      SSPanel7.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "N-")
   End If
   
End Sub

Public Sub fs_HabilitarBotones(r_boo_estado As Boolean)
   
   'Se Habilitan o desabilitan los botones dependiente del estado
   cmd_Borrar.Enabled = r_boo_estado
   cmd_Editar.Enabled = r_boo_estado
   cmd_Agrega.Enabled = r_boo_estado
   cmd_Imprim.Enabled = r_boo_estado

End Sub
