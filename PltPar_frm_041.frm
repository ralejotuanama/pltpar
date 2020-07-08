VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_Feriad_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   4200
   ClientLeft      =   2400
   ClientTop       =   3645
   ClientWidth     =   5940
   Icon            =   "PltPar_frm_041.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   4245
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5955
      _Version        =   65536
      _ExtentX        =   10504
      _ExtentY        =   7488
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
         TabIndex        =   5
         Top             =   30
         Width           =   5895
         _Version        =   65536
         _ExtentX        =   10398
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
            TabIndex        =   6
            Top             =   120
            Width           =   2685
            _Version        =   65536
            _ExtentX        =   4736
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Mantenimiento Feriados"
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
            Picture         =   "PltPar_frm_041.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   735
         Left            =   30
         TabIndex        =   7
         Top             =   720
         Width           =   5895
         _Version        =   65536
         _ExtentX        =   10398
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
         Begin VB.CommandButton cmd_Salir 
            Height          =   675
            Left            =   5190
            Picture         =   "PltPar_frm_041.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   675
            Left            =   30
            Picture         =   "PltPar_frm_041.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   675
            Left            =   720
            Picture         =   "PltPar_frm_041.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   2715
         Left            =   30
         TabIndex        =   8
         Top             =   1470
         Width           =   5895
         _Version        =   65536
         _ExtentX        =   10398
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   285
            Left            =   1290
            TabIndex        =   9
            Top             =   60
            Width           =   4245
            _Version        =   65536
            _ExtentX        =   7488
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   2325
            Left            =   30
            TabIndex        =   2
            Top             =   360
            Width           =   5835
            _ExtentX        =   10292
            _ExtentY        =   4101
            _Version        =   393216
            Rows            =   12
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel13 
            Height          =   285
            Left            =   60
            TabIndex        =   10
            Top             =   60
            Width           =   1245
            _Version        =   65536
            _ExtentX        =   2196
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Dia Feriado"
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
   End
End
Attribute VB_Name = "frm_Feriad_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Agrega_Click()

   'Se hace el llamado a el formulario y al metodo fs_GridInicia
   frm_Feriad_02.Show 1
   Call fs_GridInicia
   
End Sub

Private Sub cmd_Borrar_Click()

   'Se Refresca la grilla
   Call gs_RefrescaGrid(grd_Listad)
   
   'Se muestra ventana de pregunta si se desea eliminar el campo correspondiente
   If MsgBox("¿Está seguro de eliminar la Fecha indicada?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Puntero en reloj de arena
   Screen.MousePointer = 11
   grd_Listad.Col = 0
   modgen_g_int_DiaFer = grd_Listad.Text
      
   'Obteniendo Información del Registro y se procede al borrado del campo seleccionado
   g_str_Parame = "USP_OPE_DIAFER_BORRAR (" & modgen_g_int_DiaFer & ") "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
       Exit Sub
   End If
      
   'Puntero Normal
   Screen.MousePointer = 0
   
   'Se llama al metodo fs_GridInicia
   Call fs_GridInicia

End Sub

Private Sub cmd_Salir_Click()
   'Cerramos el formulario
   Unload Me

End Sub

Private Sub Form_Load()

   'Enviamos el nombre de la plataforma
   Me.Caption = modgen_g_str_NomPlt
   
   'Llamamos a los metodos fs_Inicia, fs_GridInicia, gs_CentraForm
   Call fs_Inicia
   Call fs_GridInicia
   Call gs_CentraForm(Me)
   
End Sub
Private Sub fs_Inicia()

   'Inicializando Columnas de Grid
   'Se le da el Ancho a las Columnas
   grd_Listad.ColWidth(0) = 1245
   grd_Listad.ColWidth(1) = 4245
   
   'Se da el Alineamiento a las columnas
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
      
   Call gs_LimpiaGrid(grd_Listad)
   
End Sub
Private Sub fs_GridInicia()
   'Se realiza la llamada del combo en la BD
   g_str_Parame = "SELECT * FROM OPE_DIAFER ORDER BY DIAFER_DIAFER ASC"
        
   Call gs_LimpiaGrid(grd_Listad)
      
   'Se pregunta si no se ejecuto, de ser afirmativo sale del metodo
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   'Se avalua si existe data en la Grilla
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
   
      MsgBox "No se han encontrado Solicitudes para esa selección.", vbExclamation, modgen_g_str_NomPlt
      
      'Se cierra la conexion a la BD
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   'Puntero en reloj de Arena
   Screen.MousePointer = 11
   
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   
   'Se hace una condicion repetitiva y mientras no se se el final del registro se ejecuta las sentencias
   Do While Not g_rst_Princi.EOF
   
      'Se va llenando los campos correspondientes a la grilla
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
                          
      'Primera Columna: Dia Feriado
      grd_Listad.Col = 0
      grd_Listad.Text = g_rst_Princi!DIAFER_DIAFER

      'Segunda Columna: Descripcion
      grd_Listad.Col = 1
      grd_Listad.Text = g_rst_Princi!DIAFER_DESFER
                    
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   'Se cierra la conexion a la BD
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   Call gs_UbiIniGrid(grd_Listad)
    
   'Puntero Normal
   Screen.MousePointer = 0

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

Private Sub SSPanel13_Click()

   'Se ordena de forma ascendente y descendente
   If Len(Trim(SSPanel13.Tag)) = 0 Or SSPanel13.Tag = "D" Then
      SSPanel13.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "N")
   Else
      SSPanel13.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "N-")
   End If
   
End Sub
