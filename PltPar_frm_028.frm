VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_PosUbi_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   7305
   ClientLeft      =   1095
   ClientTop       =   1620
   ClientWidth     =   12915
   Icon            =   "PltPar_frm_028.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   12915
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7305
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   12915
      _Version        =   65536
      _ExtentX        =   22781
      _ExtentY        =   12885
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
         Height          =   765
         Left            =   30
         TabIndex        =   13
         Top             =   3870
         Width           =   12825
         _Version        =   65536
         _ExtentX        =   22622
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
         Begin VB.CommandButton cmd_Agrega 
            Height          =   675
            Left            =   10020
            Picture         =   "PltPar_frm_028.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   675
            Left            =   10710
            Picture         =   "PltPar_frm_028.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   12090
            Picture         =   "PltPar_frm_028.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   675
            Left            =   11400
            Picture         =   "PltPar_frm_028.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   765
         Left            =   30
         TabIndex        =   14
         Top             =   6480
         Width           =   12825
         _Version        =   65536
         _ExtentX        =   22622
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
            Left            =   12090
            Picture         =   "PltPar_frm_028.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   11370
            Picture         =   "PltPar_frm_028.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   3075
         Left            =   30
         TabIndex        =   15
         Top             =   750
         Width           =   12825
         _Version        =   65536
         _ExtentX        =   22622
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   2685
            Left            =   60
            TabIndex        =   0
            Top             =   330
            Width           =   12705
            _ExtentX        =   22410
            _ExtentY        =   4736
            _Version        =   393216
            Rows            =   12
            Cols            =   5
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   4260
            TabIndex        =   16
            Top             =   60
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Departamento"
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
            TabIndex        =   17
            Top             =   60
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Código de Zona"
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
         Begin Threed.SSPanel SSPanel6 
            Height          =   285
            Left            =   6990
            TabIndex        =   21
            Top             =   60
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Provincia"
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
         Begin Threed.SSPanel SSPanel9 
            Height          =   285
            Left            =   9720
            TabIndex        =   22
            Top             =   60
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Distrito"
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
            Left            =   1530
            TabIndex        =   23
            Top             =   60
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nombre Zona"
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
         Height          =   1755
         Left            =   30
         TabIndex        =   18
         Top             =   4680
         Width           =   12825
         _Version        =   65536
         _ExtentX        =   22622
         _ExtentY        =   3096
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
         Begin VB.TextBox txt_Codigo 
            Height          =   315
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   1050
            Width           =   555
         End
         Begin VB.ComboBox cmb_Distri 
            Height          =   315
            Left            =   1680
            TabIndex        =   7
            Text            =   "cmb_DstDir"
            Top             =   720
            Width           =   11055
         End
         Begin VB.ComboBox cmb_Provin 
            Height          =   315
            Left            =   1680
            TabIndex        =   6
            Text            =   "cmb_PrvDir"
            Top             =   390
            Width           =   11055
         End
         Begin VB.ComboBox cmb_Depart 
            Height          =   315
            Left            =   1680
            TabIndex        =   5
            Text            =   "cmb_DptDir"
            Top             =   60
            Width           =   11055
         End
         Begin VB.TextBox txt_Descri 
            Height          =   315
            Left            =   1680
            MaxLength       =   250
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   1380
            Width           =   11055
         End
         Begin VB.Label Label5 
            Caption         =   "Distrito:"
            Height          =   285
            Left            =   90
            TabIndex        =   28
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Provincia:"
            Height          =   285
            Left            =   90
            TabIndex        =   27
            Top             =   420
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Código Zona:"
            Height          =   285
            Left            =   90
            TabIndex        =   24
            Top             =   1050
            Width           =   1275
         End
         Begin VB.Label Label4 
            Caption         =   "Descripción Zona:"
            Height          =   285
            Left            =   90
            TabIndex        =   20
            Top             =   1380
            Width           =   1425
         End
         Begin VB.Label Label1 
            Caption         =   "Departamento:"
            Height          =   285
            Left            =   90
            TabIndex        =   19
            Top             =   90
            Width           =   1425
         End
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   675
         Left            =   30
         TabIndex        =   25
         Top             =   30
         Width           =   12825
         _Version        =   65536
         _ExtentX        =   22622
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
         Begin Threed.SSPanel SSPanel12 
            Height          =   480
            Left            =   630
            TabIndex        =   26
            Top             =   90
            Width           =   4725
            _Version        =   65536
            _ExtentX        =   8334
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Zonas Ubicación Inmuebles"
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
            Picture         =   "PltPar_frm_028.frx":14B8
            Top             =   90
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_PosUbi_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_FlgCmb     As Integer
Dim l_str_Depart     As String
Dim l_str_Provin     As String
Dim l_str_Distri     As String

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   
   Call fs_Activa(False)
   Call gs_SetFocus(cmb_Depart)
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
   g_str_Parame = "USP_BORRAR_MNT_REFVIV (" & "'" & moddat_g_str_Codigo & "') "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
       Exit Sub
   End If
   
   Call fs_Buscar
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_Activa(True)
   Call fs_Limpia
   Call gs_SetFocus(grd_Listad)

   If grd_Listad.Rows = 0 Then
      cmd_Agrega.Enabled = True
      cmd_Editar.Enabled = False
      cmd_Borrar.Enabled = False
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
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_REFVIV WHERE "
   g_str_Parame = g_str_Parame & "REFVIV_CODZON = '" & moddat_g_str_Codigo & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   txt_Codigo.Text = Right(Trim(g_rst_Princi!REFVIV_CODZON), 2)
   txt_Descri.Text = Trim(g_rst_Princi!REFVIV_DESCRI)
   
   Call gs_BuscarCombo_Item(cmb_Depart, CInt(Left(Trim(g_rst_Princi!REFVIV_CODZON), 2)))

   Call moddat_gs_Carga_Provin(cmb_Provin, Format(Left(Trim(g_rst_Princi!REFVIV_CODZON), 2), "00"))
   Call gs_BuscarCombo_Item(cmb_Provin, CInt(Mid(Trim(g_rst_Princi!REFVIV_CODZON), 3, 2)))

   Call moddat_gs_Carga_Distri(cmb_Distri, Format(Left(Trim(g_rst_Princi!REFVIV_CODZON), 2), "00"), Format(Mid(Trim(g_rst_Princi!REFVIV_CODZON), 3, 2), "00"))
   Call gs_BuscarCombo_Item(cmb_Distri, CInt(Mid(Trim(g_rst_Princi!REFVIV_CODZON), 5, 2)))
      
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call fs_Activa(False)
   
   txt_Codigo.Enabled = False
   cmb_Depart.Enabled = False
   cmb_Provin.Enabled = False
   cmb_Distri.Enabled = False
   
   Call gs_SetFocus(txt_Descri)
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_str_Codigo     As String
   
   
   If Len(Trim(txt_Codigo.Text)) = 0 Then
      MsgBox "Debe ingresar el Código del Zona.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Codigo)
      Exit Sub
   End If
      
   txt_Codigo.Text = Format(txt_Codigo.Text, "00")
      
   If Len(Trim(txt_Descri.Text)) = 0 Then
      MsgBox "Debe ingresar la Descripción de la Zona.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Descri)
      Exit Sub
   End If
   
   If cmb_Depart.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Departamento al cual pertenece la Zona.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Depart)
      Exit Sub
   End If
   
   If cmb_Provin.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Provincia al cual pertenece la Zona.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Provin)
      Exit Sub
   End If
   
   If cmb_Distri.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Distrito al cual pertenece la Zona.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Distri)
      Exit Sub
   End If
   
   r_str_Codigo = Format(cmb_Depart.ItemData(cmb_Depart.ListIndex), "00") & Format(cmb_Provin.ItemData(cmb_Provin.ListIndex), "00") & Format(cmb_Distri.ItemData(cmb_Distri.ListIndex), "00") & txt_Codigo.Text
   
   If moddat_g_int_FlgGrb = 1 Then
      'Validar que el registro no exista
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM MNT_REFVIV WHERE "
      g_str_Parame = g_str_Parame & "REFVIV_CODZON = '" & r_str_Codigo & "' "
   
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
      
      g_str_Parame = "USP_MNT_REFVIV ("
      
      g_str_Parame = g_str_Parame & "'" & r_str_Codigo & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Descri.Text & "', "
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

Private Sub Form_Load()
   Me.Caption = modgen_g_str_NomPlt & " - Mantenimiento de Zonas de Posible Ubicación de Inmuebles"
   
   Call fs_Inicia
   
   Call fs_Activa(True)
   Call fs_Limpia
   Call fs_Buscar
   
   Call gs_CentraForm(Me)
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 1440
   grd_Listad.ColWidth(1) = 2730
   grd_Listad.ColWidth(2) = 2730
   grd_Listad.ColWidth(3) = 2730
   grd_Listad.ColWidth(4) = 2730
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter
   
   Call moddat_gs_Carga_Depart(cmb_Depart)
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   grd_Listad.Enabled = p_Activa
   cmd_Agrega.Enabled = p_Activa
   cmd_Editar.Enabled = p_Activa
   cmd_Borrar.Enabled = p_Activa
   
   txt_Codigo.Enabled = Not p_Activa
   txt_Descri.Enabled = Not p_Activa
   cmb_Depart.Enabled = Not p_Activa
   cmb_Provin.Enabled = Not p_Activa
   cmb_Distri.Enabled = Not p_Activa
   
   cmd_Grabar.Enabled = Not p_Activa
   cmd_Cancel.Enabled = Not p_Activa
End Sub

Private Sub fs_Limpia()
   txt_Codigo.Text = ""
   txt_Descri.Text = ""
   cmb_Depart.ListIndex = -1
   cmb_Depart.Text = ""
   
   cmb_Provin.Clear
   cmb_Provin.Text = ""
   
   cmb_Distri.Clear
   cmb_Distri.Text = ""
End Sub

Private Sub fs_Buscar()
   cmd_Agrega.Enabled = True
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   grd_Listad.Enabled = False
   
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_REFVIV "
   g_str_Parame = g_str_Parame & "ORDER BY REFVIV_CODZON ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
     Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = Trim(g_rst_Princi!REFVIV_CODZON)
      
      grd_Listad.Col = 1
      grd_Listad.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!REFVIV_CODZON, 2) & "0000")
      
      grd_Listad.Col = 2
      grd_Listad.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!REFVIV_CODZON, 4) & "00")
      
      grd_Listad.Col = 3
      grd_Listad.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!REFVIV_CODZON, 6))
      
      grd_Listad.Col = 4
      grd_Listad.Text = Trim(g_rst_Princi!REFVIV_DESCRI)
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If grd_Listad.Rows > 0 Then
      cmd_Editar.Enabled = True
      cmd_Borrar.Enabled = True
      grd_Listad.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub txt_Codigo_GotFocus()
   Call gs_SelecTodo(txt_Codigo)
End Sub

Private Sub txt_Codigo_KeyPress(KeyAscii As Integer)
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
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & ",-_.(<> )$/")
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

Private Sub cmb_Depart_Change()
   l_str_Depart = cmb_Depart.Text
End Sub

Private Sub cmb_Depart_Click()
   If cmb_Depart.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_Provin.Clear
         cmb_Distri.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_Provin, Format(cmb_Depart.ItemData(cmb_Depart.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_Provin)
      End If
   End If
End Sub

Private Sub cmb_Depart_GotFocus()
   Call SendMessage(cmb_Depart.hWnd, CB_SHOWDROPDOWN, 1, 0&)
   
   l_int_FlgCmb = True
   l_str_Depart = cmb_Depart.Text
End Sub

Private Sub cmb_Depart_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Depart, l_str_Depart)
      l_int_FlgCmb = True
      
      cmb_Provin.Clear
      cmb_Distri.Clear
      If cmb_Depart.ListIndex > -1 Then
         l_str_Depart = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_Provin, Format(cmb_Depart.ItemData(cmb_Depart.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_Provin)
   End If
End Sub

Private Sub cmb_Provin_Change()
   l_str_Provin = cmb_Provin.Text
End Sub

Private Sub cmb_Provin_Click()
   If cmb_Provin.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_Distri.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_Distri, Format(cmb_Depart.ItemData(cmb_Depart.ListIndex), "00"), Format(cmb_Provin.ItemData(cmb_Provin.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_Distri)
      End If
   End If
End Sub

Private Sub cmb_Provin_GotFocus()
   l_int_FlgCmb = True
   l_str_Provin = cmb_Provin.Text
End Sub

Private Sub cmb_Provin_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Provin, l_str_Provin)
      l_int_FlgCmb = True
      
      cmb_Distri.Clear
      If cmb_Provin.ListIndex > -1 Then
         l_str_Distri = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_Distri, Format(cmb_Depart.ItemData(cmb_Depart.ListIndex), "00"), Format(cmb_Provin.ItemData(cmb_Provin.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_Distri)
   End If
End Sub

Private Sub cmb_Distri_Change()
   l_str_Distri = cmb_Distri.Text
End Sub

Private Sub cmb_Distri_Click()
   If cmb_Distri.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_Codigo)
      End If
   End If
End Sub

Private Sub cmb_Distri_GotFocus()
   l_int_FlgCmb = True
   l_str_Distri = cmb_Distri.Text
End Sub

Private Sub cmb_Distri_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Distri, l_str_Distri)
      l_int_FlgCmb = True
      
      If cmb_Distri.ListIndex > -1 Then
         l_str_Distri = ""
      End If
      
      Call gs_SetFocus(txt_Codigo)
   End If
End Sub

