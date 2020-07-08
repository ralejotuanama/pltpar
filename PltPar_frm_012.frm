VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_OpeFin_1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7920
   ClientLeft      =   3300
   ClientTop       =   1875
   ClientWidth     =   8145
   Icon            =   "PltPar_frm_012.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7905
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   8175
      _Version        =   65536
      _ExtentX        =   14420
      _ExtentY        =   13944
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
         TabIndex        =   14
         Top             =   3150
         Width           =   8085
         _Version        =   65536
         _ExtentX        =   14261
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
            Left            =   5940
            Picture         =   "PltPar_frm_012.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   675
            Left            =   6630
            Picture         =   "PltPar_frm_012.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   7350
            Picture         =   "PltPar_frm_012.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   765
         Left            =   30
         TabIndex        =   15
         Top             =   7080
         Width           =   8085
         _Version        =   65536
         _ExtentX        =   14261
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
            Left            =   7350
            Picture         =   "PltPar_frm_012.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   6630
            Picture         =   "PltPar_frm_012.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   3075
         Left            =   30
         TabIndex        =   18
         Top             =   30
         Width           =   8085
         _Version        =   65536
         _ExtentX        =   14261
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
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   4736
            _Version        =   393216
            Rows            =   12
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   1530
            TabIndex        =   19
            Top             =   60
            Width           =   6135
            _Version        =   65536
            _ExtentX        =   10821
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   285
            Left            =   90
            TabIndex        =   20
            Top             =   60
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Código Operación"
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
         Height          =   3075
         Left            =   30
         TabIndex        =   21
         Top             =   3960
         Width           =   8085
         _Version        =   65536
         _ExtentX        =   14261
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
         Begin VB.TextBox txt_Abrevi 
            Height          =   315
            Left            =   2010
            MaxLength       =   25
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   720
            Width           =   6015
         End
         Begin VB.ComboBox cmb_MatCon 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   2700
            Width           =   3225
         End
         Begin VB.ComboBox cmb_OpeCon 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   2370
            Width           =   975
         End
         Begin VB.TextBox txt_Descri 
            Height          =   315
            Left            =   2010
            MaxLength       =   200
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   390
            Width           =   6015
         End
         Begin VB.TextBox txt_CodOpe 
            Height          =   315
            Left            =   2010
            MaxLength       =   6
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   60
            Width           =   1215
         End
         Begin VB.ComboBox cmb_LibCon 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1050
            Width           =   3225
         End
         Begin VB.ComboBox cmb_FlgEst 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1380
            Width           =   3225
         End
         Begin VB.ComboBox cmb_FlgCre 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1710
            Width           =   3225
         End
         Begin VB.ComboBox cmb_LavDin 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Abreviación:"
            Height          =   285
            Left            =   90
            TabIndex        =   30
            Top             =   750
            Width           =   1815
         End
         Begin VB.Label Label8 
            Caption         =   "Matriz Contable:"
            Height          =   285
            Left            =   90
            TabIndex        =   29
            Top             =   2730
            Width           =   1605
         End
         Begin VB.Label Label7 
            Caption         =   "Flag Operac. Contable:"
            Height          =   285
            Left            =   90
            TabIndex        =   28
            Top             =   2400
            Width           =   1875
         End
         Begin VB.Label Label4 
            Caption         =   "Descripción:"
            Height          =   285
            Left            =   90
            TabIndex        =   27
            Top             =   420
            Width           =   1815
         End
         Begin VB.Label Label3 
            Caption         =   "Código Operación:"
            Height          =   285
            Left            =   90
            TabIndex        =   26
            Top             =   90
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Libro Contable:"
            Height          =   285
            Left            =   90
            TabIndex        =   25
            Top             =   1080
            Width           =   1605
         End
         Begin VB.Label Label2 
            Caption         =   "Flag Estado:"
            Height          =   285
            Left            =   90
            TabIndex        =   24
            Top             =   1410
            Width           =   1605
         End
         Begin VB.Label Label5 
            Caption         =   "Flag de Créditos:"
            Height          =   285
            Left            =   90
            TabIndex        =   23
            Top             =   1740
            Width           =   1605
         End
         Begin VB.Label Label6 
            Caption         =   "Flag Lavado Dinero:"
            Height          =   285
            Left            =   90
            TabIndex        =   22
            Top             =   2070
            Width           =   1605
         End
      End
   End
End
Attribute VB_Name = "frm_OpeFin_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_MatCon()      As moddat_tpo_Genera

Private Sub cmb_LibCon_Click()
   Call gs_SetFocus(cmb_FlgEst)
End Sub

Private Sub cmb_LibCon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_LibCon_Click
   End If
End Sub

Private Sub cmb_FlgEst_Click()
   Call gs_SetFocus(cmb_FlgCre)
End Sub

Private Sub cmb_FlgEst_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_FlgEst_Click
   End If
End Sub

Private Sub cmb_FlgCre_Click()
   Call gs_SetFocus(cmb_LavDin)
End Sub

Private Sub cmb_FlgCre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_FlgCre_Click
   End If
End Sub

Private Sub cmb_LavDin_Click()
   Call gs_SetFocus(cmb_OpeCon)
End Sub

Private Sub cmb_LavDin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_LavDin_Click
   End If
End Sub

Private Sub cmb_OpeCon_Click()
   Call gs_SetFocus(cmb_MatCon)
End Sub

Private Sub cmb_OpeCon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_OpeCon_Click
   End If
End Sub

Private Sub cmb_MatCon_Click()
   Call gs_SetFocus(cmd_Grabar)
End Sub

Private Sub cmb_matcon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_MatCon_Click
   End If
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   
   Call fs_Activa(False)
   Call gs_SetFocus(txt_CodOpe)
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
   g_str_Parame = "USP_BORRAR_CRE_PRODUC (" & "'" & moddat_g_str_Codigo & "', "
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
   Call fs_Activa(True)
   Call fs_Limpia
   Call gs_SetFocus(grd_Listad)

   If grd_Listad.Rows = 0 Then
      cmd_Agrega.Enabled = True
      cmd_Editar.Enabled = False
      grd_Listad.Enabled = False
   End If
End Sub

Private Sub cmd_Editar_Click()
   Dim r_int_Contad     As Integer
   
   grd_Listad.Col = 0
   moddat_g_str_Codigo = grd_Listad.Text
         
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgGrb = 2
   
   'Obteniendo Información del Registro
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM OPERACION_TIPO WHERE "
   g_str_Parame = g_str_Parame & "OPERACION = '" & moddat_g_str_Codigo & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   
   txt_CodOpe.Text = Trim(g_rst_Genera!OPERACION)
   txt_Descri.Text = Trim(g_rst_Genera!DESCRIPCION)
   txt_Abrevi.Text = Trim(g_rst_Genera!ABREVIACION & "")
   
   cmb_MatCon.ListIndex = gf_Busca_Arregl(l_arr_MatCon, Trim(g_rst_Genera!MATRIZ_CTBL & "")) - 1
   
   Call gs_BuscarCombo_Item(cmb_LibCon, g_rst_Genera!NRO_LIBRO)
   Call gs_BuscarCombo_Item(cmb_FlgEst, g_rst_Genera!FLAG_ESTADO)
   
   Call gs_BuscarCombo_Item(cmb_LavDin, g_rst_Genera!FLAG_LAVADO_DIN & "")
   Call gs_BuscarCombo_Item(cmb_OpeCon, g_rst_Genera!FLAG_OPER_CONTABLE)
   
   For r_int_Contad = 0 To cmb_FlgCre.ListCount - 1
      cmb_FlgCre.ListIndex = r_int_Contad
      
      If g_rst_Genera!FLAG_CREDITOS = Left(cmb_FlgCre.Text, 1) Then
         Exit For
      End If
   Next r_int_Contad
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   Call fs_Activa(False)
   
   txt_CodOpe.Enabled = False
   Call gs_SetFocus(txt_Descri)
End Sub

Private Sub cmd_Grabar_Click()
   'txt_CodOpe.Text = Format(txt_CodOpe.Text, "000000")
   
   If Len(Trim(txt_CodOpe.Text)) = 0 Then
      MsgBox "Debe ingresar el Código de la Operación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_CodOpe)
      Exit Sub
   End If
      
   If Len(Trim(txt_Descri.Text)) = 0 Then
      MsgBox "Debe ingresar la Descripción de la Operación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Descri)
      Exit Sub
   End If
   
   If cmb_LibCon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Libro Contable.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_LibCon)
      Exit Sub
   End If
   
   If cmb_FlgEst.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Flag de Estado.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_FlgEst)
      Exit Sub
   End If
   
   If cmb_FlgCre.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Flag de Créditos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_FlgCre)
      Exit Sub
   End If
   
   If cmb_LavDin.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Flag de Lavado de Dinero.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_LavDin)
      Exit Sub
   End If
   
   If cmb_OpeCon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Flag de Operación Contable.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_OpeCon)
      Exit Sub
   End If
   
   If cmb_MatCon.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Matriz Contable.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_MatCon)
      Exit Sub
   End If
   
   If moddat_g_int_FlgGrb = 1 Then
      'Validar que el registro no exista
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM OPERACION_TIPO WHERE "
      g_str_Parame = g_str_Parame & "OPERACION = '" & txt_CodOpe.Text & "' "
   
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
      
      g_str_Parame = "USP_OPERACION_TIPO ("
      
      g_str_Parame = g_str_Parame & "'" & txt_CodOpe.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Descri.Text & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_LibCon.ItemData(cmb_LibCon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & CStr(cmb_FlgEst.ItemData(cmb_FlgEst.ListIndex)) & "', "
      g_str_Parame = g_str_Parame & "'" & Left(cmb_FlgCre.Text, 1) & "', "
      g_str_Parame = g_str_Parame & "'" & CStr(cmb_LavDin.ItemData(cmb_LavDin.ListIndex)) & "', "
      g_str_Parame = g_str_Parame & "'" & CStr(cmb_OpeCon.ItemData(cmb_OpeCon.ListIndex)) & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Abrevi.Text & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_MatCon(cmb_MatCon.ListIndex + 1).Genera_Codigo & "', "
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
   Me.Caption = modgen_g_str_NomPlt & " - Mantenimiento de Operaciones Financieras"
   
   Call fs_Inicia
   
   Call fs_Activa(True)
   Call fs_Limpia
   Call fs_Buscar
   
   Call gs_CentraForm(Me)
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 1425
   grd_Listad.ColWidth(1) = 6130
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   
   Call moddat_gs_Carga_LibCon(cmb_LibCon)
   Call moddat_gs_Carga_MatCon(cmb_MatCon, l_arr_MatCon)
   
   'Flag de Estado
   cmb_FlgEst.Clear
   
   cmb_FlgEst.AddItem "ACTIVO"
   cmb_FlgEst.ItemData(cmb_FlgEst.NewIndex) = 1
   
   cmb_FlgEst.AddItem "INACTIVO"
   cmb_FlgEst.ItemData(cmb_FlgEst.NewIndex) = 0
   
   cmb_FlgEst.ListIndex = -1
   
   'Flag de Créditos
   cmb_FlgCre.Clear
   
   cmb_FlgCre.AddItem "D: DESEMBOLSO"
   cmb_FlgCre.AddItem "P: PAGO CUOTAS"
   cmb_FlgCre.AddItem "O: OTROS CARGOS"
   cmb_FlgCre.AddItem "A: GASTOS ADMINISTRATIVOS"
   
   cmb_FlgCre.ListIndex = -1
   
   'Flag de Lavado de Dinero
   cmb_LavDin.Clear
   
   cmb_LavDin.AddItem "SI"
   cmb_LavDin.ItemData(cmb_LavDin.NewIndex) = 1
   
   cmb_LavDin.AddItem "NO"
   cmb_LavDin.ItemData(cmb_LavDin.NewIndex) = 0
   
   cmb_LavDin.ListIndex = -1
      
   'Flag de Operación Contable
   cmb_OpeCon.Clear
   
   cmb_OpeCon.AddItem "SI"
   cmb_OpeCon.ItemData(cmb_OpeCon.NewIndex) = 1
   
   cmb_OpeCon.AddItem "NO"
   cmb_OpeCon.ItemData(cmb_OpeCon.NewIndex) = 0
   
   cmb_OpeCon.ListIndex = -1
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   grd_Listad.Enabled = p_Activa
   cmd_Agrega.Enabled = p_Activa
   cmd_Editar.Enabled = p_Activa
   
   txt_CodOpe.Enabled = Not p_Activa
   txt_Descri.Enabled = Not p_Activa
   cmb_LibCon.Enabled = Not p_Activa
   cmb_FlgEst.Enabled = Not p_Activa
   cmb_FlgCre.Enabled = Not p_Activa
   cmb_LavDin.Enabled = Not p_Activa
   cmb_OpeCon.Enabled = Not p_Activa
   txt_Abrevi.Enabled = Not p_Activa
   cmb_MatCon.Enabled = Not p_Activa
   
   cmd_Grabar.Enabled = Not p_Activa
   cmd_Cancel.Enabled = Not p_Activa
End Sub

Private Sub fs_Limpia()
   txt_CodOpe.Text = ""
   txt_Descri.Text = ""
   txt_Abrevi.Text = ""
   cmb_LibCon.ListIndex = -1
   cmb_FlgEst.ListIndex = -1
   cmb_FlgCre.ListIndex = -1
   cmb_LavDin.ListIndex = -1
   cmb_OpeCon.ListIndex = -1
   cmb_MatCon.ListIndex = -1
End Sub

Private Sub fs_Buscar()
   cmd_Agrega.Enabled = True
   cmd_Editar.Enabled = False
   grd_Listad.Enabled = False
   
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM OPERACION_TIPO "
   g_str_Parame = g_str_Parame & "ORDER BY OPERACION ASC "

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
      grd_Listad.Text = Trim(g_rst_Genera!OPERACION)
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Genera!DESCRIPCION)
      
      g_rst_Genera.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   If grd_Listad.Rows > 0 Then
      cmd_Editar.Enabled = True
      grd_Listad.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub txt_CodOpe_GotFocus()
   Call gs_SelecTodo(txt_CodOpe)
End Sub

Private Sub txt_CodOpe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Descri)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS)
   End If
End Sub

Private Sub txt_Descri_GotFocus()
   Call gs_SelecTodo(txt_Descri)
End Sub

Private Sub txt_Descri_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Abrevi)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & ",-_.( )$/")
   End If
End Sub

Private Sub txt_Abrevi_GotFocus()
   Call gs_SelecTodo(txt_Abrevi)
End Sub

Private Sub txt_Abrevi_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_LibCon)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & ",-_.( )$/")
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


