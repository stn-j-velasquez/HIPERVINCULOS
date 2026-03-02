Imports System.Data.SqlClient
Imports System.Windows.Forms
Public Class frmMantItemsServicios

#Region "Propiedades"

    Property LaTransaccionProvieneDesdeEtapaDeCotizacion As Boolean
    Public TipoConsulta As Integer = 1
    Private strSQL As String = String.Empty
    Private oDT As New DataTable
    Dim Hp As New clsHELPER
    Dim CodigoCliente As String
    Dim Conexion As New SqlConnection
    Dim Seguridad As New ClsBtnSeguridad
    Private colEmpresa As Color
#End Region

    Private Sub frmMantItemsServicios_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            TxpCliente.ciTXT_COD.Select()
            TxpCliente.ciTXT_COD.Focus()

            'Inicializamos los TXP
            TxpCliente.CONECCION = cconnect
            TxpTemporada.CONECCION = cconnect
            TxpItems.CONECCION = cconnect
            TxpEstado.CONECCION = cconnect
            TxpProveedor.CONECCION = cconnect
            LlenarCombo()
            InhabilitaDatos()

            Dim oDt As DataTable = Hp.DevuelveDatos(String.Format("SELECT * FROM SEG_Empresas where cod_empresa = '{0}'", vemp), cSEGURIDAD)
            colEmpresa = Color.FromArgb(oDt.Rows(0)("ColorFondo_R"), oDt.Rows(0)("ColorFondo_G"), oDt.Rows(0)("ColorFondo_B"))

            Seguridad.GetBotonesJanus(vper, vemp, Me.Name, BarraOpciones, "")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub FondoDegrade(sender As Object, e As PaintEventArgs) Handles panGui.Paint
        FondoDegradeDiagonalEnPanel(sender, e, colEmpresa)
    End Sub

#Region "Opciones de Busqueda"

    Private Sub RbCliente_CheckedChanged(sender As Object, e As EventArgs) Handles RbCliente.CheckedChanged
        Try
            TipoConsulta = 1
            Inhabilita()
            grbCliente.Visible = True
            TxpCliente.ciTXT_COD.Focus()
            TxpCliente.ciTXT_COD.Text = String.Empty
            TxpCliente.ciTXT_DES.Text = String.Empty
            TxpTemporada.ciTXT_COD.Text = String.Empty
            TxpTemporada.ciTXT_DES.Text = String.Empty
            
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RbItems_CheckedChanged(sender As Object, e As EventArgs) Handles RbItems.CheckedChanged
        Try
            TipoConsulta = 2
            Inhabilita()
            grbItems.Visible = True
            TxpItems.ciTXT_COD.Focus()
            TxpItems.ciTXT_COD.Text = String.Empty
            TxpItems.ciTXT_DES.Text = String.Empty
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RbEstado_CheckedChanged(sender As Object, e As EventArgs) Handles RbEstado.CheckedChanged
        Try
            TipoConsulta = 3
            Inhabilita()
            grbEstado.Visible = True
            TxpEstado.ciTXT_COD.Focus()
            TxpEstado.ciTXT_COD.Text = String.Empty
            TxpEstado.ciTXT_DES.Text = String.Empty
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TxpCliente_TxpEvent_AntesDeConsultar(ctrlTXP As WinFormsControls.ucPROMPT_BASE) Handles TxpCliente.TxpEvent_AntesDeConsultar
        Try
            With TxpCliente
                .TipoQuery = WinFormsControls.Txp.enuTipoQuery.SP_SQL
                .CadenaSP = "select Abr_Cliente,Nom_Cliente from tg_cliente where Abr_Cliente like '" & TxpCliente.ciTXT_COD.Text & "%'"
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TxpCliente_TxpEvent_DespuesDeConsultar(ctrlTXP As WinFormsControls.ucPROMPT_BASE, oDrResultado As DataRow) Handles TxpCliente.TxpEvent_DespuesDeConsultar
        Try
            If (oDrResultado) Is Nothing Then
                TxpCliente.ciTXT_COD.Focus()
            End If
            TxpTemporada.ciTXT_COD.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TxpTemporada_TxpEvent_AntesDeConsultar(ctrlTXP As WinFormsControls.ucPROMPT_BASE) Handles TxpTemporada.TxpEvent_AntesDeConsultar
        Try
            CodigoCliente = Hp.DevuelveDato("select Cod_Cliente from tg_cliente where Abr_Cliente = '" & TxpCliente.ciTXT_COD.Text & "'", cconnect)
            With TxpTemporada
                .TipoQuery = WinFormsControls.Txp.enuTipoQuery.SP_SQL
                .CadenaSP = "select Cod_TemCli, Nom_TemCli from TG_TemCli where Cod_Cliente = '" & CodigoCliente & "' and Cod_TemCli like '" & TxpTemporada.ciTXT_COD.Text & "%' AND Flg_Activo = 'S'"
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TxpTemporada_TxpEvent_DespuesDeConsultar(ctrlTXP As WinFormsControls.ucPROMPT_BASE, oDrResultado As DataRow) Handles TxpTemporada.TxpEvent_DespuesDeConsultar
        Try
            If (oDrResultado) Is Nothing Then
                TxpTemporada.ciTXT_COD.Focus()
            End If
            TxtFamilia.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TxpItems_TxpEvent_AntesDeConsultar(ctrlTXP As WinFormsControls.ucPROMPT_BASE) Handles TxpItems.TxpEvent_AntesDeConsultar
        Try

            If Len(Trim(TxpItems.ciTXT_COD.Text)) >= 2 And Len(Trim(TxpItems.ciTXT_COD.Text)) < 8 Then

                If Len(Trim(TxpItems.ciTXT_COD.Text)) = 5 Then
                    TxpItems.ciTXT_COD.Text = CompletaCodigos(Trim(TxpItems.ciTXT_COD.Text), 8, 2)
                Else
                    TxpItems.ciTXT_COD.Text = CompletaCodigo(Trim(TxpItems.ciTXT_COD.Text), 8, 2)
                End If
            End If

            With TxpItems
                .TipoQuery = WinFormsControls.Txp.enuTipoQuery.SP_SQL
                '.CadenaSP = "SELECT Cod_Item as Código, Des_Item as Descripción FROM LG_ITEM WHERE Cod_Item like'" & TxpItems.ciTXT_COD.Text & "%'"
                .CadenaSP = "exec ES_MUESTRA__ITEMS_PROCESOS_MANUFACTURA '" & TxpItems.ciTXT_COD.Text & "'"
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Function CompletaCodigo(CodOrigen As String, longcodfinal As Integer, PosfinalCod As Integer) As String
        ' CodOrigen     = Es el codigo que sera pasado por parametro
        ' LongCodFinal  = Es el tamaño del Codigo a devolver
        ' PosFinalCod   = Es la posicion de la 1era parte del codigo
        Dim Contador As Integer
        CompletaCodigo = Mid(CodOrigen, 1, PosfinalCod).ToUpper
        For Contador = 1 To longcodfinal - Len(CodOrigen)
            CompletaCodigo = CompletaCodigo & "0"
        Next
        If Len(CodOrigen) = 2 Then
            CompletaCodigo = CompletaCodigo
        Else
            CompletaCodigo = CompletaCodigo & Mid(CodOrigen, 3, 2)
        End If

    End Function

    Public Function CompletaCodigos(CodOrigen As String, longcodfinal As Integer, PosfinalCod As Integer) As String
        ' CodOrigen     = Es el codigo que sera pasado por parametro
        ' LongCodFinal  = Es el tamaño del Codigo a devolver
        ' PosFinalCod   = Es la posicion de la 1era parte del codigo
        Dim Contador As Integer
        CompletaCodigos = Mid(CodOrigen, 1, PosfinalCod).ToUpper
        For Contador = 1 To longcodfinal - Len(CodOrigen)
            CompletaCodigos = CompletaCodigos & "0"
        Next
        If Len(CodOrigen) = 2 Then
            CompletaCodigos = CompletaCodigos
        Else
            CompletaCodigos = CompletaCodigos & Mid(CodOrigen, 3, 3)
        End If

    End Function
    Private Sub TxpItems_TxpEvent_DespuesDeConsultar(ctrlTXP As WinFormsControls.ucPROMPT_BASE, oDrResultado As DataRow) Handles TxpItems.TxpEvent_DespuesDeConsultar
        Try
            If (oDrResultado) Is Nothing Then
                TxpItems.ciTXT_COD.Focus()
            Else
                BtnBuscar.Focus()
            End If
            TxtFamilia.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TxpEstado_TxpEvent_AntesDeConsultar(ctrlTXP As WinFormsControls.ucPROMPT_BASE) Handles TxpEstado.TxpEvent_AntesDeConsultar
        Try
            With TxpEstado
                .TipoQuery = WinFormsControls.Txp.enuTipoQuery.SP_SQL
                .CadenaSP = "SELECT flg_status as Código, des_status as Descripción FROM TG_StaDes where flg_status like '" & TxpEstado.ciTXT_COD.Text & "%'"
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TxpEstado_TxpEvent_DespuesDeConsultar(ctrlTXP As WinFormsControls.ucPROMPT_BASE, oDrResultado As DataRow) Handles TxpEstado.TxpEvent_DespuesDeConsultar
        Try
            If (oDrResultado) Is Nothing Then
                TxpEstado.ciTXT_COD.Focus()
            End If
            TxtFamilia.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TxpProveedor_TxpEvent_AntesDeConsultar(ctrlTXP As WinFormsControls.ucPROMPT_BASE) Handles TxpProveedor.TxpEvent_AntesDeConsultar
        Try
            With TxpProveedor
                .TipoQuery = WinFormsControls.Txp.enuTipoQuery.SP_SQL
                .CadenaSP = "EXEC UP_SEL_PROVEEDORES_CF_ACUMULADOS '3','" & TxpProveedor.ciTXT_COD.Text & "','" & TxpProveedor.ciTXT_DES.Text & "'"
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TxpProveedor_TxpEvent_DespuesDeConsultar(ctrlTXP As WinFormsControls.ucPROMPT_BASE, oDrResultado As DataRow) Handles TxpProveedor.TxpEvent_DespuesDeConsultar
        Try
            If (oDrResultado) Is Nothing Then
                TxpProveedor.ciTXT_COD.Focus()
            End If
            TxtFamilia.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TxtFamilia_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtFamilia.KeyPress
        Try
            If e.KeyChar = CChar(ChrW(Keys.Enter)) Then
                BtnBuscar.Focus()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region

#Region "Llenar Combo"
    Private Sub LlenarCombo()
        Try
            'Declaramos los DT para los Combos
            Dim ListaFamilia As New DataTable
            Dim ListaStatus As New DataTable
            Dim ListaUnidadMedida As New DataTable
            Dim ListaClase As New DataTable
            Dim ListaModoProceso As New DataTable
            Dim ListaGrupo As New DataTable
            Dim ListaOrigen As New DataTable
            Dim ListaGradoDificultad As New DataTable
            Dim ListaTalla As New DataTable
            Dim ListaEstiloCliente As New DataTable
            Dim ListaColor As New DataTable
            Dim ListaDestino As New DataTable
            Dim ListaPO As New DataTable
            Dim ListaMotivoPreProduccion As New DataTable

            Conexion.ConnectionString = cconnect

            'Familia
            Dim cmdfamilia As New SqlCommand("SELECT cod_famitem, rtrim(ltrim(des_famitem))des_famitem FROM LG_FamIte order by 1 asc")
            cmdfamilia.Connection = Conexion
            Conexion.Open()
            ListaFamilia.Load(cmdfamilia.ExecuteReader)
            Conexion.Close()
            With cboCod_FamItem
                .DataSource = ListaFamilia
                .DisplayMember = "des_famitem"
                .ValueMember = "cod_famitem"
            End With

            'Status
            Dim cmdstatus As New SqlCommand("SELECT Flg_Status,Des_Status  FROM LG_Status_Servicios order by 1 asc")
            cmdstatus.Connection = Conexion
            Conexion.Open()
            ListaStatus.Load(cmdstatus.ExecuteReader)
            Conexion.Close()
            With cboFlg_Status
                .DataSource = ListaStatus
                .DisplayMember = "Des_Status"
                .ValueMember = "Flg_Status"
            End With

            'Unidad Medida
            Dim cmdunidadmedida As New SqlCommand("select Cod_UniMed, Des_UniMed FROM TG_UniMed order by 1 asc")
            cmdunidadmedida.Connection = Conexion
            Conexion.Open()
            ListaUnidadMedida.Load(cmdunidadmedida.ExecuteReader)
            Conexion.Close()
            With cboCod_UniMed
                .DataSource = ListaUnidadMedida
                .DisplayMember = "Des_UniMed"
                .ValueMember = "Cod_UniMed"
            End With

            'Clase
            Dim cmdclase As New SqlCommand("SELECT cod_claitem, des_claitem FROM LG_Claitem order by 1 asc")
            cmdclase.Connection = Conexion
            Conexion.Open()
            ListaClase.Load(cmdclase.ExecuteReader)
            Conexion.Close()
            With cboCod_ClaItem
                .DataSource = ListaClase
                .DisplayMember = "des_claitem"
                .ValueMember = "cod_claitem"
            End With

            'Modo Proceso
            Dim cmdModoProceso As New SqlCommand("SELECT  Flg_ModoProceso, Des_ModoProceso FROM ES_ModoProceso order by 1 asc")
            cmdModoProceso.Connection = Conexion
            Conexion.Open()
            ListaModoProceso.Load(cmdModoProceso.ExecuteReader)
            Conexion.Close()
            With cboModoProceso
                .DataSource = ListaModoProceso
                .DisplayMember = "Des_ModoProceso"
                .ValueMember = "Flg_ModoProceso"
            End With

            'Origen
            Dim cmdOrigen As New SqlCommand("SELECT cod_origen, des_origen FROM LG_Origen order by 1 asc")
            cmdOrigen.Connection = Conexion
            Conexion.Open()
            ListaOrigen.Load(cmdOrigen.ExecuteReader)
            Conexion.Close()
            With cboCod_Origen
                .DataSource = ListaOrigen
                .DisplayMember = "des_origen"
                .ValueMember = "cod_origen"
            End With

            'Grado Dificultad
            Dim cmdGradoDificultad As New SqlCommand("SELECT Tipo_Grado_Dificultad, Des_Tipo_Grado_Dificultad FROm LG_GRADO_DIFICULTAD_ITEM order by 1 asc")
            cmdGradoDificultad.Connection = Conexion
            Conexion.Open()
            ListaGradoDificultad.Load(cmdGradoDificultad.ExecuteReader)
            Conexion.Close()
            With cboGradoDif
                .DataSource = ListaGradoDificultad
                .DisplayMember = "Des_Tipo_Grado_Dificultad"
                .ValueMember = "Tipo_Grado_Dificultad"
            End With


            '=>Talla
            Dim cmdTalla As New SqlCommand("select 'S' Codigo, 'SI' Descripcion union all select 'N' Codigo, 'NO' Descripcion")
            cmdTalla.Connection = Conexion
            Conexion.Open()
            ListaTalla.Load(cmdTalla.ExecuteReader)
            With cboIde_Talla
                .DataSource = ListaTalla
                .DisplayMember = "Descripcion"
                .ValueMember = "Codigo"
            End With
            Conexion.Close()

            '=>Estilo
            Dim cmdEstilo As New SqlCommand("select 'S' Codigo, 'SI' Descripcion union all select 'N' Codigo, 'NO' Descripcion")
            cmdEstilo.Connection = Conexion
            Conexion.Open()
            ListaEstiloCliente.Load(cmdEstilo.ExecuteReader)
            With cboIde_EsCli
                .DataSource = ListaEstiloCliente
                .DisplayMember = "Descripcion"
                .ValueMember = "Codigo"
            End With
            Conexion.Close()

            '=>Color
            Dim cmdColor As New SqlCommand("select 'S' Codigo, 'SI' Descripcion union all select 'N' Codigo, 'NO' Descripcion")
            cmdColor.Connection = Conexion
            Conexion.Open()
            ListaColor.Load(cmdColor.ExecuteReader)
            With cboIde_Color
                .DataSource = ListaColor
                .DisplayMember = "Descripcion"
                .ValueMember = "Codigo"
            End With
            Conexion.Close()

            '=>Destino
            Dim cmdDestino As New SqlCommand("select 'S' Codigo, 'SI' Descripcion union all select 'N' Codigo, 'NO' Descripcion")
            cmdDestino.Connection = Conexion
            Conexion.Open()
            ListaDestino.Load(cmdDestino.ExecuteReader)
            With cboIde_Destino
                .DataSource = ListaDestino
                .DisplayMember = "Descripcion"
                .ValueMember = "Codigo"
            End With
            Conexion.Close()

            '=>PO
            Dim cmdPO As New SqlCommand("select 'S' Codigo, 'SI' Descripcion union all select 'N' Codigo, 'NO' Descripcion")
            cmdPO.Connection = Conexion
            Conexion.Open()
            ListaPO.Load(cmdPO.ExecuteReader)
            With CboIde_PO
                .DataSource = ListaPO
                .DisplayMember = "Descripcion"
                .ValueMember = "Codigo"
            End With
            Conexion.Close()

            '=>Motivo Pre-Produccion
            Dim cmdPreProduccion As New SqlCommand("SELECT cod_motprepro , des_motprepro  FROM TG_MotPrePro order by 1 asc")
            cmdPreProduccion.Connection = Conexion
            Conexion.Open()
            ListaMotivoPreProduccion.Load(cmdPreProduccion.ExecuteReader)
            With cboCod_MotPrePro
                .DataSource = ListaMotivoPreProduccion
                .DisplayMember = "des_motprepro"
                .ValueMember = "cod_motprepro"
            End With
            Conexion.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
#End Region

    Private Sub BtnBuscar_Click(sender As Object, e As EventArgs) Handles BtnBuscar.Click
        Try
            CargaLista()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub CargaLista()
        Try

            strSQL = "EXEC ES_SM_ItemServicios_ClienteTemp"
            strSQL &= vbNewLine & String.Format(" @opcion            = '{0}'", TipoConsulta)
            strSQL &= vbNewLine & String.Format(",@cod_cliente       = '{0}'", CodigoCliente)
            strSQL &= vbNewLine & String.Format(",@cod_temcli        = '{0}'", TxpTemporada.ciTXT_COD.Text)
            strSQL &= vbNewLine & String.Format(",@cod_item          = '{0}'", TxpItems.ciTXT_COD.Text)
            strSQL &= vbNewLine & String.Format(",@flg_status        = '{0}'", TxpEstado.ciTXT_COD.Text)
            strSQL &= vbNewLine & String.Format(",@cod_proveedor     = '{0}'", TxpProveedor.ciTXT_COD.Text)
            strSQL &= vbNewLine & String.Format(",@cod_famitem       = '{0}'", TxtFamilia.Text)
            strSQL &= vbNewLine & String.Format(",@COD_FABRICA       = '{0}'", "1")
            strSQL &= vbNewLine & String.Format(",@COD_ORDPRO        = '{0}'", "")
            strSQL &= vbNewLine & String.Format(",@COD_ESTCLI        = '{0}'", txtEstilo.Text)

            oDT = Hp.DevuelveDatos(strSQL, cconnect)
            GrdEstampadoBordado.DataSource = oDT
            GrdEstampadoBordado.RootTable.RowHeight = 30
            'CheckLayoutGridEx(GrdEstampadoBordado)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub GrdEstampadoBordado_SelectionChanged(sender As Object, e As EventArgs) Handles GrdEstampadoBordado.SelectionChanged
        Try
            SelectRow()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub SelectRow()
        Try
            Dim ItemBordadoEstampado As DataRow
            ItemBordadoEstampado = ObtenerDr_DeGridEx(GrdEstampadoBordado)

            If GrdEstampadoBordado.RowCount > 0 Then
                txtcoditem.Text = ItemBordadoEstampado("cod_item")
                txtDesItem.Text = ItemBordadoEstampado("Des_Item")

                cboCod_FamItem.SelectedValue = ItemBordadoEstampado("Cod_FamItem")
                cboFlg_Status.SelectedValue = ItemBordadoEstampado("Flg_Status_Ubicacion")
                cboCod_UniMed.SelectedValue = ItemBordadoEstampado("Cod_UniMed")

                If IsDBNull(ItemBordadoEstampado("Fec_Ult_Aprob_Ubicacion")) Then
                    dtpFechaUbicacion.Checked = False
                    dtpFechaUbicacion.Value = Date.Now.Date
                Else
                    dtpFechaUbicacion.Checked = True
                    dtpFechaUbicacion.Value = ItemBordadoEstampado("Fec_Ult_Aprob_Ubicacion")
                End If
                cboCod_ClaItem.SelectedValue = ItemBordadoEstampado("Cod_ClaItem")
                cboModoProceso.SelectedValue = ItemBordadoEstampado("Flg_ModoProceso")
                cboCod_GruItem.SelectedValue = ItemBordadoEstampado("Cod_GruItem")
                cboCod_Origen.SelectedValue = ItemBordadoEstampado("Cod_Origen")
                'cboGradoDif.SelectedValue = ItemBordadoEstampado("Tipo_Grado_Dificultad")
                If ItemBordadoEstampado("etiqueta_est") = "SI" Then
                    rbtSI_Etiq_est.Checked = True
                Else
                    rbtNO_Etiq_est.Checked = True
                End If
                If IsDBNull(ItemBordadoEstampado("Ubicacion")) Then
                    txtUbicacion.Text = ""
                Else
                    txtUbicacion.Text = ItemBordadoEstampado("Ubicacion")
                End If

                If IsDBNull(ItemBordadoEstampado("Comentario")) Then
                    txtComentario.Text = ""
                Else
                    txtComentario.Text = ItemBordadoEstampado("Comentario")
                End If

                txtPrecioComercial.Text = ItemBordadoEstampado("Precio_Cotizacion_Artes")
                txtTecnicaEstampado.Text = ItemBordadoEstampado("Tecnica_Estampado")
                txtDesTecnica.Text = ItemBordadoEstampado("Descripcion_Tecnica")



                If IsDBNull(ItemBordadoEstampado("Dir_Icono")) Then
                Else
                    VisualizarImgen(ItemBordadoEstampado("Dir_Icono"))
                End If

                txtCodProveedor.Text = ItemBordadoEstampado("proveedor")
                txtNombreProveedor.Text = ItemBordadoEstampado("DES_proveedor")
                txtCodItemPro.Text = ItemBordadoEstampado("cod_itemProv")
                txtUMPro.Text = ItemBordadoEstampado("cod_unimedprov")
                txtPrecio.Text = ItemBordadoEstampado("Pre_Cotizado_Proveedor")
                txtObservaciones_Proveedor.Text = ItemBordadoEstampado("Observaciones_Proveedor")
                cboGradoDif.SelectedValue = ItemBordadoEstampado("Tipo_Grado_Dificultad")
                cboIde_Talla.SelectedValue = ItemBordadoEstampado("Ide_Talla")
                cboIde_EsCli.SelectedValue = ItemBordadoEstampado("Ide_EsCli")
                cboIde_Color.SelectedValue = ItemBordadoEstampado("Ide_Color")
                cboIde_Destino.SelectedValue = ItemBordadoEstampado("Ide_Destino")
                CboIde_PO.SelectedValue = ItemBordadoEstampado("Ide_Po")
                cboCod_MotPrePro.SelectedValue = ItemBordadoEstampado("Cod_MotPrePro")

                TxtUMCotizacion.Text = ItemBordadoEstampado("Cod_UniMed_Cotizacion")
                TxtUMDscCotizacion.Text = Hp.DevuelveDato("select Des_UniMed from tg_unimed where COD_UNIMED = '" & TxtUMCotizacion.Text & "'", cconnect)


                TxtNumeroColoresBordado.Text = ItemBordadoEstampado("Num_Colores_Bordado")
                TxtTipoLavado.Text = If(IsDBNull(ItemBordadoEstampado("DEsTipoLavado")), "", ItemBordadoEstampado("DEsTipoLavado"))

                'Txp1.ciTXT_COD.Text = ItemBordadoEstampado("cod_item_pelon")
                'Txp1.ciTXT_DES.Text = Hp.DevuelveDato("select Des_Item from lg_item where Cod_Item = '" & Txp1.ciTXT_COD.Text.TrimEnd & "'", cconnect)
                'NumLargoPieza.Value = ItemBordadoEstampado("largo_pieza_pelon")
                'NumAnchoPieza.Value = ItemBordadoEstampado("ancho_pieza_pelon")
                'NumNroCapas.Value = ItemBordadoEstampado("nro_capas_pelon")
                'NumConsumo.Value = ItemBordadoEstampado("Consumo_Pelon_Bordado")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub VisualizarImgen(ByRef StrRutaImagen As String)
        Try
            Dim img As Image
            If Not String.IsNullOrEmpty(Trim(StrRutaImagen)) Then
                img = Image.FromFile(StrRutaImagen)

                Dim imgCopy As Image = New Bitmap(img)
                img.Dispose()

                PictureBox2.Image = imgCopy
            Else
                PictureBox2.Image = Nothing
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cboCod_FamItem_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboCod_FamItem.SelectedIndexChanged
        Try
            If Conexion.State = ConnectionState.Open Then
                Conexion.Close()
            End If
            cboCod_GruItem.DataSource = Nothing
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub HabilitaDatos()
        Try
            txtcoditem.Enabled = True
            txtDesItem.Enabled = True

            txtComentario.Enabled = True

            If TipoConsulta = 1 Then
                cboCod_FamItem.Enabled = False
            Else
                cboCod_FamItem.Enabled = True
            End If
            cboCod_GruItem.Enabled = True
            cboCod_UniMed.Enabled = True
            cboCod_ClaItem.Enabled = True
            cboFlg_Status.Enabled = True
            cboCod_Origen.Enabled = True
            cboCod_MotPrePro.Enabled = True
            cboIde_Talla.Enabled = True
            cboIde_Color.Enabled = True
            cboIde_EsCli.Enabled = True
            cboIde_Destino.Enabled = True
            CboIde_PO.Enabled = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub InhabilitaDatos()
        Try
            txtcoditem.Enabled = False
            txtDesItem.Enabled = False
            cboCod_FamItem.Enabled = False
            cboFlg_Status.Enabled = False
            cboCod_UniMed.Enabled = False
            cboCod_GruItem.Enabled = False
            cboCod_ClaItem.Enabled = False
            cboCod_Origen.Enabled = False
            cboCod_MotPrePro.Enabled = False
            cboIde_Talla.Enabled = False
            cboIde_Color.Enabled = False
            cboIde_EsCli.Enabled = False
            cboIde_Destino.Enabled = False
            CboIde_PO.Enabled = False
            cboGradoDif.Enabled = False
            txtComentario.Enabled = False
            txtUbicacion.Enabled = False
            dtpFechaUbicacion.Enabled = False
            cboModoProceso.Enabled = False
            Me.txtPrecioComercial.Enabled = False
            Me.txtTecnicaEstampado.Enabled = False
            Me.TxtTipoLavado.Enabled = False
            Me.txtDesTecnica.Enabled = False
            Me.txtCodProveedor.Enabled = False
            Me.txtNombreProveedor.Enabled = False
            Me.txtCodItemPro.Enabled = False
            Me.txtUMPro.Enabled = False
            Me.txtPrecio.Enabled = False
            Me.txtObservaciones_Proveedor.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub LimpiarDatos()
        Try
            txtcoditem.Text = String.Empty
            txtDesItem.Text = String.Empty
            txtComentario.Text = String.Empty
            cboCod_FamItem_SelectedIndexChanged(cboCod_FamItem, New System.EventArgs())
            cboCod_Origen.SelectedValue = "L"
            cboFlg_Status.SelectedValue = "P"
            cboCod_ClaItem.SelectedValue = "P"
            cboCod_UniMed.SelectedIndex = -1
            cboCod_MotPrePro.SelectedIndex = -1
            cboIde_Talla.SelectedIndex = 0
            cboIde_Color.SelectedIndex = 0
            cboIde_EsCli.SelectedIndex = 0
            cboIde_Destino.SelectedIndex = 0
            CboIde_PO.SelectedIndex = 0
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BarraOpciones_ItemClick(sender As Object, e As Janus.Windows.ButtonBar.ItemEventArgs) Handles BarraOpciones.ItemClick
        Try
            Dim sFLAG_Status_Arte As New VB6.FixedLengthString(1)
            Dim ItemBordadoEstampado As DataRow
            ItemBordadoEstampado = ObtenerDr_DeGridEx(GrdEstampadoBordado)
            Dim sestilo_version As String

            Select Case e.Item.Key
                Case "ADICIONAR"
                    'If RbCliente.Checked = False Then
                    '    MessageBox.Show("Debe Ingresar Cliente Tempordada antes de acceder a esta opción !! ", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    '    Exit Sub
                    'End If

                    'If RbCliente.Checked = True And (RTrim(TxpCliente.ciTXT_COD.Text) = "" Or RTrim(TxpTemporada.ciTXT_COD.Text) = "") Then
                    '    MessageBox.Show("Debe Ingresar Cliente /Tempordada ", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    '    Exit Sub
                    'End If

                    Dim frm As New frmAdicionarModificarItems

                    With frm
                        .oParent = Me
                        .Text = "Adicionar Item"
                        .Abr_Cliente = Trim(TxpCliente.ciTXT_COD.Text)
                        .sTemporada = Trim(TxpTemporada.ciTXT_COD.Text)
                        .Opcion = CStr(TipoConsulta)
                        .sTipo = "I"
                        .txtCodUM.Text = "UN"
                        .txtCodClase.Text = "P"
                        .txtCodMotivo.Text = "PD"
                        .txtCodOrigen.Text = "L"
                        .BtnAdicionaTecnica.Visible = False
                        .ShowDialog()
                    End With
                Case "MODIFICAR"
                    If GrdEstampadoBordado.RowCount > 0 Then

                        Dim frm As New frmAdicionarModificarItems

                        frm.oParent = Me
                        frm.Text = "Modificar Item"
                        frm.Abr_Cliente = TxpCliente.ciTXT_COD.Text
                        frm.sTemporada = TxpTemporada.ciTXT_COD.Text
                        frm.Opcion = CStr(TipoConsulta)
                        frm.sTipo = "U"


                        If ItemBordadoEstampado("etiqueta_est") = "SI" Then
                            frm.sEstampado = "SI"
                        Else
                            frm.sEstampado = "NO"
                        End If
                        frm.txtcoditem.Text = ItemBordadoEstampado("Cod_Itemx")
                        frm.txtDesItem.Text = ItemBordadoEstampado("Des_Item")
                        frm.txtCodFamilia.Text = ItemBordadoEstampado("cod_FamItem")
                        frm.txtDesFamilia.Text = ItemBordadoEstampado("des_famitem")
                        frm.txtCodUM.Text = ItemBordadoEstampado("cod_UniMed")
                        frm.txtDesUM.Text = ItemBordadoEstampado("Des_UniMed")
                        frm.txtCodClase.Text = ItemBordadoEstampado("cod_ClaItem")
                        frm.txtDesClase.Text = ItemBordadoEstampado("des_claitem")
                        frm.txtCodGrupo.Text = ItemBordadoEstampado("cod_GruItem")
                        frm.txtDesGrupo.Text = ItemBordadoEstampado("des_famgruite")
                        frm.txtCodStatus.Text = ItemBordadoEstampado("Flg_Status")
                        frm.txtDesStatus.Text = ItemBordadoEstampado("des_status")
                        frm.txtCodTipoVersion.Text = ItemBordadoEstampado("Tip_version")
                        'frmAdicionarModificarItems.txtDesTipoVersion = RTrim(DGridLista.Value(DGridLista.Columns("Descripcion").Index))
                        frm.TxtModo.Text = ItemBordadoEstampado("Flg_ModoProceso")
                        frm.TxtDes_modo.Text = ItemBordadoEstampado("Des_ModoProceso")
                        frm.txtCodMotivo.Text = ItemBordadoEstampado("Cod_MotPrePro")
                        frm.txtDesMotivo.Text = ItemBordadoEstampado("des_motprepro")
                        frm.txtCodOrigen.Text = ItemBordadoEstampado("cod_origen")
                        frm.txtDesOrigen.Text = ItemBordadoEstampado("des_origen")
                        frm.txtUbicacion.Text = ItemBordadoEstampado("Ubicacion")
                        frm.txtComentario.Text = ItemBordadoEstampado("Comentario")
                        frm.txtCodProveedor.Text = ItemBordadoEstampado("proveedor")
                        frm.txtNombreProveedor.Text = ItemBordadoEstampado("des_proveedor")
                        frm.txtPrecio.Text = ItemBordadoEstampado("Pre_cotizado_proveedor")
                        frm.txtObservacionesProv.Text = ItemBordadoEstampado("Observaciones_proveedor")
                        frm.txtCodItemProv.Text = ItemBordadoEstampado("cod_itemProv")
                        frm.txtUniMedProv.Text = ItemBordadoEstampado("cod_unimedprov")
                        

                        If IsDBNull(ItemBordadoEstampado("Dir_Icono")) Then
                            frm.txtDirIcono.Text = ""
                            frm.strImagenCambio = ""
                        Else
                            frm.txtDirIcono.Text = ItemBordadoEstampado("Dir_Icono")
                            frm.strImagenCambio = ItemBordadoEstampado("Dir_Icono")
                        End If

                        frm.txtCodTecnica.Text = ItemBordadoEstampado("cod_tecnica")
                        frm.TxtDesTecnica.Text = ItemBordadoEstampado("descripcion_tecnica")
                        frm.txtPrecioComercial.Text = ItemBordadoEstampado("Precio_Cotizacion_Artes")
                        frm.txtTecnicaEstampado.Text = ItemBordadoEstampado("Tecnica_Estampado")
                        frm.TxtCaracteristica_Tela.Text = ItemBordadoEstampado("Caracteristica_Tela_Estampados")
                        frm.TxpUnidadMedida.ciTXT_COD.Text = ItemBordadoEstampado("Cod_UniMed_Cotizacion")
                        frm.TxtNumeroColoresBordado.Text = ItemBordadoEstampado("Num_Colores_Bordado")
                        frm.TxtCodTipoLavado.Text = ItemBordadoEstampado("TipoLavado")
                        frm.TxtDesTipoLavado.Text = IIf(IsDBNull(ItemBordadoEstampado("DEsTipoLavado")), "", ItemBordadoEstampado("DEsTipoLavado"))

                        If ItemBordadoEstampado("cod_FamItem") = "BD" Or ItemBordadoEstampado("cod_FamItem") = "AB" Then
                            frm.lblNumeroColoresBordado.Visible = True
                            frm.TxtNumeroColoresBordado.Visible = True
                        Else
                            frm.lblNumeroColoresBordado.Visible = False
                            frm.TxtNumeroColoresBordado.Visible = False
                        End If

                        frm.TxpUnidadMedida.ciTXT_DES.Text = Hp.DevuelveDato("select Des_UniMed from tg_unimed where COD_UNIMED = '" & frm.TxpUnidadMedida.ciTXT_COD.Text & "'", cconnect)

                        frm.txtGradoDif.Text = ItemBordadoEstampado("Tipo_Grado_Dificultad")
                        frm.txtGradodifDes.Text = cboGradoDif.Text
                        'frm.Busca_TipGradoDif((1))

                        If RTrim(FixNulos((frm.DtpFec_Prevista_Aprobacion.Value), VariantType.String)) = "" Then
                            ''frmAdicionarModificarItems.DtpFec_Prevista_Aprobacion.CheckBox = False
                        Else
                            'frmAdicionarModificarItems.DtpFec_Prevista_Aprobacion.CheckBox = True

                        End If

                        Call BuscaCombo(ItemBordadoEstampado("Ide_Talla"), 2, (frm.cboIde_TallaX))
                        Call BuscaCombo(ItemBordadoEstampado("Ide_Color"), 2, (frm.cboIde_Color))
                        Call BuscaCombo(ItemBordadoEstampado("Ide_EsCli"), 2, (frm.cboIde_EsCli))

                        If Not IsDBNull(ItemBordadoEstampado("Ide_Destino")) Then

                            Call BuscaCombo(ItemBordadoEstampado("Ide_Destino"), 2, (frm.cboIde_Destino))
                        Else
                            frm.cboIde_Destino.SelectedIndex = -1
                        End If

                        If Not IsDBNull(ItemBordadoEstampado("Ide_Po")) Then
                            Call BuscaCombo(ItemBordadoEstampado("Ide_Po"), 2, (frm.CboIde_PO))
                        Else
                            frm.CboIde_PO.SelectedIndex = -1
                        End If

                        sFLAG_Status_Arte.Value = Trim(ItemBordadoEstampado("FLG_APROBACION_UBICACION_ARTE"))
                        If sFLAG_Status_Arte.Value = "S" Then
                            frm.optSI.Checked = True
                            frm.optNO.Checked = False
                        Else
                            frm.optSI.Checked = False
                            frm.optNO.Checked = True
                        End If

                        frm.Frame3.Enabled = False

                        If Not String.IsNullOrEmpty(cboCod_FamItem.Text) Then

                            If cboCod_FamItem.SelectedValue.ToString = "BD" Or cboCod_FamItem.SelectedValue.ToString = "AB" Then
                                frm.nUpLargo_Bordado.Enabled = True
                                frm.nUpAncho_Bordado.Enabled = True
                                frm.nUpNumero_de_Puntadas_Bordado.Enabled = True
                                frm.nUpTiempo_Maquina_Bordado.Enabled = True
                                frm.nUpTiempo_Limpieza_Bordado.Enabled = True
                            End If
                        End If

                        frm.nUpLargo_Bordado.Text = ItemBordadoEstampado("Largo_Bordado")
                        frm.nUpAncho_Bordado.Text = ItemBordadoEstampado("Ancho_Bordado")
                        frm.nUpNumero_de_Puntadas_Bordado.Value = ItemBordadoEstampado("Número_de_Puntadas_Bordado")
                        frm.nUpTiempo_Maquina_Bordado.Value = ItemBordadoEstampado("Tiempo_Máquina_Bordado")
                        frm.nUpTiempo_Limpieza_Bordado.Value = ItemBordadoEstampado("Tiempo_Limpieza_Bordado")
                        frm.BtnAdicionaTecnica.Visible = True

                        If Not String.IsNullOrEmpty(cboCod_FamItem.Text) Then
                            If cboCod_FamItem.SelectedValue.ToString = "BD" Or cboCod_FamItem.SelectedValue.ToString = "AB" Then
                                frm.PnBordado.Enabled = True
                            End If
                        End If

                        'frm.Txp1.ciTXT_COD.Text = ItemBordadoEstampado("cod_item_pelon")
                        'frm.Txp1.ciTXT_DES.Text = Hp.DevuelveDato("select Des_Item from lg_item where Cod_Item = '" & Txp1.ciTXT_COD.Text.TrimEnd & "'", cconnect)
                        'frm.NumLargoPieza.Value = ItemBordadoEstampado("largo_pieza_pelon")
                        'frm.NumAnchoPieza.Value = ItemBordadoEstampado("ancho_pieza_pelon")
                        'frm.NumNroCapas.Value = ItemBordadoEstampado("nro_capas_pelon")
                        'frm.NumConsumo.Value = ItemBordadoEstampado("Consumo_Pelon_Bordado")
                        frm.ShowDialog()

                    Else
                        MessageBox.Show("Debe seleccionar un item para acceder a esta opcion", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If

                Case "ELIMINAR"
                    If GrdEstampadoBordado.RowCount > 0 Then
                        If (MessageBox.Show("¿Esta seguro de eliminar el registro?", "Eliminar", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.Yes) Then
                            EliminarItem()
                        End If
                    Else
                        MessageBox.Show("Debe seleccionar un item para acceder a esta opcion", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If

                Case "PELONES"
                    If GrdEstampadoBordado.RowCount > 0 Then
                        If cboCod_FamItem.SelectedValue.ToString = "BD" Or cboCod_FamItem.SelectedValue.ToString = "AB" Then
                            Using oPel As New frmMantItemsServicios_Pelones
                                With oPel
                                    .TxtCodItem.Text = ItemBordadoEstampado("cod_item")
                                    .TxtDesItem.Text = ItemBordadoEstampado("Des_Item")
                                    .ShowDialog()
                                End With
                            End Using
                        Else
                            MessageBox.Show("Opcion solo para Bordados", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    End If

                Case "LEADTIME"
                    If GrdEstampadoBordado.RowCount > 0 Then
                        Dim i As Integer
                        i = GrdEstampadoBordado.Row
                        Dim frm As New Lead_Tima_de_Abastecimiento()
                        frm.sCodItem = ItemBordadoEstampado("cod_item")
                        frm.sDcsItem = ItemBordadoEstampado("Des_Item")
                        frm.sLTAbastecimiento = ItemBordadoEstampado("Lead_Time_Abastecimiento")
                        frm.sLTTravesia = ItemBordadoEstampado("Lead_Time_Travesia_Desaduanaje")
                        If frm.ShowDialog() = DialogResult.OK Then
                            Call CargaLista()
                            GrdEstampadoBordado.Row = i
                        End If
                    End If
                Case "ADICIONARTECNICA"
                    If GrdEstampadoBordado.RowCount > 0 Then
                        Dim i As Integer
                        i = GrdEstampadoBordado.Row
                        Dim frm As New FrmAdicionarTecnica()
                        frm.vCodItem = ItemBordadoEstampado("cod_item")
                        frm.vDscItem = ItemBordadoEstampado("Des_Item")
                        If frm.ShowDialog() = DialogResult.OK Then

                            GrdEstampadoBordado.Row = i
                        End If
                    End If
                Case "CAMBIOESTADO"
                    If GrdEstampadoBordado.RowCount > 0 Then

                        If Trim(ItemBordadoEstampado("FLG_STATUS_UBICACION")) <> "P" Then
                            If MsgBox("Esta seguro de cambiar de estado", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "AVISO") = MsgBoxResult.Yes Then
                                Dim strSQL As String
                                strSQL = " exec ES_Cambia_Status_Ubicacion '" & ItemBordadoEstampado("Cod_Itemx") & "','P' "
                                Call ExecuteSQL(cconnectVB6, strSQL)
                                Call CargaLista()
                            End If
                        Else
                            Dim frm As New frmCambioEstadoItems(ItemBordadoEstampado("Cod_Itemx"), ItemBordadoEstampado("FLG_STATUS_UBICACION"))
                            If frm.ShowDialog() = DialogResult.OK Then
                                Call CargaLista()
                            End If
                        End If
                    Else
                        MessageBox.Show("Debe seleccionar un item para acceder a esta opcion", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Case "IMPRESION"
                        If GrdEstampadoBordado.RowCount > 0 Then
                            If RbCliente.Checked Then
                                Dim frm As New frmImprimirCBEA(CodigoCliente, TxpTemporada.ciTXT_COD.Text)
                                If frm.ShowDialog() = DialogResult.OK Then
                                    Call CargaLista()
                                End If
                            Else
                                MessageBox.Show("Debe ingresar Cliente Temporada", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            End If
                        End If
                Case "COMBINACIONES"
                    If cboIde_Color.SelectedValue = "S" Or CboIde_PO.SelectedValue = "S" Then Exit Sub
                    If GrdEstampadoBordado.RowCount > 0 Then
                        Using oComb As New frmMantItemComb
                            With oComb
                                .Text = "COMBINACIONES DE ITEM:" & ItemBordadoEstampado("Cod_Item") & " " & ItemBordadoEstampado("Des_Item")
                                .Codigo_item = ItemBordadoEstampado("Cod_Item")
                                .TxtCodItem.Text = ItemBordadoEstampado("Cod_Item")
                                .TxtDesItem.Text = ItemBordadoEstampado("Des_Item")
                                .codCliente = ItemBordadoEstampado("cod_cliente")
                                .codTemporada = ItemBordadoEstampado("cod_temcli")
                                .sopcion = TipoConsulta
                                .CARGA_GRID()
                                .ShowDialog()
                            End With
                        End Using
                    Else
                        MessageBox.Show("Debe seleccionar un Item para acceder a esta opcion", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If

                Case "PROVEEDOR"
                    If GrdEstampadoBordado.RowCount > 0 Then
                        Dim frm As New frmManItemProvShort
                        frm.sUniMedDefault = ItemBordadoEstampado("cod_unimed")
                        frm.varCod_item = ItemBordadoEstampado("cod_item")
                        frm.varCod_Proveedor = ItemBordadoEstampado("Cod_Proveedor")
                        frm.Text = "Item Proveedor  Item :" & ItemBordadoEstampado("cod_item")
                        frm.CARGA_GRID()
                        If frm.ShowDialog() = DialogResult.OK Then
                            CargaLista()
                        End If
                    Else
                        MessageBox.Show("Debe seleccionar un Item para acceder a esta opcion", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Case "TEMPORADA"
                    If TxpCliente.ciTXT_COD.Text <> String.Empty And TxpTemporada.ciTXT_COD.Text <> String.Empty Then
                        Dim frm As New frmAdItemTemCli
                        frm.sCod_Cliente = CodigoCliente
                        frm.oParent = Me
                        frm.sCod_Temcli = TxpTemporada.ciTXT_COD.Text
                        frm.ShowDialog()
                    Else
                        MessageBox.Show("Debe ingresar Cliente Temporada", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Case "VERESTILOSNP"
                    If GrdEstampadoBordado.RowCount > 0 Then
                        Dim frm As New frmEstilosNPs
                        frm.codCliente = ItemBordadoEstampado("cod_cliente")
                        frm.codTemporada = ItemBordadoEstampado("cod_temcli")
                        frm.codItem = ItemBordadoEstampado("cod_item")
                        frm.ShowDialog()
                    End If
                Case "ELIMDETEMP"
                        If GrdEstampadoBordado.RowCount > 0 Then
                            If MsgBox("Esta seguro de eliminar el item " & ItemBordadoEstampado("COD_ITEM") & " de esta tempodada", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "AVISO") = MsgBoxResult.Yes Then
                                EliminarItemTemporada()
                            End If
                        Else
                            MessageBox.Show("Debe seleccionar un Item para acceder a esta opcion", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                Case "MODPRECOM"
                        If GrdEstampadoBordado.RowCount = 0 Then Exit Sub
                        famPrecioComercial.Visible = True
                        txtPC_item.Text = ItemBordadoEstampado("Cod_Itemx")
                        txtPC_Item_Des.Text = ItemBordadoEstampado("Des_Item")
                        txtPC_Pecio.Text = ItemBordadoEstampado("Precio_Cotizacion_Artes")
                        txtPC_Pecio.Focus()
                Case "IMPTICKET"

                        If GrdEstampadoBordado.RowCount > 0 Then
                            With FrmImpTicket

                                .V1 = ItemBordadoEstampado("Cod_Itemx")
                                .ShowDialog()
                            End With
                        End If
                Case "COMPOSICION"
                    If GrdEstampadoBordado.RowCount > 0 Then
                        Dim frm As New frmMantHilosItem
                        frm.Codigo_item = ItemBordadoEstampado("Cod_Itemx")
                        frm.txtDes_Item.Text = ItemBordadoEstampado("Des_Item")
                        frm.CARGA_GRID()
                        frm.ShowDialog()
                    Else
                        MsgBox("Debe seleccionar un Item para acceder a esta opcion")
                    End If
                Case "DESARROLLOLAV"

                    '    Dim strSQL As String
                    'If GrdEstampadoBordado.RowCount > 0 Then


                    '    strSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(TxpCliente.ciTXT_COD.Text) & "'"

                    '    Dim frm As New frmSolicitudesLavanderia

                    '    With frm
                    '        .Cod_Item = ItemBordadoEstampado("cod_item")
                    '        .COD_PROVEEDOR = ItemBordadoEstampado("Cod_Proveedor")


                    '        .txtFND_ClienteCod.Text = Trim(TxpCliente.ciTXT_COD.Text)
                    '        .txtFND_ClienteDes.Text = UCase(TxpCliente.ciTXT_DES.Text)

                    '        .txtFND_TemporadaCod.Text = UCase(TxpTemporada.ciTXT_COD.Text)


                    '        .txtFND_TemporadaDes.Text = UCase(TxpTemporada.ciTXT_DES.Text)

                    '        .txtFND_CodItem.Text = ItemBordadoEstampado("cod_item")

                    '        .scod_Item2 = ItemBordadoEstampado("cod_item")
                    '        .VAR_CLIENTE = Trim(TxpCliente.ciTXT_COD.Text)

                    '        .COD_CLIENTE = DevuelveCampo(strSQL, cconnectVB6)
                    '        .NOM_CLIENTE = UCase(TxpCliente.ciTXT_DES.Text)

                    '        .VAR_TEMPORADA = UCase(TxpTemporada.ciTXT_COD.Text)
                    '        .NOM_TEMPORADA = UCase(TxpCliente.ciTXT_DES.Text)

                    '        .txtcliente.Text = UCase(TxpCliente.ciTXT_COD.Text)
                    '        .txtNom_Cliente.Text = UCase(TxpCliente.ciTXT_DES.Text)

                    '        .txttemporada.Text = UCase(TxpTemporada.ciTXT_COD.Text)
                    '        .txtNom_TemCli.Text = UCase(TxpTemporada.ciTXT_DES.Text)

                    '        .txtcliente.Enabled = False
                    '        .txtNom_Cliente.Enabled = False
                    '        .txttemporada.Enabled = False
                    '        .txtNom_TemCli.Enabled = False

                    '        .CARGA_GRID()
                    '        .ShowDialog()
                    '    End With
                    '    CargaLista()
                    'Else
                    '    MsgBox("Debe seleccionar un Item para acceder a esta opcion")
                    'End If
                Case "SALIR"
                        Me.Close()
                Case "PRODUCCION"
                        Dim strSQL As String
                    If GrdEstampadoBordado.RowCount > 0 Then

                        strSQL = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(TxpCliente.ciTXT_COD.Text) & "'"
                        Dim frm As New frmSolicitudesProduccion

                        With frm
                            .COD_ITEM = ItemBordadoEstampado("cod_item")
                            .COD_PROVEEDOR = ItemBordadoEstampado("Cod_Proveedor")
                            .txtFND_ClienteCod.Text = UCase(Trim(TxpCliente.ciTXT_COD.Text))
                            .txtFND_ClienteDes.Text = UCase(Trim(TxpCliente.ciTXT_DES.Text))

                            .txtFND_TemporadaCod.Text = UCase(TxpTemporada.ciTXT_COD.Text)
                            .txtFND_TemporadaDes.Text = UCase(TxpTemporada.ciTXT_DES.Text)
                            .txtFND_CodItem.Text = ItemBordadoEstampado("cod_item")
                            .scod_Item2 = ItemBordadoEstampado("cod_item")

                            .VAR_CLIENTE = Trim(TxpCliente.ciTXT_COD.Text)

                            .COD_CLIENTE = DevuelveCampo(strSQL, cconnectVB6)
                            .NOM_CLIENTE = UCase(Trim(TxpCliente.ciTXT_DES.Text))

                            .VAR_TEMPORADA = UCase(TxpTemporada.ciTXT_COD.Text)
                            .NOM_TEMPORADA = UCase(TxpCliente.ciTXT_DES.Text)

                            .txtcliente.Text = UCase(TxpCliente.ciTXT_COD.Text)
                            .txtNom_Cliente.Text = UCase(Trim(TxpCliente.ciTXT_DES.Text))

                            .txttemporada.Text = UCase(TxpTemporada.ciTXT_COD.Text)
                            .txtNom_TemCli.Text = UCase(TxpTemporada.ciTXT_DES.Text)

                            .txtcliente.Enabled = False
                            .txtNom_Cliente.Enabled = False
                            .txttemporada.Enabled = False
                            .txtNom_TemCli.Enabled = False

                            .CARGA_GRID()

                            .txtModoProceso_Cod.Text = UCase(Trim(RTrim(ItemBordadoEstampado("Cod_MotPrePro"))))
                            .txtModoProceso_Des.Text = UCase(Trim(RTrim(ItemBordadoEstampado("Des_ModoProceso"))))
                            .TxtCod_Tecnica.Text = UCase(Trim(RTrim(ItemBordadoEstampado("cod_tecnica"))))
                            .TxtDes_Tecnica.Text = UCase(Trim(RTrim(ItemBordadoEstampado("descripcion_tecnica"))))

                            .ShowDialog()
                        End With
                        CargaLista()
                    Else
                        MsgBox("Debe seleccionar un Item para acceder a esta opcion")
                    End If

                Case "IMPUBIAR"
                    If GrdEstampadoBordado.RowCount > 0 Then

                        sestilo_version = ""
                        Dim frm As New frmEstilosNPs2

                        frm.codCliente = ItemBordadoEstampado("cod_cliente")
                        frm.codTemporada = ItemBordadoEstampado("cod_temcli")
                        frm.codItem = ItemBordadoEstampado("cod_item")
                        frm.ShowDialog()
                    End If
                        Call Imprimir_Ubicacion_Arte(sestilo_version)
                Case "IMPRISTRIKE"
                        Reporte_Pendiente()
                Case "IMPRILAV"
                        Reporte_Pendiente_LAVADO()
                Case "STRIKEREQ"
                        FRA_STRIKE_REQ.Visible = True
                        DtpFec_Inicio.Value = Today
                        DtpFec_Fin.Value = Today


                Case "MODIFICARGRADODIF"
                    Dim frm As New frmMantItemServiciosModGradDif
                    frm.oParent = Me
                    frm.sCod_Item = ItemBordadoEstampado("cod_itemx")
                    frm.txtGradoDif.Text = ItemBordadoEstampado("tipo_grado_dificultad")
                    frm.Busca_TipGradoDif(1)
                    frm.ShowDialog()
                    CargaLista()
                Case "IMPRIMIRLISTA"
                    IMPRIMIR_LISTA()
                Case "MEDIDA"
                    If cboIde_Talla.SelectedValue.ToString = "S" Then Exit Sub
                    Dim frm As New FrmMantMed
                    frm.Cod_Item = ItemBordadoEstampado("Cod_item")
                    frm.Tipo_Item = "I"
                    frm.vMostarPanel = True
                    frm.vPnMedDspLav = True
                    frm.ShowDialog()
                Case "DATOTECNICO"
                    Dim Frmx As New FrmMantItemsServicos_DatosTecnicos
                    Frmx.vlCod_Item = ItemBordadoEstampado("Cod_item")
                    Frmx.vlDesItem = ItemBordadoEstampado("Des_Item")
                    Frmx.ShowDialog()
                    'cod_item

                Case "ESTMEDIDAS"
                    If GrdEstampadoBordado.RowCount = 0 Then Return
                    Using oEncog As New FrmMantItemsServicos_EstMedidaEncog
                        With oEncog
                            .TxtCodItem.Text = ItemBordadoEstampado("Cod_Item")
                            .TxtDesItem.Text = ItemBordadoEstampado("Des_Item")
                            .ShowDialog()
                        End With
                    End Using
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub EliminarItem()
        Try
            Dim strSQL As String
            Dim ItemBordadoEstampado As DataRow
            ItemBordadoEstampado = ObtenerDr_DeGridEx(GrdEstampadoBordado)

            strSQL = "UP_MAN_ITEMS2 " & TipoConsulta & ",'D','" & ItemBordadoEstampado("cod_itemx") & "','" & ItemBordadoEstampado("cod_FamItem") & _
                     "','" & ItemBordadoEstampado("cod_GruItem") & "','" & ItemBordadoEstampado("cod_UniMed") & "','" & ItemBordadoEstampado("Des_Item") & _
                     "','" & ItemBordadoEstampado("cod_ClaItem") & "','" & ItemBordadoEstampado("cod_origen") & "','" & ItemBordadoEstampado("Ide_Talla") & _
                     "','" & ItemBordadoEstampado("Ide_Color") & "','" & ItemBordadoEstampado("Ide_EsCli") & "','" & ItemBordadoEstampado("Ide_Destino") & _
                     "','" & ItemBordadoEstampado("Cod_MotPrePro") & "','" & ItemBordadoEstampado("cod_cliente") & "','" & ItemBordadoEstampado("cod_temcli") & _
                     "','" & ItemBordadoEstampado("Comentario") & "','" & ItemBordadoEstampado("Ide_Po") & "','" & vusu & "','" & ItemBordadoEstampado("Ubicacion") & _
                     "','" & ItemBordadoEstampado("Flg_Status") & "','" & ItemBordadoEstampado("Tip_version") & "','" & ItemBordadoEstampado("Flg_ModoProceso") & _
                     "','','" & ItemBordadoEstampado("proveedor") & "','" & ItemBordadoEstampado("cod_itemprov") & "','" & ItemBordadoEstampado("cod_unimedprov") & _
                     "','" & ItemBordadoEstampado("Pre_cotizado_proveedor") & "','" & ItemBordadoEstampado("Observaciones_proveedor") & "'"

            Call ExecuteSQL(cconnectVB6, strSQL)
            Call CargaLista()

            Exit Sub
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub EliminarItemTemporada()
        Try
            Dim strSQL As String
            Dim ItemBordadoEstampado As DataRow
            ItemBordadoEstampado = ObtenerDr_DeGridEx(GrdEstampadoBordado)

            strSQL = "LG_ELIMINA_ITEM_TEMPORADA '" & ItemBordadoEstampado("cod_itemx") & "','" & ItemBordadoEstampado("cod_cliente") & "','" & ItemBordadoEstampado("cod_temcli") & "'"

            Call ExecuteSQL(cconnectVB6, strSQL)
            Call CargaLista()

            Exit Sub
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub IMPRIMIR_LISTA()
        Try
            Dim oo As Object
            Dim strSQL As String

            Dim Adors1 As New ADODB.Recordset
            '  Dim rutaLogo As String


            strSQL = "EXEC ES_SM_ItemServicios_ClienteTemp '" & TipoConsulta & "','" & CodigoCliente & "','" & TxpTemporada.ciTXT_COD.Text & "','" & TxpItems.ciTXT_COD.Text & "','" & TxpEstado.ciTXT_COD.Text & "', '" & TxpProveedor.ciTXT_COD.Text & "','" & TxtFamilia.Text & "' "

            Adors1 = CargarRecordSetDesconectado(strSQL, cconnectVB6)

            If Adors1.RecordCount > 0 Then
                oo = CreateObject("Excel.Application")

                oo.workbooks.Open(vruta & "\rptControlBordadosEstampados.XLT")

                oo.Visible = True

                oo.DisplayAlerts = False

                oo.run("Reporte", Adors1)

                oo = Nothing
            Else
                MsgBox("No hay datos para mostrar")
            End If
            Exit Sub

            oo = Nothing
            '   MsgBox("Hubo error en la impresion del Reporte" & Err.Description, MsgBoxStyle.Critical, "Impresion")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TxpCliente_TxpEvent_FalloAlConsultarSQL(ctrlTXP As WinFormsControls.ucPROMPT_BASE, xERROR_SQL As SqlException) Handles TxpCliente.TxpEvent_FalloAlConsultarSQL

    End Sub

    Private Sub cmdAceptar_Click(sender As Object, e As EventArgs) Handles cmdAceptar.Click
        Dim strSQL As String
        Try
            strSQL = "EXEC ES_ACTUALIZA_PRECIO_COMERCIAL_ARTES '" & txtPC_item.Text & "', " & CStr(Val(txtPC_Pecio.Text)) & ",'" & vusu & "','" & ComputerName() & "'"
            Call ExecuteSQL(cconnectVB6, strSQL)
            MsgBox("EL Precio Comercial ha sido actualizado por el indicado" & vbNewLine & "de manera satisfactoria......", MsgBoxStyle.Information, Me.Text)

            'Call CmdCancelar_Click(cmdCancelar, New System.EventArgs())
            'Call FunctBuscar_ActionClick(FunctBuscar, New AxFunctionsButtons.__FunctButt_ActionClickEvent(0, 0, ""))

            Dim sCLAVE As Int32

            sCLAVE = GrdEstampadoBordado.Row

            CargaLista()
            GrdEstampadoBordado.FirstRow.Equals(sCLAVE)
            Salir()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


        

        
    End Sub

    Private Sub cmdCancelar_Click(sender As Object, e As EventArgs) Handles cmdCancelar.Click
        Salir()
    End Sub

    Sub Salir()
        Try
            txtPC_item.Text = ""
            txtPC_Pecio.Text = ""
            famPrecioComercial.Visible = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub Imprimir_Ubicacion_Arte(ByRef est_ver As String)
        Try
            Dim ItemBordadoEstampado As DataRow
            ItemBordadoEstampado = ObtenerDr_DeGridEx(GrdEstampadoBordado)

            If RbCliente.Checked = False Then
                MsgBox("Solo se puede imprimir si la opcion CLIENTE esta marcada......", MsgBoxStyle.Exclamation)
                Exit Sub
            End If

            Dim oRs As New ADODB.Recordset
            Dim sCod_Item, strSQL, sCod_Cliente As String

            sCod_Item = ItemBordadoEstampado("cod_itemx")
            sCod_Cliente = DevuelveCampo("SELECT COD_CLIENTE FROM TG_CLIENTE WHERE ABR_CLIENTE = '" & Trim(TxpCliente.ciTXT_COD.Text) & "'", cconnectVB6)

            strSQL = "EXEC es_reporte_ubicacion_artes_cabecera '" & sCod_Cliente & "', '" & TxpTemporada.ciTXT_COD.Text & "', '" & sCod_Item & "','" & est_ver & "'"
            oRs = CargarRecordSetDesconectado(strSQL, cconnectVB6)
            If oRs.RecordCount = 0 Then
                MsgBox("No se han encontrado datos para la impresión.....", MsgBoxStyle.Exclamation)
                Exit Sub
            End If

            Dim oo As Object
            Dim sCodItem, sRutaLogo, sTitulo, sDesItem As String

            oo = CreateObject("excel.application")
            strSQL = "SELECT Ruta_Logo = ISNULL(Ruta_Logo, '') From SEGURIDAD..SEG_EMPRESAS WHERE Cod_Empresa = '" & vemp & "'"
            sRutaLogo = DevuelveCampo(strSQL, cconnectVB6)
            oo.workbooks.Open(vruta & "\rptUbicacionArte.XLT")
            oo.Visible = True
            oo.DisplayAlerts = False
            oo.run("reporte", sRutaLogo, oRs)
            oo = Nothing
            Exit Sub

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Reporte_Pendiente()
        Try
            Dim oo As Object
            Dim vRutaLogo As String
            Dim strSQL As String
            Dim oRs As New ADODB.Recordset


            strSQL = "LG_MUESTRA_item_solicitud_aplicaciones '1', 0, '','',''"


            'oRs = GetRecordset(cconnectVB6, strSQL)
            oRs = CargarRecordSetDesconectado(strSQL, cconnectVB6)

            'CargarRecordSetDesconectado()

            If oRs.RecordCount = 0 Then
                MsgBox("No Existen Registros")
                Exit Sub
            End If

            strSQL = "SELECT Ruta_Logo From SEGURIDAD..SEG_EMPRESAS " & "WHERE Cod_Empresa = '" & vemp & "'"
            vRutaLogo = DevuelveCampo(strSQL, cconnectVB6)

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            oo = CreateObject("excel.application")

            oo.workbooks.Open(vruta & "\rptStrikeOff_Listado.XLT")

            oo.DisplayAlerts = False
            oo.Visible = True

            Dim sOpcion As String
            Dim sCliente As String
            Dim sTemporada As String
            Dim sNumSolicitud As String
            Dim sCodItem As String

            sOpcion = "TODOS LOS PENDIENTES"
            sCliente = ""
            sTemporada = ""
            sNumSolicitud = ""
            sCodItem = ""

            oo.run("reporte", vRutaLogo, oRs, sOpcion, sCliente, sTemporada, sNumSolicitud, sCodItem)

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            oo = Nothing
            Exit Sub
Fin:
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            MsgBox(Err.Description)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub Reporte_Pendiente_LAVADO()
        Try
            Dim oo As Object
            Dim vRutaLogo As String
            Dim strSQL As String
            Dim oRs As ADODB.Recordset
            strSQL = "LG_MUESTRA_item_solicitud_Desarrollos_Lavados '1', 0, '','',''"

            'oRs = CargarRecordSetDesconectado(cconnectVB6, strSQL)

            oRs = CargarRecordSetDesconectado(strSQL, cconnectVB6)


            If oRs.RecordCount = 0 Then
                MsgBox("No Existen Registros")
                Exit Sub
            End If

            strSQL = "SELECT Ruta_Logo From SEGURIDAD..SEG_EMPRESAS " & "WHERE Cod_Empresa = '" & vemp & "'"
            vRutaLogo = DevuelveCampo(strSQL, cconnectVB6)

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            oo = CreateObject("excel.application")

            oo.workbooks.Open(vruta & "\rptDesarrollo_Lavado_Listado.XLT")

            oo.DisplayAlerts = False
            oo.Visible = True

            Dim sOpcion As String
            Dim sCliente As String
            Dim sTemporada As String
            Dim sNumSolicitud As String
            Dim sCodItem As String

            sOpcion = "TODOS LOS PENDIENTES"
            sCliente = ""
            sTemporada = ""
            sNumSolicitud = ""
            sCodItem = ""

            oo.run("reporte", vRutaLogo, oRs, sOpcion, sCliente, sTemporada, sNumSolicitud, sCodItem)

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            oo = Nothing
            Exit Sub
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Reporte_strike_requerido()
        Try
            Dim oo As Object
            Dim vRutaLogo As String
            Dim strSQL As String
            Dim oRs As ADODB.Recordset


            strSQL = "LG_MUESTRA_item_solicitud_aplicaciones '5', 0, '','','','" & DtpFec_Inicio.Value & "','" & DtpFec_Fin.Value & "'"


            oRs = GetRecordset(cconnectVB6, strSQL)

            If oRs.RecordCount = 0 Then
                MsgBox("No Existen Registros")
                Exit Sub
            End If

            strSQL = "SELECT Ruta_Logo From SEGURIDAD..SEG_EMPRESAS " & "WHERE Cod_Empresa = '" & vemp & "'"

            vRutaLogo = DevuelveCampo(strSQL, cconnectVB6)

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            oo = CreateObject("excel.application")

            oo.workbooks.Open(vruta & "\rptStrikeOff_Listado.XLT")

            oo.DisplayAlerts = False
            oo.Visible = True

            Dim sOpcion As String
            Dim sCliente As String
            Dim sTemporada As String
            Dim sNumSolicitud As String
            Dim sCodItem As String

            sOpcion = "TODOS LOS PENDIENTES"
            sCliente = ""
            sTemporada = ""
            sNumSolicitud = ""
            sCodItem = ""
            oo.run("reporte", vRutaLogo, oRs, sOpcion, sCliente, sTemporada, sNumSolicitud, sCodItem)

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            oo = Nothing
            Exit Sub
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub rbtEstilo_CheckedChanged(sender As Object, e As EventArgs) Handles rbtEstilo.CheckedChanged
        Try
            TipoConsulta = 6
            Inhabilita()
            grbEstiloCliente.Visible = True
            txtEstilo.Focus()
            txtEstilo.Text = String.Empty

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub














    ' === Deep-link desde la bandeja por COD_CLIENTE ===
    Public Sub DeepLinkPorItemPorCodCliente(codCliente As String,
                                        codTemCli As String,
                                        codItem As String,
                                        Optional accion As String = "seleccionar")
        Try
            ' 1) Guarda COD_CLIENTE en el campo privado que usa CargaLista()
            Me.CodigoCliente = codCliente

            ' 2) TipoConsulta por Ítem
            Me.TipoConsulta = 2
            If RbItems.Checked = False Then RbItems.Checked = True

            ' 3) Rellena filtros visibles
            TxpCliente.ciTXT_COD.Text = Hp.DevuelveDato("SELECT Abr_Cliente FROM tg_cliente WHERE Cod_Cliente = '" & codCliente.Replace("'", "''") & "'", cconnect)
            TxpTemporada.ciTXT_COD.Text = codTemCli
            TxpItems.ciTXT_COD.Text = codItem

            ' 4) Carga grilla
            CargaLista()




            ' 5) Selecciona y enfoca fila del ítem (SIN lanzar excepción si no existe)
            Dim idx As Integer = BuscarIndicePorItem(codItem)
            If idx >= 0 AndAlso idx < GrdEstampadoBordado.RowCount Then
                GrdEstampadoBordado.Row = idx
                Try
                    Dim mi = GrdEstampadoBordado.GetType().GetMethod("EnsureVisible", {GetType(Integer)})
                    If mi IsNot Nothing Then mi.Invoke(GrdEstampadoBordado, New Object() {idx})
                Catch
                    ' ignorar visibilidad
                End Try
                GrdEstampadoBordado.Focus()
            Else
                ' Nada de MessageBox; si no lo encuentra, sal de forma silenciosa
                Exit Sub
            End If




            ' 6) Acción opcional
            Select Case accion.Trim().ToLowerInvariant()
                Case "seleccionar"
                ' nada extra
                Case "modificar"
                    BarraOpciones_ItemClick(BarraOpciones, New Janus.Windows.ButtonBar.ItemEventArgs(BarraOpciones.Groups(0).Items("MODIFICAR")))
                Case "proveedor"
                    BarraOpciones_ItemClick(BarraOpciones, New Janus.Windows.ButtonBar.ItemEventArgs(BarraOpciones.Groups(0).Items("PROVEEDOR")))
                Case "combinaciones"
                    BarraOpciones_ItemClick(BarraOpciones, New Janus.Windows.ButtonBar.ItemEventArgs(BarraOpciones.Groups(0).Items("COMBINACIONES")))
                Case "imprimirlista"
                    IMPRIMIR_LISTA()
                Case "precio"
                    BarraOpciones_ItemClick(BarraOpciones, New Janus.Windows.ButtonBar.ItemEventArgs(BarraOpciones.Groups(0).Items("MODPRECOM")))
            End Select

        Catch ex As Exception
            MessageBox.Show(ex.Message, "DeepLinkPorItemPorCodCliente", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Busca el índice de la fila por Cod_Item(_x)
    Private Function BuscarIndicePorItem(codItem As String) As Integer
        Try
            Dim dt = TryCast(GrdEstampadoBordado.DataSource, DataTable)
            If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return -1

            Dim colKey As String = Nothing
            If dt.Columns.Contains("Cod_Itemx") Then colKey = "Cod_Itemx"
            If colKey Is Nothing AndAlso dt.Columns.Contains("cod_itemx") Then colKey = "cod_itemx"
            If colKey Is Nothing AndAlso dt.Columns.Contains("Cod_Item") Then colKey = "Cod_Item"
            If colKey Is Nothing AndAlso dt.Columns.Contains("cod_item") Then colKey = "cod_item"
            If colKey Is Nothing Then Return -1

            For i As Integer = 0 To dt.Rows.Count - 1
                Dim val As String = If(dt.Rows(i)(colKey) Is DBNull.Value, Nothing, dt.Rows(i)(colKey).ToString().Trim())
                If String.Equals(val, codItem, StringComparison.OrdinalIgnoreCase) Then
                    Return i
                End If
            Next
            Return -1
        Catch
            Return -1
        End Try
    End Function













    Sub Inhabilita()
        grbCliente.Visible = False
        grbItems.Visible = False
        grbEstado.Visible = False
        grbEstiloCliente.Visible = False
    End Sub

    Private Sub GrdEstampadoBordado_FormattingRow(sender As Object, e As Janus.Windows.GridEX.RowLoadEventArgs) Handles GrdEstampadoBordado.FormattingRow

    End Sub
End Class