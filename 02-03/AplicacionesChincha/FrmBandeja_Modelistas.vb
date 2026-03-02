Imports System.Data.SqlClient
Imports System.Threading
Imports AplicacionesCS.Clases
Imports AplicacionesCS.Helpers.clsHelpers
Imports AplicacionesCS.Helpers.clsMessages
Imports AplicacionesCS.Models
Imports Janus.Windows
Imports Janus.Windows.GridEX
Imports WinFormsControls
Imports System.Reflection




Public Class FrmBandeja_Modelistas
    Private oHP As New clsHELPER
    Private strSQL As String
    Private ds As New DataTable
    Private colEmpresa As Color
    Private sTipoBusqueda As String = "1"
    Private rs As New ADODB.Recordset

    Private Sub FrmBandeja_Modelistas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim oDtColor As DataTable = oHP.DevuelveDatos(String.Format("SELECT * FROM SEG_Empresas where cod_empresa = '{0}'", vemp), cSEGURIDAD)
        colEmpresa = Color.FromArgb(oDtColor.Rows(0)("ColorFondo_R"), oDtColor.Rows(0)("ColorFondo_G"), oDtColor.Rows(0)("ColorFondo_B"))
        Panel1.BackColor = colEmpresa
        txtCod_Estado.Text = ""
        txtCod_Estado.Enabled = False
        txtDes_estado.Text = ""
        txtDes_estado.Enabled = False
        txtCod_Cliente.Text = ""
        txtCod_Cliente.Enabled = False
        txtDes_Cliente.Text = ""
        txtDes_Cliente.Enabled = False

        txtCod_Estpro.Text = ""
        txtCod_Estpro.Enabled = False
        txtCod_Version.Text = ""
        txtCod_Version.Enabled = False
        Cargar()
    End Sub

    Private Sub Cargar()

        Try
            strSQL = ""
            strSQL = $"EXEC SP_MUESTRA_BANDEJA_MODELISTAS '{txtCod_Estpro.Text.Trim}','{txtCod_Version.Text.Trim}','{txtCod_Cliente.Text.Trim}','{txtCod_Estado.Text.Trim}','{sTipoBusqueda.Trim}'"
            ds = oHP.DevuelveDatos(strSQL, cCONNECT, True)

            AsignaDatosGridEx(GridEX1, ds, True, True)
            CheckLayoutGridEx(GridEX1)

            OcultaColumnasGridEx(GridEX1, New String() {"Cod_cliente", "Cod_temcli", "Cod_estpro", "Cod_version", "Codigo_status"})
            FormatoColumnas(GridEX1, New List(Of ModelGridEx) From {
            New ModelGridEx With {.columna = "Nom_Cliente", .ancho_columna = 180, .texto_cabecera = "CLIENTE"},
            New ModelGridEx With {.columna = "Nom_Temcli", .ancho_columna = 100, .texto_cabecera = "TEMPORADA"},
            New ModelGridEx With {.columna = "Cod_estcli", .ancho_columna = 100, .texto_cabecera = "ESTILO"},
            New ModelGridEx With {.columna = "EPV", .ancho_columna = 70, .alineacion = TextAlignment.Center, .texto_cabecera = "EPV"},
            New ModelGridEx With {.columna = "Secuencia_Version", .ancho_columna = 60, .alineacion = TextAlignment.Center, .texto_cabecera = "SEC. VERSION"},
            New ModelGridEx With {.columna = "Cod_tela", .ancho_columna = 80, .texto_cabecera = "COD TELA"},
            New ModelGridEx With {.columna = "Des_tela", .ancho_columna = 380, .texto_cabecera = "DESCRIPCION TELA"},
            New ModelGridEx With {.columna = "Descripcion_Hoja", .ancho_columna = 250, .texto_cabecera = "DESCRIPCION HOJA"},
            New ModelGridEx With {.columna = "Cod_item", .ancho_columna = 80, .alineacion = TextAlignment.Center, .texto_cabecera = "ITEM"},
            New ModelGridEx With {.columna = "Fecha_Plan_FT", .ancho_columna = 125, .alineacion = TextAlignment.Center, .texto_cabecera = "FECHA PLAN FT"},
            New ModelGridEx With {.columna = "Estado", .ancho_columna = 150, .texto_cabecera = "ESTADO"},
            New ModelGridEx With {.columna = "Fec_status", .ancho_columna = 125, .alineacion = TextAlignment.Center, .texto_cabecera = "FECHA ESTADO"},
            New ModelGridEx With {.columna = "Usuario_Aprobo", .ancho_columna = 90, .texto_cabecera = "USUARIO APROBO"},
            New ModelGridEx With {.columna = "Analista_Prenda", .ancho_columna = 200, .texto_cabecera = "ANALISTA DE PRENDA"},
            New ModelGridEx With {.columna = "Modelista", .ancho_columna = 200, .texto_cabecera = "MODELISTA"}
        })

            If Not IsNothing(ds) Then
                If ds.Columns.Count > 0 Then
                    rs = AplicacionesCS.Procedimientos.DTtoRSconvert.ConvertToRecordSet(ds)
                End If
            End If


            '-========================== PARA COLOREAR COLUMNA DE STATUS ===============================


            Dim colEstado = GridEX1.RootTable.Columns("Codigo_status")
            colEstado.EditType = Janus.Windows.GridEX.EditType.TextBox

            Dim colColorear = GridEX1.RootTable.Columns("Estado")
            colColorear.EditType = Janus.Windows.GridEX.EditType.TextBox


            Dim Sin_aprobaciones As New Janus.Windows.GridEX.GridEXFormatCondition(colEstado, Janus.Windows.GridEX.ConditionOperator.Equal, " ")
            Sin_aprobaciones.TargetColumn = colColorear
            Sin_aprobaciones.FormatStyle.ForeColor = Color.White
            Sin_aprobaciones.FormatStyle.BackColor = Color.Orange
            GridEX1.RootTable.FormatConditions.Add(Sin_aprobaciones)

            Dim Listo As New Janus.Windows.GridEX.GridEXFormatCondition(colEstado, Janus.Windows.GridEX.ConditionOperator.Equal, "L")
            Listo.TargetColumn = colColorear
            Listo.FormatStyle.ForeColor = Color.Black
            Listo.FormatStyle.BackColor = Color.GreenYellow
            GridEX1.RootTable.FormatConditions.Add(Listo)

            Dim Cerrado As New Janus.Windows.GridEX.GridEXFormatCondition(colEstado, Janus.Windows.GridEX.ConditionOperator.Equal, "C")
            Cerrado.TargetColumn = colColorear
            Cerrado.FormatStyle.ForeColor = Color.White
            Cerrado.FormatStyle.BackColor = Color.Red
            GridEX1.RootTable.FormatConditions.Add(Cerrado)


            '-========================== HIPERVINCULO  ===============================

            With GridEX1.RootTable.Columns("EPV")
                .ColumnType = Janus.Windows.GridEX.ColumnType.Link
                .CellStyle.ForeColor = Color.Blue
                .CellStyle.BackColor = Color.LightYellow
                .CellStyle.FontUnderline = Janus.Windows.GridEX.TriState.True
                .TextAlignment = Janus.Windows.GridEX.TextAlignment.Center
            End With

            '==================== HIPERVINCULO PARA ITEM ==========================
            With GridEX1.RootTable.Columns("Cod_item")
                .ColumnType = Janus.Windows.GridEX.ColumnType.Link
                .CellStyle.ForeColor = Color.Blue
                .CellStyle.BackColor = Color.LightYellow
                .CellStyle.FontUnderline = Janus.Windows.GridEX.TriState.True
                .TextAlignment = Janus.Windows.GridEX.TextAlignment.Center
            End With






        Catch ex As Exception
            _excep(ex, Me)
        End Try
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Cargar()
    End Sub







    Private Sub GridEX1_LinkClicked(sender As Object, e As Janus.Windows.GridEX.ColumnActionEventArgs) Handles GridEX1.LinkClicked
        Try
            If e.Column Is Nothing Then Exit Sub
            Dim colKey As String = e.Column.Key.ToUpperInvariant().Trim()

            If GridEX1.CurrentRow Is Nothing OrElse GridEX1.CurrentRow.RowType <> Janus.Windows.GridEX.RowType.Record Then Exit Sub

            ' Campos base que usaremos en ambos casos
            Dim EstiloPropio As String = GridEX1.GetValue("Cod_estpro").ToString().Trim()
            Dim Cod_Version As String = GridEX1.GetValue("Cod_version").ToString().Trim()
            Dim Secuencia_Version As Integer = CInt(GridEX1.GetValue("Secuencia_Version"))
            Dim Cod_Cliente As String = GridEX1.GetValue("Cod_cliente").ToString().Trim()
            Dim Cod_Temcli As String = GridEX1.GetValue("Cod_temcli").ToString().Trim()
            Dim Cod_Estcli As String = GridEX1.GetValue("Cod_estcli").ToString().Trim()
            Dim Codigo_Estampado As String = GridEX1.GetValue("Cod_item").ToString().Trim()

            If colKey = "EPV" Then
                ' --- Comportamiento actual para EPV (Hoja de Medidas) ---
                If String.IsNullOrEmpty(EstiloPropio) Then Exit Sub

                If Not String.IsNullOrEmpty(Codigo_Estampado) Then
                    ' Con aplicaciones (cuando hay ítem)
                    Formulario_Hojas_Medidas_Aplicaciones(EstiloPropio, Cod_Version, Secuencia_Version, Cod_Cliente, Cod_Temcli, Cod_Estcli)
                Else
                    ' Estándar
                    Formulario_Hojas_Medidas_EstProver(EstiloPropio, Cod_Version, Cod_Cliente, Cod_Temcli, Cod_Estcli)
                End If

            ElseIf colKey = "COD_ITEM" Then
                ' --- Nuevo: abrir el mantenimiento de Ítems/Servicios con deep-link ---
                If String.IsNullOrWhiteSpace(Codigo_Estampado) Then Exit Sub

                AbrirMantItemsServiciosPorItem(Cod_Cliente, Cod_Temcli, Codigo_Estampado, "seleccionar")
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Formulario_Hojas_Medidas_Aplicaciones(EstiloPropio As String, Cod_Version As String, Secuencia_Version As Integer, Cod_Cliente As String, Cod_Temcli As String, Cod_Estcli As String)
        Try
            Dim sPath As String = My.Application.Info.DirectoryPath
            Dim oClsForm As Object = Nothing
            Dim sDllName As String = String.Empty
            Dim tDllName As String = "Estilo_Propio"


            Dim Nombre_Cliente As String = oHP.DevuelveDato($"SELECT CONCAT(LTRIM(RTRIM(Abr_Cliente)),'-',LTRIM(RTRIM(Nom_Cliente))  ) AS Nom_Cliente FROM TG_Cliente WHERE LTRIM(RTRIM(Cod_Cliente)) = '{Cod_Cliente.Trim}'", cCONNECT)
            Dim Nombre_Temporada As String = oHP.DevuelveDato($"SELECT CONCAT( LTRIM(RTRIM(Cod_TemCli)),'-',LTRIM(RTRIM(Nom_TemCli))  ) FROM TG_TemCli WHERE LTRIM(RTRIM(Cod_Cliente)) = '{Cod_Cliente.Trim}' AND LTRIM(RTRIM(Cod_TemCli)) = '{Cod_Temcli.Trim}'", cCONNECT)


            sDllName = sPath & "\" & tDllName & ".exe"
            objFormDLL = System.Reflection.Assembly.LoadFrom(sDllName)
            If tDllName.Trim.Length = 0 Then Exit Sub
            oFormObjDLL = objFormDLL.CreateInstance(tDllName & ".clsForm", True)
            oFormObjDLL.Cod_Empresa = vemp
            oFormObjDLL.UserName = vusu
            oFormObjDLL.Cod_Perfil = vper
            oFormObjDLL.Rutas = sPath
            oFormObjDLL.ConnectEmpresa = cCONNECT
            oFormObjDLL.ConnectSeguridad = cSEGURIDAD
            oFormObjDLL.ConnectVB60 = cCONNECTVB6
            oClsForm = oFormObjDLL.getform("frmHojaMedidaApli")
            oClsForm.sAccion = "V"
            oClsForm.sCod_EstPro = EstiloPropio
            oClsForm.sCod_Version = Cod_Version
            oClsForm.sCod_Cliente = Cod_Cliente
            oClsForm.sCod_TemCli = Cod_Temcli

            oClsForm.TxtCliente.text = Nombre_Cliente
            oClsForm.TxtTemporada.text = Nombre_Temporada
            oClsForm.TxtEstiloCliente.text = Cod_Estcli
            oClsForm.TxtEstiloPropio.text = EstiloPropio
            oClsForm.TxtVersion.text = Cod_Version

            oClsForm.sDesde_Bandeja_Pendientes = True
            oClsForm.sSec_Version_Desde_Bandeja_Pendientes = Secuencia_Version
            oClsForm.ShowDialog()
            oClsForm = Nothing
        Catch ex As Exception
            _excep(ex, Me)
        End Try
    End Sub













    ' Abre frmMantItemsServicios del ensamblado de tablas y ejecuta el deep-link por ítem
    Private Sub AbrirMantItemsServiciosPorItem(codCliente As String,
                                           codTemCli As String,
                                           codItem As String,
                                           Optional accion As String = "seleccionar")
        Try
            Dim sPath As String = My.Application.Info.DirectoryPath

            '=== AJUSTA este nombre al ensamblado/namespace real donde vive frmMantItemsServicios ===
            Dim tDllName As String = "TablasEst"   ' <- por ejemplo: "TablasEst", "TablasEstilos", etc.
            '===========================================================================

            Dim sDllName As String = System.IO.Path.Combine(sPath, tDllName & ".exe")
            If Not IO.File.Exists(sDllName) Then
                ' Si es librería, cambia a .dll
                sDllName = System.IO.Path.Combine(sPath, tDllName & ".dll")
            End If

            Dim asm As Assembly = Assembly.LoadFrom(sDllName)

            ' 1) Intento con patrón clsForm (mismo que usas en Estilo_Propio)
            Dim clsFormFullName As String = tDllName & ".clsForm"
            Dim host As Object = Nothing
            Try
                host = asm.CreateInstance(clsFormFullName, True)
            Catch
                host = Nothing
            End Try

            If host IsNot Nothing Then
                ' Transferir contexto (igual que en Formulario_Hojas_Medidas_*)
                host.Cod_Empresa = vemp
                host.UserName = vusu
                host.Cod_Perfil = vper
                host.Rutas = sPath
                host.ConnectEmpresa = cCONNECT
                host.ConnectSeguridad = cSEGURIDAD
                host.ConnectVB60 = cCONNECTVB6

                ' Obtener el form solicitado
                Dim frm As Object = host.getform("frmMantItemsServicios")

                ' Invocar el deep-link público del form
                Dim mi As MethodInfo = frm.GetType().GetMethod("DeepLinkPorItemPorCodCliente")
                If mi IsNot Nothing Then
                    mi.Invoke(frm, New Object() {codCliente, codTemCli, codItem, accion})
                Else
                    MessageBox.Show("No se encontró el método DeepLinkPorItemPorCodCliente en frmMantItemsServicios.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If

                ' Mostrar modal
                Dim showDlg = frm.GetType().GetMethod("ShowDialog", Type.EmptyTypes)
                If showDlg IsNot Nothing Then
                    showDlg.Invoke(frm, Nothing)
                Else
                    ' Fallback a Show si no hay ShowDialog
                    Dim showMi = frm.GetType().GetMethod("Show", Type.EmptyTypes)
                    If showMi IsNot Nothing Then showMi.Invoke(frm, Nothing)
                End If

                Return
            End If

            ' 2) Fallback: instanciar el form directamente por nombre completo
            Dim frmTypeName As String = tDllName & ".frmMantItemsServicios"
            Dim frmType As Type = asm.GetType(frmTypeName, throwOnError:=True)

            Dim frmInst As Object = Activator.CreateInstance(frmType)

            ' Deep-link
            Dim miDL As MethodInfo = frmType.GetMethod("DeepLinkPorItemPorCodCliente")
            If miDL IsNot Nothing Then
                miDL.Invoke(frmInst, New Object() {codCliente, codTemCli, codItem, accion})
            Else
                MessageBox.Show("No se encontró el método DeepLinkPorItemPorCodCliente en frmMantItemsServicios.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

            ' Mostrar
            Dim showDialogMi = frmType.GetMethod("ShowDialog", Type.EmptyTypes)
            If showDialogMi IsNot Nothing Then
                showDialogMi.Invoke(frmInst, Nothing)
            Else
                Dim showMi = frmType.GetMethod("Show", Type.EmptyTypes)
                If showMi IsNot Nothing Then showMi.Invoke(frmInst, Nothing)
            End If

        Catch ex As Exception
            _excep(ex, Me)
        End Try
    End Sub


























    Private Sub Formulario_Hojas_Medidas_EstProver(EstiloPropio As String, Cod_Version As String, Cod_Cliente As String, Cod_Temcli As String, Cod_estcli As String)
        Try
            Dim sPath As String = My.Application.Info.DirectoryPath
            Dim oClsForm As Object = Nothing
            Dim sDllName As String = String.Empty
            Dim tDllName As String = "Estilo_Propio"


            Dim Nombre_Cliente As String = oHP.DevuelveDato($"SELECT CONCAT(LTRIM(RTRIM(Abr_Cliente)),'-',LTRIM(RTRIM(Nom_Cliente))  ) AS Nom_Cliente FROM TG_Cliente WHERE LTRIM(RTRIM(Cod_Cliente)) = '{Cod_Cliente.Trim}'", cCONNECT)
            Dim Nombre_Temporada As String = oHP.DevuelveDato($"SELECT CONCAT( LTRIM(RTRIM(Cod_TemCli)),'-',LTRIM(RTRIM(Nom_TemCli))  ) FROM TG_TemCli WHERE LTRIM(RTRIM(Cod_Cliente)) = '{Cod_Cliente.Trim}' AND LTRIM(RTRIM(Cod_TemCli)) = '{Cod_Temcli.Trim}'", cCONNECT)


            sDllName = sPath & "\" & tDllName & ".exe"
            objFormDLL = System.Reflection.Assembly.LoadFrom(sDllName)
            If tDllName.Trim.Length = 0 Then Exit Sub
            oFormObjDLL = objFormDLL.CreateInstance(tDllName & ".clsForm", True)
            oFormObjDLL.Cod_Empresa = vemp
            oFormObjDLL.UserName = vusu
            oFormObjDLL.Cod_Perfil = vper
            oFormObjDLL.Rutas = sPath
            oFormObjDLL.ConnectEmpresa = cCONNECT
            oFormObjDLL.ConnectSeguridad = cSEGURIDAD
            oFormObjDLL.ConnectVB60 = cCONNECTVB6
            oClsForm = oFormObjDLL.getform("frmHojadeMedidas_Principal")
            oClsForm.SESTILO = EstiloPropio
            oClsForm.SVERSION = Cod_Version
            oClsForm.sCod_Cliente = Cod_Cliente
            oClsForm.sCod_TemCli = Cod_Temcli

            oClsForm.TxtCliente.text = Nombre_Cliente
            oClsForm.TxtTemporada.text = Nombre_Temporada
            oClsForm.TxtEstiloCliente.text = Cod_estcli
            oClsForm.TxtEstiloPropio.text = EstiloPropio
            oClsForm.TxtVersion.text = Cod_Version

            oClsForm.sDesde_Bandeja_Pendientes = True
            oClsForm.ShowDialog()
            oClsForm = Nothing
        Catch ex As Exception
            _excep(ex, Me)
        End Try
    End Sub
    Private Sub ButtonBar1_ItemClick(sender As Object, e As ButtonBar.ItemEventArgs) Handles ButtonBar1.ItemClick
        Select Case e.Item.Key
            Case "EXPORTAR"
                Imprimir()
        End Select
    End Sub

    Private Sub Imprimir()
        Try
            Dim oo As Object = CreateObject("excel.application")
            oo.Workbooks.Open(vRuta & "\rptBandeja_Modelistas.xltm")
            oo.Visible = True
            oo.DisplayAlerts = False
            oo.Run("REPORTE", rs)
            oo.Visible = True
            oo = Nothing
        Catch ex As Exception
            MessageBox.Show(ex.Message, "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub rdbEPV_CheckedChanged(sender As Object, e As EventArgs) Handles rdbEPV.CheckedChanged
        If rdbEPV.Checked = True Then
            txtCod_Estado.Text = ""
            txtCod_Estado.Enabled = False
            txtDes_estado.Text = ""
            txtDes_estado.Enabled = False
            txtCod_Cliente.Text = ""
            txtCod_Cliente.Enabled = False
            txtDes_Cliente.Text = ""
            txtDes_Cliente.Enabled = False

            txtCod_Estpro.Text = ""
            txtCod_Estpro.Enabled = True
            txtCod_Version.Text = ""
            txtCod_Version.Enabled = True
            sTipoBusqueda = "1"
        End If
    End Sub

    Private Sub rdbEstado_CheckedChanged(sender As Object, e As EventArgs) Handles rdbEstado.CheckedChanged
        If rdbEstado.Checked = True Then
            txtCod_Cliente.Text = ""
            txtCod_Cliente.Enabled = False
            txtDes_Cliente.Text = ""
            txtDes_Cliente.Enabled = False
            txtCod_Estpro.Text = ""
            txtCod_Estpro.Enabled = False
            txtCod_Version.Text = ""
            txtCod_Version.Enabled = False

            txtCod_Estado.Text = ""
            txtCod_Estado.Enabled = True
            txtDes_estado.Text = ""
            txtDes_estado.Enabled = True
            sTipoBusqueda = "2"
        End If
    End Sub

    Private Sub rdbCliente_CheckedChanged(sender As Object, e As EventArgs) Handles rdbCliente.CheckedChanged
        If rdbCliente.Checked = True Then
            txtCod_Estpro.Text = ""
            txtCod_Estpro.Enabled = False
            txtCod_Version.Text = ""
            txtCod_Version.Enabled = False
            txtCod_Estado.Text = ""
            txtCod_Estado.Enabled = False
            txtDes_estado.Text = ""
            txtDes_estado.Enabled = False

            txtCod_Cliente.Text = ""
            txtCod_Cliente.Enabled = True
            txtDes_Cliente.Text = ""
            txtDes_Cliente.Enabled = True
            sTipoBusqueda = "3"
        End If
    End Sub

    Private Sub txtCod_Cliente_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCod_Cliente.KeyPress
        If e.KeyChar = Chr(Keys.Enter) Then
            Try
                strSQL = ""
                strSQL = $"SELECT LTRIM(RTRIM(Cod_Cliente)) AS Cod_Cliente, LTRIM(RTRIM(Nom_Cliente)) AS Nom_Cliente FROM TG_Cliente"
                ds = oHP.DevuelveDatos(strSQL, cCONNECT)
                clsRetornaMuestraDatos._retorna_muestra_datos_eb(ds, txtCod_Cliente, txtDes_Cliente)
                btnBuscar.Select()
            Catch ex As Exception
                MessageBox.Show("Error al cargar los datos: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub txtCod_Estado_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCod_Estado.KeyPress
        If e.KeyChar = Chr(Keys.Enter) Then
            Try
                strSQL = ""
                strSQL = $"Select Codigo, Descripcion_status As Descripcion FROM ESTADO_BANDEJA_MODELISTAS"
                ds = oHP.DevuelveDatos(strSQL, cCONNECT)
                clsRetornaMuestraDatos._retorna_muestra_datos_eb(ds, txtCod_Estado, txtDes_estado)
                btnBuscar.Select()
            Catch ex As Exception
                MessageBox.Show("Error al cargar los datos: " & ex.Message)
            End Try
        End If
    End Sub
End Class