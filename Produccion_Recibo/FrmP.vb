Imports System.Windows.Forms
Imports System.IO
Imports System.Data.SqlClient
Imports System.Data
Imports System.Drawing.Text
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Windows.Forms
Imports CrystalDecisions.ReportSource
Imports CrystalDecisions.Enterprise.InfoObject
Imports System.Web.UI.WebControls
Public Class FrmP
#Region "Variables"
    Public con As New Conexion
    Dim oCompany As SAPbobsCOM.Company
    Dim connectionString As String = Conexion.ObtenerConexion.ConnectionString
    Public Shared SQL_Conexion As SqlConnection = New SqlConnection()
    Dim reader As StreamReader
    Dim entra As String
    Dim sale As String
    Public Shared ba As New List(Of String)
    Public Shared batch As New List(Of String)
    Public Shared quantity As New List(Of Double)
    Public Shared ancho As New List(Of Double)
    Private Const CP_NOCLOSE_BUTTON As Integer = &H200
#End Region
    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim myCp As CreateParams = MyBase.CreateParams
            myCp.ClassStyle = myCp.ClassStyle Or CP_NOCLOSE_BUTTON
            Return myCp
        End Get
    End Property
    Public Sub New()
        MyBase.New()
        InitializeComponent()
    End Sub

    Public Function CargaGridLecturaBC()
        'Dim query1 As String
        'query1 = ""
        'Dim query2 As String
        'query2 = ""
        'Try
        '    reader = File.OpenText(System.Windows.Forms.Application.StartupPath + "\Recibo\" + txtOrder.Text + ".txt")
        '    entra = System.Windows.Forms.Application.StartupPath + "\Recibo\" + txtOrder.Text + ".txt"
        '    Dim line As String = Nothing

        '    Dim lines As Integer = 0
        '    While (reader.Peek() <> -1)
        '        line = reader.ReadLine()
        '        Dim rnum As Integer = DGV.Rows.Add()
        '        DGV.Rows.Item(rnum).Cells(0).Value = line.ToString
        '        System.Windows.Forms.Application.DoEvents()
        '    End While
        '    reader.Close()
        'Catch ex As Exception
        '    MessageBox.Show("El Picking del numero de orden no existe")
        'End Try


    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        'DGV.Rows.Clear()
        CargaGridLecturaBC()
        Dim turnoAM_I As DateTime = CType("6:00:00 AM", DateTime)
        Dim turnoAM_F As DateTime = CType("6:00:00 PM", DateTime)

        Dim result As Integer = 0
        Dim result2 As Integer = 0
        result = DateTime.Compare(turnoAM_I, TimeOfDay.ToShortTimeString)
        result2 = DateTime.Compare(turnoAM_F, TimeOfDay.ToShortTimeString)
        If result = -1 Then
            If result2 = 1 Then
                ComboBox1.Text = "Turno 1"
            Else
                ComboBox1.Text = "Turno 2"
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim result As Integer = MessageBox.Show("Desea limpiar el objeto?", "Atencion", MessageBoxButtons.YesNoCancel)
        If result = DialogResult.Cancel Then
            MessageBox.Show("Cancelado")
        ElseIf result = DialogResult.No Then
            MessageBox.Show("Puede continuar!")
        ElseIf result = DialogResult.Yes Then
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim result As Integer = MessageBox.Show("Desea salir del modulo?", "Atencion", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            MessageBox.Show("Puede continuar")
        ElseIf result = DialogResult.Yes Then
            Try
                con.oCompany.Disconnect()
            Catch
            End Try
            Me.Hide()
        End If
    End Sub


    Private Function transferencia()
        Dim i As Integer = 0
        Dim oReturn As Integer = -1
        Dim oError As Integer = 0
        Dim errMsg As String = ""
        Dim sql As String
        Dim items As String
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim quantitys As Double

        Try
            If con.Connected = True Then

                Dim frm As New Produccion_Recibo.Batchs
                Dim oreceipt As SAPbobsCOM.Documents = con.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)

                'Contador de lineas para la orden
                sql = ("SELECT T0.PlannedQty FROM OWOR T0 WHERE T0.DocNum = '" + txtOrder.Text + "'")
                oRecordSet = con.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery(sql)
                If oRecordSet.RecordCount > 0 Then
                    quantitys = oRecordSet.Fields.Item(0).Value
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                oRecordSet = Nothing
                GC.Collect()
                oreceipt.Lines.BaseEntry = txtOrder.Text
                oreceipt.Lines.BaseType = 202
                oreceipt.Lines.Quantity = quantitys
                oreceipt.Lines.AccountCode = "_SYS00000000039"
                oreceipt.Lines.COGSCostingCode = "SL4"
                oreceipt.Lines.TransactionType = SAPbobsCOM.BoTransactionTypeEnum.botrntReject
                'true continua, false termina

                'batch.Clear()

                '    While i < (DGV.Rows.Count)
                '        batch.Add(DGV.Rows(i).Cells(0).Value.ToString)
                '        i = i + 1
                '    End While
                frm.load(txtOrder.Text, quantitys, txtOrder.Text)
                frm.ShowDialog()

                'encabezado de impresion
                sql = ("SELECT T0.PostDate,T0.ItemCode, t1.ItemName,t0.Warehouse, t0.u_comment
                        FROM OWOR T0 
                        inner join OITM T1 on t1.ItemCode = t0.ItemCode
                        WHERE T0.DocNum = '" + txtOrder.Text + "'")
                oRecordSet = con.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery(sql)
                Dim postdate As Date
                Dim itemcode As String
                Dim itemname As String
                Dim warehouse As String
                Dim comment As String
                If oRecordSet.RecordCount > 0 Then
                    postdate = oRecordSet.Fields.Item(0).Value
                    itemcode = oRecordSet.Fields.Item(1).Value
                    itemname = oRecordSet.Fields.Item(2).Value
                    warehouse = oRecordSet.Fields.Item(3).Value
                    comment = oRecordSet.Fields.Item(4).Value
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                oRecordSet = Nothing
                GC.Collect()

                Dim cont As Integer
                For cont = 0 To ba.Count - 1
                    oreceipt.Lines.BatchNumbers.BatchNumber = ba.Item(cont)
                    oreceipt.Lines.BatchNumbers.Quantity = Decimal.Round((quantity.Item(cont)), 3, MidpointRounding.AwayFromZero)
                    oreceipt.Lines.BatchNumbers.UserFields.Fields.Item("U_Ancho").Value = ancho.Item(cont)
                    oreceipt.Lines.BatchNumbers.Add()
                Next

                oreceipt.Lines.Add()
                oReturn = oreceipt.Add


                If oReturn <> 0 Then
                    MessageBox.Show(con.oCompany.GetLastErrorDescription)
                    ba.Clear()
                    quantity.Clear()
                End If
                If oReturn <> 0 Then
                    con.oCompany.GetLastError(oError, errMsg)
                    reader.Close()
                    MsgBox("Error en los Items: " & items & "Error Referente a: " & errMsg)
                    My.Computer.FileSystem.WriteAllText(System.Windows.Forms.Application.StartupPath + "\Log\LogErroresTransferencia.txt", DateTime.Now.ToString("dd/MM/yyyy-HH:mm:ss") & "Error en los Items: " & items & "Error Referente a: " & errMsg & vbCrLf, True)
                    sale = System.Windows.Forms.Application.StartupPath + "\Temp\Recibo\" + txtOrder.Text + "-ERROR-" & DateTime.Now.ToString("dd-MM-yyyy-HH-mm-ss") & ".txt"
                    File.Move(entra, sale)
                Else
                    Dim cont2 As Integer
                    For cont2 = 0 To ba.Count - 1
                        'barcode 4 , desc 2 , anch 6, pes 5, itmcod 1, het 7, coi 8,whs 9
                        imprime(ba.Item(cont2), itemname, ancho.Item(cont2), Decimal.Round((quantity.Item(cont2)), 3, MidpointRounding.AwayFromZero), itemcode, postdate, warehouse, comment, ComboBox1.Text)
                    Next
                    ba.Clear()
                    quantity.Clear()
                    MessageBox.Show("Recibo Realizado con Exito")
                    sale = System.Windows.Forms.Application.StartupPath + "\Temp\Recibo\" + txtOrder.Text + "-TransFerido" & DateTime.Now.ToString("dd-MM-yyyy-HH-mm-ss") & ".txt"
                    File.Move(entra, sale)
                    txtOrder.Clear()
                End If
            End If

        Catch ex As Exception
            MessageBox.Show("Verifique que El numero de documento")
        End Try
    End Function

    Private Sub imprime(barcode As String, desc As String, anch As String, pes As String, itmcod As String, postdate As Date, warehouse As String, comment As String, turno As String)
        Dim Report1 As New CrystalDecisions.CrystalReports.Engine.ReportDocument()
        Report1.PrintOptions.PaperOrientation = PaperOrientation.Portrait
        Report1.Load(Application.StartupPath + "\Report\InformeR.rpt", CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault.OpenReportByDefault)
        ''-----------------------------------------ENCABEZADO NO CAMBIA POR IMPRESION------------------------------------------
        Report1.SetParameterValue("docnum", txtOrder.Text)
        Report1.SetParameterValue("docdate", postdate)
        ''------------------------------------------DETALLE TRAE DATOS POR PARAMETROS------------------------------------------
        Report1.SetParameterValue("CodBatch", barcode) 'col4
        Report1.SetParameterValue("descripcion", desc) 'col2
        Report1.SetParameterValue("pesoreal", pes) 'col5
        Report1.SetParameterValue("anchotira", anch) 'col6
        Report1.SetParameterValue("bobina", itmcod) 'col1
        Report1.SetParameterValue("coil", comment)
        Report1.SetParameterValue("almacen", warehouse)
        Report1.SetParameterValue("turno", turno)
        'Report1.SetParameterValue("fechacorte", Now.ToShortDateString)
        'CrystalReportViewer1.ReportSource = Report1
        Report1.PrintToPrinter(1, False, 0, 0)
    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            Dim result As Integer = MessageBox.Show("Desea realizar el Recibo de Produccion?", "Atencion", MessageBoxButtons.YesNoCancel)
            If result = DialogResult.Cancel Then
                MessageBox.Show("Cancelado")
            ElseIf result = DialogResult.No Then
                MessageBox.Show("No se realizara el Recibo")
                Exit Sub
            ElseIf result = DialogResult.Yes Then
                If con.Connected Then
                    transferencia()
                Else
                    con.MakeConnectionSAP()
                    If con.Connected Then
                        transferencia()
                    Else
                        MessageBox.Show("Error de Conexion, intente Nuevamente")
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox("Error: " + ex.Message.ToString)
        End Try
    End Sub

    Private Sub EnterClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode.Equals(Keys.Enter) Then
            Button1_Click(1, e)
        End If
    End Sub

    Private Sub txtOrder_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtOrder.KeyDown
        Call EnterClick(sender, e)
    End Sub

    Private Sub FrmP_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtOrder.Select()
        Dim turnoAM_I As DateTime = CType("6:00:00 AM", DateTime)
        Dim turnoAM_F As DateTime = CType("6:00:00 PM", DateTime)

        Dim result As Integer = 0
        Dim result2 As Integer = 0
        result = DateTime.Compare(turnoAM_I, TimeOfDay.ToShortTimeString)
        result2 = DateTime.Compare(turnoAM_F, TimeOfDay.ToShortTimeString)
        If result = -1 Then
            If result2 = 1 Then
                ComboBox1.Text = "Turno 1"
            Else
                ComboBox1.Text = "Turno 2"
            End If
        End If
        ComboBox1.DropDownStyle = ComboBoxStyle.DropDownList
    End Sub
End Class
