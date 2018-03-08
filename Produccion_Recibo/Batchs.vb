Imports System.Data.SqlClient
Public Class Batchs
    Public con As New Conexion
    Public batchsnum As Double
    Dim cantidadR As Double
    Dim objectCode As Integer
    Public Shared SQL_Conexion As SqlConnection = New SqlConnection()
    Dim connectionString As String = Conexion.ObtenerConexion.ConnectionString
    Private Const CP_NOCLOSE_BUTTON As Integer = &H200

    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim myCp As CreateParams = MyBase.CreateParams
            myCp.ClassStyle = myCp.ClassStyle Or CP_NOCLOSE_BUTTON
            Return myCp
        End Get
    End Property
    Friend Sub load(itemcode As String, cantidad As Double, objectcodes As Integer)
        Label4.Text = itemcode
        Label2.Text = cantidad
        cantidadR = cantidad
        objectCode = objectcodes
        CargaItems(itemcode)
    End Sub

    Public Function CargaItems(itemcode As String)
        Dim valor As Double
        valor = cantidadR / FrmP.batch.Count
        For Each item In FrmP.batch
            DGV.Rows.Add(item, valor)
        Next
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If Decimal.Round(Label6.Text, 3, MidpointRounding.AwayFromZero) = 0 Then
            Dim i As Integer = 0
            While i < (DGV.Rows.Count - 1)
                FrmP.ba.Add(DGV.Rows(i).Cells.Item(0).Value.ToString)
                FrmP.quantity.Add(DGV.Rows(i).Cells.Item(1).Value.ToString)
                FrmP.ancho.Add(DGV.Rows(i).Cells.Item(2).Value.ToString)
                i = i + 1
            End While
            Me.Hide()
        Else
            MessageBox.Show("Verifique que la cantidad requerida concuerde con el total consumido")
        End If
    End Sub

    Private Sub BatchsFase2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        Dim suma As Double
        Dim queda As Double
        For Each row As DataGridViewRow In DGV.Rows
            suma += Val(row.Cells(1).Value)
        Next

        queda = Decimal.Round((Convert.ToDouble(Label2.Text) - suma), 3, MidpointRounding.AwayFromZero)
        If queda = 0 Then
            Button2.Visible = True
            Button1.Visible = True
            Label6.Text = queda
            Label6.Refresh()
        Else
            'queda = Convert.ToDouble(Label2.Text) - suma
            Label6.Text = queda
            Label6.Refresh()
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
            MessageBox.Show("Cancele el Objeto e inicie nuevamente")
            Me.Hide()
        End If
    End Sub
End Class