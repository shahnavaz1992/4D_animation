Imports System.Data.OleDb
Imports System.Windows.Forms
'Imports System.Math
Imports Autoanimation.Class1
Imports System.Collections.Generic
Imports System.Text
Imports Autodesk.Navisworks.Api.Plugins
Imports Autodesk.Navisworks.Api
Imports Autodesk.Navisworks.Api.Application


Public Class SelectForm

    Public ReDataTable As New DataTable

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim OpenFileDialog1 As New OpenFileDialog
        OpenFileDialog1.Title = "Open File..."
        OpenFileDialog1.Multiselect = False
        OpenFileDialog1.Filter = "All Files|*.*"
        OpenFileDialog1.ShowDialog()
        TextBox1.Text = OpenFileDialog1.FileName

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If RadioButton1.Checked = True Then
            Button4_Click(sender, e)
        End If
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim re As New Class1

        Dim pathDB As String = Me.TextBox1.Text

        '*******************************************
        ' Conect to Access

        Dim connetionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & pathDB

        Dim myAccessConn As OleDbConnection
        myAccessConn = New OleDbConnection(connetionString)
        myAccessConn.Open()
        Dim dscmd As New OleDbDataAdapter("Select * from path", myAccessConn)
        Dim myDataSet As New DataSet
        Dim da As New OleDbCommandBuilder(dscmd)
        dscmd.Fill(myDataSet, "path")
        ReDataTable = myDataSet.Tables("path")

        For i As Integer = 0 To ReDataTable.Rows.Count - 1


            If ReDataTable.Rows(i).Item("PathID").ToString = "Vertical move" Then
                re.ReHoistUp(i, ReDataTable)
            End If

            If ReDataTable.Rows(i).Item("PathID").ToString = "Object rotation" Then
                re.RerotateObjAlongAxis(i, ReDataTable)
            End If

            If ReDataTable.Rows(i).Item("PathID").ToString = "Boom rotation" Then
                re.RerotateCraneAlongZ(i, ReDataTable)
            End If

            If ReDataTable.Rows(i).Item("PathID").ToString = "Walk" Then
                re.Rewalk(i, ReDataTable)
            End If

            If ReDataTable.Rows(i).Item("PathID").ToString = "Boom Up" Then
                re.ReboomUp(i, ReDataTable)
            End If

        Next

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles TimeCalculator.Click

        Dim Scale As Double = CDbl(TimeScale.Text)
        Dim pathDB As String = Me.TextBox1.Text
        Dim connetionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & pathDB
        Dim myAccessConn As OleDbConnection

        Dim path As DataTable
        Dim CraneFigures As DataTable

        myAccessConn = New OleDbConnection(connetionString)
        myAccessConn.Open()
        Dim OleDbDataAdapter1 As New OleDbDataAdapter("Select * from path", myAccessConn)
        Dim myDataSet As New DataSet
        Dim da As New OleDbCommandBuilder(OleDbDataAdapter1)
        OleDbDataAdapter1.Fill(myDataSet, "path")
        path = myDataSet.Tables("path")

        Dim OleDbDataAdapter2 As New OleDbDataAdapter("Select * from CraneFigures", myAccessConn)
        Dim db As New OleDbCommandBuilder(OleDbDataAdapter2)
        OleDbDataAdapter2.Fill(myDataSet, "CraneFigures")
        CraneFigures = myDataSet.Tables("CraneFigures")



        '''''''' calculate start and finish time

        Dim st As Double = 0
        Dim sf As Double = 0
        Dim Envelope As String
        Dim ObjectNum As String
        Dim CraneID As String
        Dim delta As Double

        path.Rows(0).Item("StartTime") = 0
        For i As Integer = 0 To path.Rows.Count - 1

            Envelope = path.Rows(i).Item("ObjectName")
            ObjectNum = Strings.Right(Envelope, 12)
            CraneID = "Crane" & ObjectNum

            If path.Rows(i).Item("Rotation") <> 0 Then
                delta = Math.Abs(path.Rows(i).Item("Rotation"))
            ElseIf path.Rows(i).Item("X") <> 0 Then
                delta = Math.Abs(path.Rows(i).Item("X"))
            ElseIf path.Rows(i).Item("Y") <> 0 Then
                delta = Math.Abs(path.Rows(i).Item("Y"))
            Else
                delta = Math.Abs(path.Rows(i).Item("Z"))
            End If

            For j As Integer = 0 To CraneFigures.Rows.Count - 1
                If CraneFigures.Rows(j).Item("CraneName") = CraneID And CraneFigures.Rows(j).Item("PathID") = path.Rows(i).Item("PathID") Then

                    path.Rows(i).Item("FinishTime") = CInt(path.Rows(i).Item("StartTime") + (delta / Scale / CraneFigures.Rows(j).Item("Velocity")))
                    If i <> path.Rows.Count - 1 Then
                        If path.Rows(i + 1).Item("ObjectName") = path.Rows(i).Item("ObjectName") Then
                            path.Rows(i + 1).Item("StartTime") = CInt(path.Rows(i).Item("FinishTime") + 1)
                        Else
                            path.Rows(i + 1).Item("StartTime") = 0
                        End If
                    End If
                End If
            Next

        Next

        'Dim deleteCommand2 As New OleDbCommand("Delete * FROM path;", myAccessConn)
        'Dim dj As OleDbCommandBuilder
        'deleteCommand2.ExecuteNonQuery()
        'deleteCommand2.Dispose()

        OleDbDataAdapter1.Update(myDataSet, "path")

        MessageBox.Show("Time calculates")
        myAccessConn.Close()
    End Sub
End Class