Imports System.Data.OleDb
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.Text
Imports Autodesk.Navisworks.Api.Plugins
Imports Autodesk.Navisworks.Api
Imports Autodesk.Navisworks.Api.Application

Imports Autodesk.Navisworks.Api.Timeliner
Imports Autodesk.Navisworks.Animator



<PluginAttribute("animator", "F.Shahnavaz", DisplayName:="4Danimation")>
Public Class Class1
    Inherits AddInPlugin

    Private oDoc As Document
    Public myAccessConn As OleDbConnection

    Dim MySelectForm As New SelectForm


    Private connetionString As String

    Public myDataTable As DataTable



    Public Overrides Function Execute(ParamArray parameters() As String) As Integer


        MySelectForm.ShowDialog()
        Dim pathDB As String = MySelectForm.TextBox1.Text

        '*******************************************
        ' Conect to Access

        connetionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & pathDB

        Try
            myAccessConn = New OleDbConnection(connetionString)
            myAccessConn.Open()

            Dim dscmd As New OleDbDataAdapter("Select * from path", myAccessConn)
            Dim myDataSet As New DataSet
            Dim da As New OleDbCommandBuilder(dscmd)
            dscmd.Fill(myDataSet, "path")
            myDataTable = myDataSet.Tables("path")





            '''''' start animation
            m_Index = 0

            AddHandler Autodesk.Navisworks.Api.Application.Idle, AddressOf Idle_EventHandler



            myAccessConn.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        Return 0


    End Function

    Dim m_Index As Integer = 0
    Dim oMoveStep As Double = 2
    Dim oRotateStep As Double = 0.1




    Private Sub Idle_EventHandler(ByVal sender As Object, ByVal e As EventArgs)

        oDoc = ActiveDocument

        Dim maxfinish As Integer = 0
        For r = 0 To myDataTable.Rows.Count - 1
            If myDataTable.Rows(r).Item("FinishTime") > maxfinish Then
                maxfinish = myDataTable.Rows(r).Item("FinishTime")
            End If
        Next

        'get handle of Navisworks window

        Dim hWnd As IntPtr = Gui.MainWindow.Handle

            Dim s As Double
            Dim f As Double

            ' If m_Index < 200 Then
            For i As Integer = 0 To myDataTable.Rows.Count - 1

                s = myDataTable.Rows(i).Item("StartTime")
                f = myDataTable.Rows(i).Item("FinishTime")
                If myDataTable.Rows(i).Item("PathID").ToString = "Vertical move" And s <= m_Index And m_Index < f Then
                    HoistUp(i, s, f)
                End If

                If myDataTable.Rows(i).Item("PathID").ToString = "Object rotation" And s <= m_Index And m_Index < f Then
                    rotateObjAlongAxis(i, s, f)
                End If

                If myDataTable.Rows(i).Item("PathID").ToString = "Boom rotation" And s <= m_Index And m_Index < f Then
                    rotateCraneAlongZ(i, s, f)
                End If

            If myDataTable.Rows(i).Item("PathID").ToString = "Walk" And s <= m_Index And m_Index < f Then
                walk(i, s, f)
            End If

            If myDataTable.Rows(i).Item("PathID").ToString = "Boom Up" And s <= m_Index And m_Index < f Then
                boomUp(i, s, f)
            End If


        Next
            m_Index = m_Index + 1

        If m_Index > maxfinish Then
            RemoveHandler Autodesk.Navisworks.Api.Application.Idle, AddressOf Idle_EventHandler
            MySelectForm.ShowDialog()
        End If

    End Sub



#Region "transform Object"

    Public Sub HoistUp(ByVal row As Integer, start As Double, finish As Double)

        '  oDoc = ActiveDocument
        Dim newvector3d As Vector3D
        Dim NewOverrideTrans As Transform3D
        'Dim Rotation As Rotation3D
        Dim NewIndentityM As Matrix3
        Dim duration As Double
        duration = finish - start

        Dim h As Double = (CDbl(myDataTable.Rows(row).Item("Z"))) / duration

        newvector3d = New Vector3D(0, 0, h)

        Dim SelectionModelItems1 As New ModelItemCollection
        Dim SelectionModelItems2 As New ModelItemCollection


        NewIndentityM = New Matrix3



        Dim MySearch1 As Search
        Dim MySearch2 As Search
        Dim MySearchConditionGroup1 As New List(Of SearchCondition)
        Dim MySearchConditionGroup2 As New List(Of SearchCondition)

        Dim SearchCondition1 As SearchCondition
        SearchCondition1 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Name")

        Dim SearchCondition2 As SearchCondition
        SearchCondition2 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Name")

        Dim Envelope As String = myDataTable.Rows(row).Item("ObjectName")
        Dim ObjectNum As String = Right(Envelope, 12)
        Dim Rigging As String = "Rigging" & ObjectNum
        SearchCondition1 = SearchCondition1.DisplayStringContains(Envelope) '  "Rigging2710-PR-525P"
        SearchCondition2 = SearchCondition2.DisplayStringContains(Rigging)
        MySearchConditionGroup1.Add(SearchCondition1)
        MySearchConditionGroup2.Add(SearchCondition2)


        MySearch1 = New Search
        MySearch1.Selection.SelectAll()
        MySearch1.SearchConditions.AddGroup(MySearchConditionGroup1)
        SelectionModelItems1 = MySearch1.FindAll(oDoc, False)

        MySearch2 = New Search
        MySearch2.Selection.SelectAll()
        MySearch2.SearchConditions.AddGroup(MySearchConditionGroup2)
        SelectionModelItems2 = MySearch2.FindAll(oDoc, False)


        NewOverrideTrans = New Transform3D(NewIndentityM, newvector3d)
        ActiveDocument.Models.OverridePermanentTransform(SelectionModelItems1, NewOverrideTrans, True)
        ActiveDocument.Models.OverridePermanentTransform(SelectionModelItems2, NewOverrideTrans, True)
    End Sub

    Public Sub rotateObjAlongAxis(ByVal row As Integer, start As Double, finish As Double)

        oDoc = ActiveDocument


        Dim SelectionModelItems1 As New ModelItemCollection
        Dim SelectionModelItems2 As New ModelItemCollection

        Dim duration As Double
        duration = finish - start

        Dim MySearch1 As Search
        Dim MySearch2 As Search
        Dim MySearchConditionGroup1 As New List(Of SearchCondition)
        Dim MySearchConditionGroup2 As New List(Of SearchCondition)

        Dim SearchCondition1 As SearchCondition
        SearchCondition1 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Name")

        Dim SearchCondition2 As SearchCondition
        SearchCondition2 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Name")
        ''''
        Dim Envelope As String = myDataTable.Rows(row).Item("ObjectName")
        Dim ObjectNum As String = Right(Envelope, 12)
        Dim Rigging As String = "Rigging" & ObjectNum
        SearchCondition1 = SearchCondition1.DisplayStringContains(Envelope) '  "Rigging2710-PR-525P"
        SearchCondition2 = SearchCondition2.DisplayStringContains(Rigging)
        MySearchConditionGroup1.Add(SearchCondition1)
        MySearchConditionGroup2.Add(SearchCondition2)


        MySearch1 = New Search
        MySearch1.Selection.SelectAll()
        MySearch1.SearchConditions.AddGroup(MySearchConditionGroup1)
        SelectionModelItems1 = MySearch1.FindAll(oDoc, False)

        MySearch2 = New Search
        MySearch2.Selection.SelectAll()
        MySearch2.SearchConditions.AddGroup(MySearchConditionGroup2)
        SelectionModelItems2 = MySearch2.FindAll(oDoc, False)

        Dim R As Double = CDbl(myDataTable.Rows(row).Item("Rotation"))

        Dim oCenterP As Point3D = SelectionModelItems1.BoundingBox.Center
        Dim oMoveBackToOrig As Vector3D = New Vector3D(-oCenterP.X, -oCenterP.Y, -oCenterP.Z)

        Dim oMoveBackToOrigM As Transform3D = Transform3D.CreateTranslation(oMoveBackToOrig)

        '  set the axis we will rotate around （0,0,1）

        Dim odeltaA As UnitVector3D = New UnitVector3D(0, 0, 1)

        ' Create delta of Quaternion: axis is Z,       

        Dim delta As Rotation3D = New Rotation3D(odeltaA, (R * Math.PI / 180) / duration)

        'create a transform from a matrix with a rotation.

        Dim oNewOverrideTrans As Transform3D = New Transform3D(delta)


        Dim oFinalM As Transform3D = Transform3D.Multiply(oMoveBackToOrigM, oNewOverrideTrans)

        oFinalM = Transform3D.Multiply(oFinalM, oMoveBackToOrigM.Inverse())

        oDoc.Models.OverridePermanentTransform(SelectionModelItems1, oFinalM, True)
        oDoc.Models.OverridePermanentTransform(SelectionModelItems2, oFinalM, True)
    End Sub

    Public Sub rotateCraneAlongZ(ByVal row As Integer, start As Double, finish As Double)

        oDoc = ActiveDocument

        Dim duration As Double
        duration = finish - start


        Dim SelectionModelItems1 As New ModelItemCollection
        Dim SelectionModelItems2 As New ModelItemCollection
        Dim SelectionModelItems3 As New ModelItemCollection
        Dim SelectionModelItems4 As New ModelItemCollection
        Dim SelectionModelItems5 As New ModelItemCollection
        Dim SelectionModelItems6 As New ModelItemCollection
        Dim SelectionModelItems7 As New ModelItemCollection
        Dim SelectionModelItems8 As New ModelItemCollection
        Dim SelectionModelItems9 As New ModelItemCollection
        Dim SelectionModelItems10 As New ModelItemCollection

        Dim MySearch1 As Search
        Dim MySearch2 As Search
        Dim MySearch3 As Search


        Dim MySearchConditionGroup1 As New List(Of SearchCondition)
        Dim MySearchConditionGroup2 As New List(Of SearchCondition)
        Dim MySearchConditionGroup3 As New List(Of SearchCondition)


        Dim SearchCondition1 As SearchCondition
        SearchCondition1 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Name")

        Dim SearchCondition2 As SearchCondition
        SearchCondition2 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Name")

        Dim SearchCondition3 As SearchCondition
        SearchCondition3 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Name")



        Dim SearchCondition8 As SearchCondition
        SearchCondition8 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Type")




        Dim Envelope As String = myDataTable.Rows(row).Item("ObjectName")
        Dim ObjectNum As String = Right(Envelope, 12)
        Dim Rigging As String = "Rigging" & ObjectNum
        Dim Crane As String = "Crane" & ObjectNum

        SearchCondition1 = SearchCondition1.DisplayStringContains(Envelope)
        SearchCondition2 = SearchCondition2.DisplayStringContains(Rigging)
        SearchCondition3 = SearchCondition3.DisplayStringContains(Crane)


        MySearchConditionGroup1.Add(SearchCondition1)
        MySearchConditionGroup2.Add(SearchCondition2)
        MySearchConditionGroup3.Add(SearchCondition3)



        MySearch1 = New Search
        MySearch1.Selection.SelectAll()
        MySearch1.SearchConditions.AddGroup(MySearchConditionGroup1)
        SelectionModelItems1 = MySearch1.FindAll(oDoc, False)

        MySearch2 = New Search
        MySearch2.Selection.SelectAll()
        MySearch2.SearchConditions.AddGroup(MySearchConditionGroup2)
        SelectionModelItems2 = MySearch2.FindAll(oDoc, False)

        MySearch3 = New Search
        MySearch3.Selection.SelectAll()
        MySearch3.SearchConditions.AddGroup(MySearchConditionGroup3)
        SelectionModelItems3 = MySearch3.FindAll(oDoc, False)


        For Each ModelItem In SelectionModelItems3
            For Each insert In ModelItem.Children
                For Each block In insert.Children
                    If block.DisplayName = "Boom_CC 2800" Then
                        SelectionModelItems4.Add(block)
                    End If
                Next
            Next
        Next

        For Each ModelItem In SelectionModelItems3
            For Each insert In ModelItem.Children
                For Each block In insert.Children
                    If block.DisplayName = "Mast_CC 2800" Then
                        SelectionModelItems5.Add(block)
                    End If
                Next
            Next
        Next

        For Each ModelItem In SelectionModelItems3
            For Each insert In ModelItem.Children
                For Each block In insert.Children
                    If block.DisplayName = "Superlift_CC 2800" Then
                        SelectionModelItems6.Add(block)
                    End If
                Next
            Next
        Next

        For Each ModelItem In SelectionModelItems3
            For Each insert In ModelItem.Children
                For Each block In insert.Children
                    If block.DisplayName = "LoadLines_CC 2800" Then
                        SelectionModelItems7.Add(block)
                    End If
                Next
            Next
        Next

        For Each ModelItem In SelectionModelItems3
            For Each block In ModelItem.Children
                If block.ClassDisplayName = "Line" Then
                    SelectionModelItems8.Add(block)
                End If
            Next
        Next



        For Each ModelItem In SelectionModelItems3
            For Each insert In ModelItem.Children
                For Each block In insert.Children
                    If block.DisplayName = "Body_CC 2800" Then
                        SelectionModelItems9.Add(block)
                    End If
                Next
            Next
        Next

        For Each ModelItem In SelectionModelItems3
            For Each insert In ModelItem.Children
                For Each block In insert.Children
                    If block.DisplayName = "Crawler_CC 2800" Then
                        SelectionModelItems10.Add(block)
                    End If
                Next
            Next
        Next



        Dim oCenterP As Point3D = SelectionModelItems10.BoundingBox.Center()
        Dim oMinP As Point3D = SelectionModelItems10.BoundingBox.Min()
        Dim oMaxP As Point3D = SelectionModelItems10.BoundingBox.Max()
        Dim oMoveBackToOrig As Vector3D = New Vector3D(-oCenterP.X, -oCenterP.Y, -oCenterP.Z)

        Dim oMoveBackToOrigM As Transform3D = Transform3D.CreateTranslation(oMoveBackToOrig)

        '  set the axis we will rotate around （0,0,1）

        Dim odeltaA As UnitVector3D = New UnitVector3D(0, 0, 1)

        ' Create delta of Quaternion: axis is Z,       
        Dim RB As Double = myDataTable.Rows(row).Item("Rotation")
        Dim delta As Rotation3D = New Rotation3D(odeltaA, RB * (Math.PI / 180) / duration) ' -0.785

        'create a transform from a matrix with a rotation.

        Dim oNewOverrideTrans As Transform3D = New Transform3D(delta)


        Dim oFinalM As Transform3D = Transform3D.Multiply(oMoveBackToOrigM, oNewOverrideTrans)

        oFinalM = Transform3D.Multiply(oFinalM, oMoveBackToOrigM.Inverse())



        oDoc.Models.OverridePermanentTransform(SelectionModelItems1, oFinalM, True)
        oDoc.Models.OverridePermanentTransform(SelectionModelItems2, oFinalM, True)
        'oDoc.Models.OverridePermanentTransform(SelectionModelItems3, oFinalM, True)
        oDoc.Models.OverridePermanentTransform(SelectionModelItems4, oFinalM, True)
        oDoc.Models.OverridePermanentTransform(SelectionModelItems5, oFinalM, True)
        oDoc.Models.OverridePermanentTransform(SelectionModelItems6, oFinalM, True)
        oDoc.Models.OverridePermanentTransform(SelectionModelItems7, oFinalM, True)
        oDoc.Models.OverridePermanentTransform(SelectionModelItems8, oFinalM, True)
        oDoc.Models.OverridePermanentTransform(SelectionModelItems9, oFinalM, True)
    End Sub

    Public Sub boomUp(ByVal row As Integer, start As Double, finish As Double)

        oDoc = ActiveDocument

        Dim duration As Double
        duration = finish - start


        Dim SelectionModelItems1 As New ModelItemCollection
        Dim SelectionModelItems2 As New ModelItemCollection
        Dim SelectionModelItems3 As New ModelItemCollection
        Dim SelectionModelItems4 As New ModelItemCollection
        Dim SelectionModelItems5 As New ModelItemCollection
        Dim SelectionModelItems6 As New ModelItemCollection
        Dim SelectionModelItems7 As New ModelItemCollection
        Dim SelectionModelItems8 As New ModelItemCollection
        Dim SelectionModelItems9 As New ModelItemCollection
        Dim SelectionModelItems10 As New ModelItemCollection

        Dim MySearch1 As Search
        Dim MySearch2 As Search
        Dim MySearch3 As Search


        Dim MySearchConditionGroup1 As New List(Of SearchCondition)
        Dim MySearchConditionGroup2 As New List(Of SearchCondition)
        Dim MySearchConditionGroup3 As New List(Of SearchCondition)


        Dim SearchCondition1 As SearchCondition
        SearchCondition1 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Name")

        Dim SearchCondition2 As SearchCondition
        SearchCondition2 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Name")

        Dim SearchCondition3 As SearchCondition
        SearchCondition3 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Name")



        Dim SearchCondition8 As SearchCondition
        SearchCondition8 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Type")




        Dim Envelope As String = myDataTable.Rows(row).Item("ObjectName")
        Dim ObjectNum As String = Right(Envelope, 12)
        Dim Rigging As String = "Rigging" & ObjectNum
        Dim Crane As String = "Crane" & ObjectNum

        SearchCondition1 = SearchCondition1.DisplayStringContains(Envelope)
        SearchCondition2 = SearchCondition2.DisplayStringContains(Rigging)
        SearchCondition3 = SearchCondition3.DisplayStringContains(Crane)


        MySearchConditionGroup1.Add(SearchCondition1)
        MySearchConditionGroup2.Add(SearchCondition2)
        MySearchConditionGroup3.Add(SearchCondition3)



        MySearch1 = New Search
        MySearch1.Selection.SelectAll()
        MySearch1.SearchConditions.AddGroup(MySearchConditionGroup1)
        SelectionModelItems1 = MySearch1.FindAll(oDoc, False)

        MySearch2 = New Search
        MySearch2.Selection.SelectAll()
        MySearch2.SearchConditions.AddGroup(MySearchConditionGroup2)
        SelectionModelItems2 = MySearch2.FindAll(oDoc, False)

        MySearch3 = New Search
        MySearch3.Selection.SelectAll()
        MySearch3.SearchConditions.AddGroup(MySearchConditionGroup3)
        SelectionModelItems3 = MySearch3.FindAll(oDoc, False)


        For Each ModelItem In SelectionModelItems3
            For Each insert In ModelItem.Children
                For Each block In insert.Children
                    If block.DisplayName = "Boom_CC 2800" Then
                        SelectionModelItems4.Add(block)
                    End If
                Next
            Next
        Next



        For Each ModelItem In SelectionModelItems3
            For Each insert In ModelItem.Children
                For Each block In insert.Children
                    If block.DisplayName = "LoadLines_CC 2800" Then
                        SelectionModelItems7.Add(block)
                    End If
                Next
            Next
        Next

        For Each ModelItem In SelectionModelItems3
            For Each block In ModelItem.Children
                If block.ClassDisplayName = "Line" Then
                    SelectionModelItems8.Add(block)
                End If
            Next
        Next



        For Each ModelItem In SelectionModelItems3
            For Each insert In ModelItem.Children
                For Each block In insert.Children
                    If block.DisplayName = "Crawler_CC 2800" Then
                        SelectionModelItems10.Add(block)
                    End If
                Next
            Next
        Next



        Dim oCenterP As Point3D = SelectionModelItems4.BoundingBox.Center()
        Dim oMinP As Point3D = SelectionModelItems4.BoundingBox.Min()
        Dim oMaxP As Point3D = SelectionModelItems4.BoundingBox.Max()
        Dim oMoveBackToOrig As Vector3D = New Vector3D(-oMinP.X, -oMinP.Y, -oMinP.Z)

        Dim oMoveBackToOrigM As Transform3D = Transform3D.CreateTranslation(oMoveBackToOrig)

        '  set the axis we will rotate around （0,0,1）

        Dim odeltaA As UnitVector3D = New UnitVector3D(0, 1, 0)

        ' Create delta of Quaternion: axis is Z,       
        Dim RB As Double = myDataTable.Rows(row).Item("Rotation")
        Dim delta As Rotation3D = New Rotation3D(odeltaA, RB * (Math.PI / 180) / duration) ' -0.785

        'create a transform from a matrix with a rotation.

        Dim oNewOverrideTrans As Transform3D = New Transform3D(delta)


        Dim oFinalM As Transform3D = Transform3D.Multiply(oMoveBackToOrigM, oNewOverrideTrans)

        oFinalM = Transform3D.Multiply(oFinalM, oMoveBackToOrigM.Inverse())


        ''''''''
        Dim oCenterP2 As Point3D = SelectionModelItems7.BoundingBox.Center()
        Dim oMinP2 As Point3D = SelectionModelItems7.BoundingBox.Min()
        Dim oMaxP2 As Point3D = SelectionModelItems7.BoundingBox.Max()
        Dim oMoveBackToOrig2 As Vector3D = New Vector3D(-oMaxP2.X, -oMaxP2.Y, -oMaxP2.Z)

        Dim oMoveBackToOrigM2 As Transform3D = Transform3D.CreateTranslation(oMoveBackToOrig2)

        '  set the axis we will rotate around （0,0,1）

        Dim odeltaA2 As UnitVector3D = New UnitVector3D(0, 1, 0)

        ' Create delta of Quaternion: axis is Z,       
        Dim RB2 As Double = -myDataTable.Rows(row).Item("Rotation")
        Dim delta2 As Rotation3D = New Rotation3D(odeltaA2, RB2 * (Math.PI / 180) / duration) ' -0.785

        'create a transform from a matrix with a rotation.

        Dim oNewOverrideTrans2 As Transform3D = New Transform3D(delta2)


        Dim oFinalM2 As Transform3D

        oFinalM2 = oFinalM
        oFinalM2 = Transform3D.Multiply(oFinalM2, oMoveBackToOrigM2)
        oFinalM2 = Transform3D.Multiply(oFinalM2, oNewOverrideTrans2)
        oFinalM2 = Transform3D.Multiply(oFinalM2, oMoveBackToOrigM2.Inverse())
        'oFinalM2 = Transform3D.Multiply(oFinalM, oFinalM2)


        oDoc.Models.OverridePermanentTransform(SelectionModelItems1, oFinalM2, True)
        oDoc.Models.OverridePermanentTransform(SelectionModelItems2, oFinalM2, True)
        'oDoc.Models.OverridePermanentTransform(SelectionModelItems3, oFinalM, True)
        oDoc.Models.OverridePermanentTransform(SelectionModelItems4, oFinalM, True)
        'oDoc.Models.OverridePermanentTransform(SelectionModelItems5, oFinalM, True)
        'oDoc.Models.OverridePermanentTransform(SelectionModelItems6, oFinalM, True)
        oDoc.Models.OverridePermanentTransform(SelectionModelItems7, oFinalM2, True)
        oDoc.Models.OverridePermanentTransform(SelectionModelItems8, oFinalM, True)
        'oDoc.Models.OverridePermanentTransform(SelectionModelItems9, oFinalM, True)
    End Sub

    Public Sub walk(ByVal row As Integer, start As Double, finish As Double)

        oDoc = ActiveDocument
        Dim newvector3d As Vector3D
        Dim NewOverrideTrans As Transform3D
        'Dim Rotation As Rotation3D
        Dim NewIndentityM As Matrix3

        Dim duration As Double
        duration = finish - start


        Dim wx As Double = myDataTable.Rows(row).Item("X") / duration
        Dim wy As Double = myDataTable.Rows(row).Item("Y") / duration
        newvector3d = New Vector3D(wx, wy, 0)

        Dim SelectionModelItems1 As New ModelItemCollection
        Dim SelectionModelItems2 As New ModelItemCollection
        Dim SelectionModelItems3 As New ModelItemCollection


        NewIndentityM = New Matrix3



        Dim MySearch1 As Search
        Dim MySearch2 As Search
        Dim MySearch3 As Search
        Dim MySearchConditionGroup1 As New List(Of SearchCondition)
        Dim MySearchConditionGroup2 As New List(Of SearchCondition)
        Dim MySearchConditionGroup3 As New List(Of SearchCondition)

        Dim SearchCondition1 As SearchCondition
        SearchCondition1 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Name")

        Dim SearchCondition2 As SearchCondition
        SearchCondition2 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Name")

        Dim SearchCondition3 As SearchCondition
        SearchCondition3 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Name")

        Dim Envelope As String = myDataTable.Rows(row).Item("ObjectName")
        Dim ObjectNum As String = Right(Envelope, 12)
        Dim Rigging As String = "Rigging" & ObjectNum
        Dim Crane As String = "Crane" & ObjectNum

        SearchCondition1 = SearchCondition1.DisplayStringContains(Envelope) '  "Rigging2710-PR-525P"
        SearchCondition2 = SearchCondition2.DisplayStringContains(Rigging)
        SearchCondition3 = SearchCondition3.DisplayStringContains(Crane)
        MySearchConditionGroup1.Add(SearchCondition1)
        MySearchConditionGroup2.Add(SearchCondition2)
        MySearchConditionGroup3.Add(SearchCondition3)


        MySearch1 = New Search
        MySearch1.Selection.SelectAll()
        MySearch1.SearchConditions.AddGroup(MySearchConditionGroup1)
        SelectionModelItems1 = MySearch1.FindAll(oDoc, False)

        MySearch2 = New Search
        MySearch2.Selection.SelectAll()
        MySearch2.SearchConditions.AddGroup(MySearchConditionGroup2)
        SelectionModelItems2 = MySearch2.FindAll(oDoc, False)

        MySearch3 = New Search
        MySearch3.Selection.SelectAll()
        MySearch3.SearchConditions.AddGroup(MySearchConditionGroup3)
        SelectionModelItems3 = MySearch3.FindAll(oDoc, False)


        NewOverrideTrans = New Transform3D(NewIndentityM, newvector3d)
        ActiveDocument.Models.OverridePermanentTransform(SelectionModelItems1, NewOverrideTrans, True)
        ActiveDocument.Models.OverridePermanentTransform(SelectionModelItems2, NewOverrideTrans, True)
        ActiveDocument.Models.OverridePermanentTransform(SelectionModelItems3, NewOverrideTrans, True)
    End Sub

#End Region


    Public Sub ReHoistUp(ByVal row As Integer, ByRef myDatatable As DataTable)

        oDoc = ActiveDocument
        Dim newvector3d As Vector3D
        Dim NewOverrideTrans As Transform3D
        'Dim Rotation As Rotation3D
        Dim NewIndentityM As Matrix3



        Dim h As Double = -(CDbl(myDatatable.Rows(row).Item("Z")))

        newvector3d = New Vector3D(0, 0, h)

        Dim SelectionModelItems1 As New ModelItemCollection
        Dim SelectionModelItems2 As New ModelItemCollection


        NewIndentityM = New Matrix3



        Dim MySearch1 As Search
        Dim MySearch2 As Search
        Dim MySearchConditionGroup1 As New List(Of SearchCondition)
        Dim MySearchConditionGroup2 As New List(Of SearchCondition)

        Dim SearchCondition1 As SearchCondition
        SearchCondition1 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Name")

        Dim SearchCondition2 As SearchCondition
        SearchCondition2 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Name")

        Dim Envelope As String = myDatatable.Rows(row).Item("ObjectName")
        Dim ObjectNum As String = Right(Envelope, 12)
        Dim Rigging As String = "Rigging" & ObjectNum
        SearchCondition1 = SearchCondition1.DisplayStringContains(Envelope) '  "Rigging2710-PR-525P"
        SearchCondition2 = SearchCondition2.DisplayStringContains(Rigging)
        MySearchConditionGroup1.Add(SearchCondition1)
        MySearchConditionGroup2.Add(SearchCondition2)


        MySearch1 = New Search
        MySearch1.Selection.SelectAll()
        MySearch1.SearchConditions.AddGroup(MySearchConditionGroup1)
        SelectionModelItems1 = MySearch1.FindAll(oDoc, False)

        MySearch2 = New Search
        MySearch2.Selection.SelectAll()
        MySearch2.SearchConditions.AddGroup(MySearchConditionGroup2)
        SelectionModelItems2 = MySearch2.FindAll(oDoc, False)


        NewOverrideTrans = New Transform3D(NewIndentityM, newvector3d)
        ActiveDocument.Models.OverridePermanentTransform(SelectionModelItems1, NewOverrideTrans, True)
        ActiveDocument.Models.OverridePermanentTransform(SelectionModelItems2, NewOverrideTrans, True)
    End Sub

    Public Sub RerotateObjAlongAxis(ByVal row As Integer, ByRef myDatatable As DataTable)

        oDoc = ActiveDocument


        Dim SelectionModelItems1 As New ModelItemCollection
        Dim SelectionModelItems2 As New ModelItemCollection

        Dim MySearch1 As Search
        Dim MySearch2 As Search
        Dim MySearchConditionGroup1 As New List(Of SearchCondition)
        Dim MySearchConditionGroup2 As New List(Of SearchCondition)

        Dim SearchCondition1 As SearchCondition
        SearchCondition1 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Name")

        Dim SearchCondition2 As SearchCondition
        SearchCondition2 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Name")
        ''''
        Dim Envelope As String = myDatatable.Rows(row).Item("ObjectName")
        Dim ObjectNum As String = Right(Envelope, 12)
        Dim Rigging As String = "Rigging" & ObjectNum
        SearchCondition1 = SearchCondition1.DisplayStringContains(Envelope) '  "Rigging2710-PR-525P"
        SearchCondition2 = SearchCondition2.DisplayStringContains(Rigging)
        MySearchConditionGroup1.Add(SearchCondition1)
        MySearchConditionGroup2.Add(SearchCondition2)


        MySearch1 = New Search
        MySearch1.Selection.SelectAll()
        MySearch1.SearchConditions.AddGroup(MySearchConditionGroup1)
        SelectionModelItems1 = MySearch1.FindAll(oDoc, False)

        MySearch2 = New Search
        MySearch2.Selection.SelectAll()
        MySearch2.SearchConditions.AddGroup(MySearchConditionGroup2)
        SelectionModelItems2 = MySearch2.FindAll(oDoc, False)

        Dim R As Double = -CDbl(myDatatable.Rows(row).Item("Rotation"))

        Dim oCenterP As Point3D = SelectionModelItems1.BoundingBox.Center
        Dim oMoveBackToOrig As Vector3D = New Vector3D(-oCenterP.X, -oCenterP.Y, -oCenterP.Z)

        Dim oMoveBackToOrigM As Transform3D = Transform3D.CreateTranslation(oMoveBackToOrig)

        '  set the axis we will rotate around （0,0,1）

        Dim odeltaA As UnitVector3D = New UnitVector3D(0, 0, 1)

        ' Create delta of Quaternion: axis is Z,       

        Dim delta As Rotation3D = New Rotation3D(odeltaA, (R * Math.PI / 180))

        'create a transform from a matrix with a rotation.

        Dim oNewOverrideTrans As Transform3D = New Transform3D(delta)


        Dim oFinalM As Transform3D = Transform3D.Multiply(oMoveBackToOrigM, oNewOverrideTrans)

        oFinalM = Transform3D.Multiply(oFinalM, oMoveBackToOrigM.Inverse())

        oDoc.Models.OverridePermanentTransform(SelectionModelItems1, oFinalM, True)
        oDoc.Models.OverridePermanentTransform(SelectionModelItems2, oFinalM, True)
    End Sub

    Public Sub RerotateCraneAlongZ(ByVal row As Integer, ByRef myDatatable As DataTable)

        oDoc = ActiveDocument


        Dim SelectionModelItems1 As New ModelItemCollection
        Dim SelectionModelItems2 As New ModelItemCollection
        Dim SelectionModelItems3 As New ModelItemCollection
        Dim SelectionModelItems4 As New ModelItemCollection
        Dim SelectionModelItems5 As New ModelItemCollection
        Dim SelectionModelItems6 As New ModelItemCollection
        Dim SelectionModelItems7 As New ModelItemCollection
        Dim SelectionModelItems8 As New ModelItemCollection
        Dim SelectionModelItems9 As New ModelItemCollection
        Dim SelectionModelItems10 As New ModelItemCollection

        Dim MySearch1 As Search
        Dim MySearch2 As Search
        Dim MySearch3 As Search


        Dim MySearchConditionGroup1 As New List(Of SearchCondition)
        Dim MySearchConditionGroup2 As New List(Of SearchCondition)
        Dim MySearchConditionGroup3 As New List(Of SearchCondition)


        Dim SearchCondition1 As SearchCondition
        SearchCondition1 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Name")

        Dim SearchCondition2 As SearchCondition
        SearchCondition2 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Name")

        Dim SearchCondition3 As SearchCondition
        SearchCondition3 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Name")



        Dim SearchCondition8 As SearchCondition
        SearchCondition8 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Type")




        Dim Envelope As String = myDatatable.Rows(row).Item("ObjectName")
        Dim ObjectNum As String = Right(Envelope, 12)
        Dim Rigging As String = "Rigging" & ObjectNum
        Dim Crane As String = "Crane" & ObjectNum

        SearchCondition1 = SearchCondition1.DisplayStringContains(Envelope)
        SearchCondition2 = SearchCondition2.DisplayStringContains(Rigging)
        SearchCondition3 = SearchCondition3.DisplayStringContains(Crane)


        MySearchConditionGroup1.Add(SearchCondition1)
        MySearchConditionGroup2.Add(SearchCondition2)
        MySearchConditionGroup3.Add(SearchCondition3)



        MySearch1 = New Search
        MySearch1.Selection.SelectAll()
        MySearch1.SearchConditions.AddGroup(MySearchConditionGroup1)
        SelectionModelItems1 = MySearch1.FindAll(oDoc, False)

        MySearch2 = New Search
        MySearch2.Selection.SelectAll()
        MySearch2.SearchConditions.AddGroup(MySearchConditionGroup2)
        SelectionModelItems2 = MySearch2.FindAll(oDoc, False)

        MySearch3 = New Search
        MySearch3.Selection.SelectAll()
        MySearch3.SearchConditions.AddGroup(MySearchConditionGroup3)
        SelectionModelItems3 = MySearch3.FindAll(oDoc, False)


        For Each ModelItem In SelectionModelItems3
            For Each insert In ModelItem.Children
                For Each block In insert.Children
                    If block.DisplayName = "Boom_CC 2800" Then
                        SelectionModelItems4.Add(block)
                    End If
                Next
            Next
        Next

        For Each ModelItem In SelectionModelItems3
            For Each insert In ModelItem.Children
                For Each block In insert.Children
                    If block.DisplayName = "Mast_CC 2800" Then
                        SelectionModelItems5.Add(block)
                    End If
                Next
            Next
        Next

        For Each ModelItem In SelectionModelItems3
            For Each insert In ModelItem.Children
                For Each block In insert.Children
                    If block.DisplayName = "Superlift_CC 2800" Then
                        SelectionModelItems6.Add(block)
                    End If
                Next
            Next
        Next

        For Each ModelItem In SelectionModelItems3
            For Each insert In ModelItem.Children
                For Each block In insert.Children
                    If block.DisplayName = "LoadLines_CC 2800" Then
                        SelectionModelItems7.Add(block)
                    End If
                Next
            Next
        Next

        For Each ModelItem In SelectionModelItems3
            For Each block In ModelItem.Children
                If block.ClassDisplayName = "Line" Then
                    SelectionModelItems8.Add(block)
                End If
            Next
        Next



        For Each ModelItem In SelectionModelItems3
            For Each insert In ModelItem.Children
                For Each block In insert.Children
                    If block.DisplayName = "Body_CC 2800" Then
                        SelectionModelItems9.Add(block)
                    End If
                Next
            Next
        Next

        For Each ModelItem In SelectionModelItems3
            For Each insert In ModelItem.Children
                For Each block In insert.Children
                    If block.DisplayName = "Crawler_CC 2800" Then
                        SelectionModelItems10.Add(block)
                    End If
                Next
            Next
        Next



        Dim oCenterP As Point3D = SelectionModelItems10.BoundingBox.Center()
        Dim oMinP As Point3D = SelectionModelItems10.BoundingBox.Min()
        Dim oMaxP As Point3D = SelectionModelItems10.BoundingBox.Max()
        Dim oMoveBackToOrig As Vector3D = New Vector3D(-oCenterP.X, -oCenterP.Y, -oCenterP.Z)

        Dim oMoveBackToOrigM As Transform3D = Transform3D.CreateTranslation(oMoveBackToOrig)

        '  set the axis we will rotate around （0,0,1）

        Dim odeltaA As UnitVector3D = New UnitVector3D(0, 0, 1)

        ' Create delta of Quaternion: axis is Z,       
        Dim RB As Double = -CDbl(myDatatable.Rows(row).Item("Rotation"))
        Dim delta As Rotation3D = New Rotation3D(odeltaA, RB * (Math.PI / 180)) ' -0.785

        'create a transform from a matrix with a rotation.

        Dim oNewOverrideTrans As Transform3D = New Transform3D(delta)


        Dim oFinalM As Transform3D = Transform3D.Multiply(oMoveBackToOrigM, oNewOverrideTrans)

        oFinalM = Transform3D.Multiply(oFinalM, oMoveBackToOrigM.Inverse())



        oDoc.Models.OverridePermanentTransform(SelectionModelItems1, oFinalM, True)
        oDoc.Models.OverridePermanentTransform(SelectionModelItems2, oFinalM, True)
        'oDoc.Models.OverridePermanentTransform(SelectionModelItems3, oFinalM, True)
        oDoc.Models.OverridePermanentTransform(SelectionModelItems4, oFinalM, True)
        oDoc.Models.OverridePermanentTransform(SelectionModelItems5, oFinalM, True)
        oDoc.Models.OverridePermanentTransform(SelectionModelItems6, oFinalM, True)
        oDoc.Models.OverridePermanentTransform(SelectionModelItems7, oFinalM, True)
        oDoc.Models.OverridePermanentTransform(SelectionModelItems8, oFinalM, True)
        oDoc.Models.OverridePermanentTransform(SelectionModelItems9, oFinalM, True)
    End Sub

    Public Sub ReboomUp(ByVal row As Integer, ByRef myDatatable As DataTable)

        oDoc = ActiveDocument


        Dim SelectionModelItems1 As New ModelItemCollection
        Dim SelectionModelItems2 As New ModelItemCollection
        Dim SelectionModelItems3 As New ModelItemCollection
        Dim SelectionModelItems4 As New ModelItemCollection
        Dim SelectionModelItems5 As New ModelItemCollection
        Dim SelectionModelItems6 As New ModelItemCollection
        Dim SelectionModelItems7 As New ModelItemCollection
        Dim SelectionModelItems8 As New ModelItemCollection
        Dim SelectionModelItems9 As New ModelItemCollection
        Dim SelectionModelItems10 As New ModelItemCollection

        Dim MySearch1 As Search
        Dim MySearch2 As Search
        Dim MySearch3 As Search


        Dim MySearchConditionGroup1 As New List(Of SearchCondition)
        Dim MySearchConditionGroup2 As New List(Of SearchCondition)
        Dim MySearchConditionGroup3 As New List(Of SearchCondition)


        Dim SearchCondition1 As SearchCondition
        SearchCondition1 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Name")

        Dim SearchCondition2 As SearchCondition
        SearchCondition2 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Name")

        Dim SearchCondition3 As SearchCondition
        SearchCondition3 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Name")



        Dim SearchCondition8 As SearchCondition
        SearchCondition8 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Type")




        Dim Envelope As String = myDatatable.Rows(row).Item("ObjectName")
        Dim ObjectNum As String = Right(Envelope, 12)
        Dim Rigging As String = "Rigging" & ObjectNum
        Dim Crane As String = "Crane" & ObjectNum

        SearchCondition1 = SearchCondition1.DisplayStringContains(Envelope)
        SearchCondition2 = SearchCondition2.DisplayStringContains(Rigging)
        SearchCondition3 = SearchCondition3.DisplayStringContains(Crane)


        MySearchConditionGroup1.Add(SearchCondition1)
        MySearchConditionGroup2.Add(SearchCondition2)
        MySearchConditionGroup3.Add(SearchCondition3)



        MySearch1 = New Search
        MySearch1.Selection.SelectAll()
        MySearch1.SearchConditions.AddGroup(MySearchConditionGroup1)
        SelectionModelItems1 = MySearch1.FindAll(oDoc, False)

        MySearch2 = New Search
        MySearch2.Selection.SelectAll()
        MySearch2.SearchConditions.AddGroup(MySearchConditionGroup2)
        SelectionModelItems2 = MySearch2.FindAll(oDoc, False)

        MySearch3 = New Search
        MySearch3.Selection.SelectAll()
        MySearch3.SearchConditions.AddGroup(MySearchConditionGroup3)
        SelectionModelItems3 = MySearch3.FindAll(oDoc, False)


        For Each ModelItem In SelectionModelItems3
            For Each insert In ModelItem.Children
                For Each block In insert.Children
                    If block.DisplayName = "Boom_CC 2800" Then
                        SelectionModelItems4.Add(block)
                    End If
                Next
            Next
        Next



        For Each ModelItem In SelectionModelItems3
            For Each insert In ModelItem.Children
                For Each block In insert.Children
                    If block.DisplayName = "LoadLines_CC 2800" Then
                        SelectionModelItems7.Add(block)
                    End If
                Next
            Next
        Next

        For Each ModelItem In SelectionModelItems3
            For Each block In ModelItem.Children
                If block.ClassDisplayName = "Line" Then
                    SelectionModelItems8.Add(block)
                End If
            Next
        Next



        For Each ModelItem In SelectionModelItems3
            For Each insert In ModelItem.Children
                For Each block In insert.Children
                    If block.DisplayName = "Crawler_CC 2800" Then
                        SelectionModelItems10.Add(block)
                    End If
                Next
            Next
        Next



        Dim oCenterP As Point3D = SelectionModelItems4.BoundingBox.Center()
        Dim oMinP As Point3D = SelectionModelItems4.BoundingBox.Min()
        Dim oMaxP As Point3D = SelectionModelItems4.BoundingBox.Max()
        Dim oMoveBackToOrig As Vector3D = New Vector3D(-oMinP.X, -oMinP.Y, -oMinP.Z)

        Dim oMoveBackToOrigM As Transform3D = Transform3D.CreateTranslation(oMoveBackToOrig)

        '  set the axis we will rotate around （0,0,1）

        Dim odeltaA As UnitVector3D = New UnitVector3D(0, 1, 0)

        ' Create delta of Quaternion: axis is Z,       
        Dim RB As Double = -CDbl(myDatatable.Rows(row).Item("Rotation"))
        Dim delta As Rotation3D = New Rotation3D(odeltaA, RB * (Math.PI / 180)) ' -0.785

        'create a transform from a matrix with a rotation.

        Dim oNewOverrideTrans As Transform3D = New Transform3D(delta)


        Dim oFinalM As Transform3D = Transform3D.Multiply(oMoveBackToOrigM, oNewOverrideTrans)

        oFinalM = Transform3D.Multiply(oFinalM, oMoveBackToOrigM.Inverse())


        ''''''''
        Dim oCenterP2 As Point3D = SelectionModelItems7.BoundingBox.Center()
        Dim oMinP2 As Point3D = SelectionModelItems7.BoundingBox.Min()
        Dim oMaxP2 As Point3D = SelectionModelItems7.BoundingBox.Max()
        Dim oMoveBackToOrig2 As Vector3D = New Vector3D(-oMaxP2.X, -oMaxP2.Y, -oMaxP2.Z)

        Dim oMoveBackToOrigM2 As Transform3D = Transform3D.CreateTranslation(oMoveBackToOrig2)

        '  set the axis we will rotate around （0,0,1）

        Dim odeltaA2 As UnitVector3D = New UnitVector3D(0, 1, 0)

        ' Create delta of Quaternion: axis is Z,       
        Dim RB2 As Double = myDatatable.Rows(row).Item("Rotation")
        Dim delta2 As Rotation3D = New Rotation3D(odeltaA2, RB2 * (Math.PI / 180)) ' -0.785

        'create a transform from a matrix with a rotation.

        Dim oNewOverrideTrans2 As Transform3D = New Transform3D(delta2)


        Dim oFinalM2 As Transform3D

        oFinalM2 = oFinalM
        oFinalM2 = Transform3D.Multiply(oFinalM2, oMoveBackToOrigM2)
        oFinalM2 = Transform3D.Multiply(oFinalM2, oNewOverrideTrans2)
        oFinalM2 = Transform3D.Multiply(oFinalM2, oMoveBackToOrigM2.Inverse())
        'oFinalM2 = Transform3D.Multiply(oFinalM, oFinalM2)


        oDoc.Models.OverridePermanentTransform(SelectionModelItems1, oFinalM2, True)
        oDoc.Models.OverridePermanentTransform(SelectionModelItems2, oFinalM2, True)
        'oDoc.Models.OverridePermanentTransform(SelectionModelItems3, oFinalM, True)
        oDoc.Models.OverridePermanentTransform(SelectionModelItems4, oFinalM, True)
        'oDoc.Models.OverridePermanentTransform(SelectionModelItems5, oFinalM, True)
        'oDoc.Models.OverridePermanentTransform(SelectionModelItems6, oFinalM, True)
        oDoc.Models.OverridePermanentTransform(SelectionModelItems7, oFinalM2, True)
        oDoc.Models.OverridePermanentTransform(SelectionModelItems8, oFinalM, True)
        'oDoc.Models.OverridePermanentTransform(SelectionModelItems9, oFinalM, True)
    End Sub

    Public Sub Rewalk(ByVal row As Integer, ByRef myDatatable As DataTable)

        oDoc = ActiveDocument
        Dim newvector3d As Vector3D
        Dim NewOverrideTrans As Transform3D
        'Dim Rotation As Rotation3D
        Dim NewIndentityM As Matrix3



        Dim wx As Double = -CDbl(myDatatable.Rows(row).Item("X"))
        Dim wy As Double = -CDbl(myDatatable.Rows(row).Item("Y"))
        newvector3d = New Vector3D(wx, wy, 0)

        Dim SelectionModelItems1 As New ModelItemCollection
        Dim SelectionModelItems2 As New ModelItemCollection
        Dim SelectionModelItems3 As New ModelItemCollection


        NewIndentityM = New Matrix3



        Dim MySearch1 As Search
        Dim MySearch2 As Search
        Dim MySearch3 As Search
        Dim MySearchConditionGroup1 As New List(Of SearchCondition)
        Dim MySearchConditionGroup2 As New List(Of SearchCondition)
        Dim MySearchConditionGroup3 As New List(Of SearchCondition)

        Dim SearchCondition1 As SearchCondition
        SearchCondition1 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Name")

        Dim SearchCondition2 As SearchCondition
        SearchCondition2 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Name")

        Dim SearchCondition3 As SearchCondition
        SearchCondition3 = SearchCondition.HasPropertyByDisplayName _
                ("Item", "Name")

        Dim Envelope As String = myDatatable.Rows(row).Item("ObjectName")
        Dim ObjectNum As String = Right(Envelope, 12)
        Dim Rigging As String = "Rigging" & ObjectNum
        Dim Crane As String = "Crane" & ObjectNum

        SearchCondition1 = SearchCondition1.DisplayStringContains(Envelope) '  "Rigging2710-PR-525P"
        SearchCondition2 = SearchCondition2.DisplayStringContains(Rigging)
        SearchCondition3 = SearchCondition3.DisplayStringContains(Crane)
        MySearchConditionGroup1.Add(SearchCondition1)
        MySearchConditionGroup2.Add(SearchCondition2)
        MySearchConditionGroup3.Add(SearchCondition3)


        MySearch1 = New Search
        MySearch1.Selection.SelectAll()
        MySearch1.SearchConditions.AddGroup(MySearchConditionGroup1)
        SelectionModelItems1 = MySearch1.FindAll(oDoc, False)

        MySearch2 = New Search
        MySearch2.Selection.SelectAll()
        MySearch2.SearchConditions.AddGroup(MySearchConditionGroup2)
        SelectionModelItems2 = MySearch2.FindAll(oDoc, False)

        MySearch3 = New Search
        MySearch3.Selection.SelectAll()
        MySearch3.SearchConditions.AddGroup(MySearchConditionGroup3)
        SelectionModelItems3 = MySearch3.FindAll(oDoc, False)


        NewOverrideTrans = New Transform3D(NewIndentityM, newvector3d)
        ActiveDocument.Models.OverridePermanentTransform(SelectionModelItems1, NewOverrideTrans, True)
        ActiveDocument.Models.OverridePermanentTransform(SelectionModelItems2, NewOverrideTrans, True)
        ActiveDocument.Models.OverridePermanentTransform(SelectionModelItems3, NewOverrideTrans, True)
    End Sub

End Class
