Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Interop.PowerPoint
Imports System.Xml

Public Class Ribbon1
    Public LRMargin As Double
    Public TBMargin As Double
    Public tagName As String
    Public tagValue As String
    Public shpName As String
    Public sldName As String
    Private elementSource As Shape
    Private elementPageBreakHeader As Shape
    Private element16ptText As Shape
    Private myPPT As New Application


    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        ddWingdings.SelectedItemIndex = -1


        Dim newItem As RibbonDropDownItem
        Dim reader As XmlReader = XmlReader.Create("C:\Users\adrak\AppData\Local\Efficient Elements\Efficient Elements for presentations\User\MyElements.xml")
        Dim output As New List(Of String)

        While reader.Read
            If reader.HasAttributes And reader.Name = "element" Then
                reader.MoveToNextAttribute()
                If reader.Name = "name" Then
                    output.Add(New String(reader.Value))
                End If
            End If

        End While

        Dim i As Integer
        For i = 0 To output.Count - 1
            newItem = Globals.Factory.GetRibbonFactory.CreateRibbonDropDownItem()
            newItem.Label = output(i)

            ddElements.Items.Add(newItem)


        Next




    End Sub

    Private Sub ebTBMargin_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles ebTBMargin.TextChanged

        TBMargin = CDbl(ebTBMargin.Text) * 2800

    End Sub

    Private Sub btnLRMargin_Click(sender As Object, e As RibbonControlEventArgs) Handles btnLRMargin.Click

        Globals.ThisAddIn.SetLRMargins(LRMargin)

    End Sub

    Private Sub ebLRMargin_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles ebLRMargin.TextChanged

        LRMargin = CDbl(ebLRMargin.Text) * 2800

    End Sub

    Private Sub btnTBMargin_Click(sender As Object, e As RibbonControlEventArgs) Handles btnTBMargin.Click

        Globals.ThisAddIn.SetTBMargins(TBMargin)

    End Sub

    Private Sub ebSuggestions_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles ebSuggestions.TextChanged

        IO.File.AppendAllText("C:\Users\adrak\Documents\Visual Studio 2015\Projects\ECG_Tools\ECG_Tools\ppSuggestions.txt", ebSuggestions.Text & Chr(13) & Chr(11))

        ebSuggestions.Text = ""

    End Sub

    Private Sub btnApplyTag_Click(sender As Object, e As RibbonControlEventArgs) Handles btnApplyTags.Click

        'Globals.ThisAddIn.ApplyTagtoShape(tagName, tagValue)

    End Sub

    Private Sub ebTagType_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles ebTagName.TextChanged

        tagName = ebTagName.Text

    End Sub

    Private Sub ebTagValue_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles ebTagValue.TextChanged

        tagValue = ebTagValue.Text

    End Sub



    Private Sub btnSetMaster_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSetMaster.Click

        'Globals.ThisAddIn.SetMasterTable()



    End Sub

    Private Sub btnShapeSizeAdjus_Click(sender As Object, e As RibbonControlEventArgs)

        Globals.ThisAddIn.SetSizeShape()

    End Sub

    Private Sub btnRowHeight_Click(sender As Object, e As RibbonControlEventArgs) Handles btnAutoRowHeight.Click

        Globals.ThisAddIn.AutoSizeTableHeight()

    End Sub

    Private Sub btnRefreshTag_Click(sender As Object, e As RibbonControlEventArgs) Handles btnRefreshTags.Click

        Globals.ThisAddIn.UpdateSelection()

        ebTagName.Text = Globals.ThisAddIn.shpRange(1).Tags.Name(1)
        ebTagValue.Text = Globals.ThisAddIn.shpRange(1).Tags.Value(1)

    End Sub



    Private Sub btnRefreshName_Click(sender As Object, e As RibbonControlEventArgs) Handles btnRefreshNames.Click

        Globals.ThisAddIn.UpdateSelection()

        ebShpName.Text = Globals.ThisAddIn.shpRange.Name
        ebSldName.Text = Globals.ThisAddIn.sldRange.Name

    End Sub

    Private Sub ddElements_SelectionChanged(sender As Object, e As RibbonControlEventArgs) Handles ddElements.ButtonClick

        Globals.ThisAddIn.UpdateSelection()

        Select Case ddElements.SelectedItem.Label
            Case "Source"
                elementSource = Globals.ThisAddIn.shpRange(1)
            Case "PB Header"
                elementPageBreakHeader = Globals.ThisAddIn.shpRange(1)
            Case "16pt TextBox"
                element16ptText = Globals.ThisAddIn.shpRange(1)
        End Select

    End Sub

    Private Sub btnInsertSource_Click(sender As Object, e As RibbonControlEventArgs) Handles btnInsertSource.Click

        Globals.ThisAddIn.UpdateSelection()

        elementSource.Copy()

        Globals.ThisAddIn.sldRange(1).Shapes.Paste()

    End Sub

    Private Sub btnInsertPBHeader_Click(sender As Object, e As RibbonControlEventArgs) Handles btnInsertPBHeader.Click

        Globals.ThisAddIn.UpdateSelection()

        elementPageBreakHeader.Copy()

        Globals.ThisAddIn.sldRange(1).Shapes.Paste()

    End Sub

    Private Sub btn16ptText_Click(sender As Object, e As RibbonControlEventArgs) Handles btn16ptTextBox.Click

        Globals.ThisAddIn.UpdateSelection()

        element16ptText.Copy()

        Globals.ThisAddIn.sldRange(1).Shapes.Paste()

    End Sub

    Private Sub btnSetTop_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSetTop.Click

        Globals.ThisAddIn.SetTop()

    End Sub

    Private Sub btnSetLeft_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSetLeft.Click

        Globals.ThisAddIn.SetLeft()

    End Sub

    Private Sub btnApplyNames_Click(sender As Object, e As RibbonControlEventArgs) Handles btnApplyNames.Click

        Globals.ThisAddIn.UpdateSelection()

        Globals.ThisAddIn.shpRange(1).Name = ebShpName.Text

    End Sub

    Private Sub ddWingdings_SelectionChanged(sender As Object, e As RibbonControlEventArgs) Handles ddWingdings.SelectionChanged

        Globals.ThisAddIn.InsertWingdings(ddWingdings.SelectedItemIndex)

    End Sub

    Private Sub ddElements_SelectionChanged_1(sender As Object, e As RibbonControlEventArgs) Handles ddElements.SelectionChanged
        InsertElement(ddElements.SelectedItem.Label)

    End Sub

    Private Sub btnHorizontalAnchor_Click(sender As Object, e As RibbonControlEventArgs) Handles btnHorizontalAnchor.Click
        Globals.ThisAddIn.UpdateSelection()
        Dim myshape As Shape

        myshape = Globals.ThisAddIn.shpRange(1)

        With myshape.TextFrame
            Select Case .HorizontalAnchor
                Case Microsoft.Office.Core.MsoHorizontalAnchor.msoAnchorNone
                    .HorizontalAnchor = Microsoft.Office.Core.MsoHorizontalAnchor.msoAnchorCenter
                Case Microsoft.Office.Core.MsoHorizontalAnchor.msoAnchorCenter
                    .HorizontalAnchor = Microsoft.Office.Core.MsoHorizontalAnchor.msoAnchorNone
            End Select
        End With

    End Sub

    Private Sub btnZeroMargin_Click(sender As Object, e As RibbonControlEventArgs) Handles btnZeroMargin.Click

        Dim myshapes As ShapeRange
        myshapes = myPPT.ActiveWindow.Selection.ShapeRange

        Dim shp As Shape

        For Each shp In myshapes
            With shp.TextFrame
                .MarginBottom = 0
                .MarginLeft = 0
                .MarginRight = 0
                .MarginTop = 0
            End With
        Next shp

    End Sub
End Class
