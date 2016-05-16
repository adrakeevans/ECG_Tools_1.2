Imports Microsoft.Office.Core.MsoThemeColorIndex
Imports Microsoft.Office.Core.MsoTriState
Imports Microsoft.Office.Interop.PowerPoint.PpBorderType
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.PowerPoint


Public Class ThisAddIn
    Public mTable As PowerPoint.Table
    Public tMaster As Object
    Public shpRange As PowerPoint.ShapeRange
    Public sldRange As PowerPoint.SlideRange




    Private Sub ThisAddIn_Startup() Handles Me.Startup



    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub


    Sub SetMasterTable()

        On Error Resume Next
        mTable = Application.ActiveWindow.Selection.ShapeRange(1).Table
        tMaster = Application.ActiveWindow.Selection.ShapeRange(1)

    End Sub

    Sub FormatDrakeTables()
        Dim TableTop As Double

        Dim ThisTable As PowerPoint.Table

        Dim FirstRowCheck As String

        Dim reFormatCheck As Boolean
        Dim variant1 As Object


        FirstRowCheck = ""
        On Error Resume Next

        If Application.ActiveWindow.Selection.ShapeRange.Count = 2 Then
            ThisTable = Application.ActiveWindow.Selection.ShapeRange(2).Table
            FirstRowCheck = Application.ActiveWindow.Selection.ShapeRange(1).TextFrame.TextRange.Text
            FirstRowCheck = Replace(FirstRowCheck, Chr(11), "")
            FirstRowCheck = Replace(FirstRowCheck, Chr(13), "")
            Application.ActiveWindow.Selection.ShapeRange(1).Delete()
        Else
            ThisTable = Application.ActiveWindow.Selection.ShapeRange(1).Table
        End If
        With ThisTable
            variant1 = .Parent.Table.Cell(1, 1)
            TableTop = .Parent.Top

            If Not FirstRowCheck = "" Then
                .Rows.Add(1)
                .Cell(1, 1).Shape.TextFrame.TextRange.Text = FirstRowCheck
            End If

            If .Rows(1).Height = .Rows(2).Height Then
                reFormatCheck = True
            Else
                reFormatCheck = False
            End If

            'Remove outermost borders from table
            Dim clm As Object
            For Each rw In .Rows
                rw.Cells.Borders(ppBorderTop).Visible = msoFalse
                rw.Cells.Borders(ppBorderBottom).Visible = msoFalse
                rw.Cells.Borders(ppBorderLeft).Visible = msoFalse
                rw.Cells.Borders(ppBorderRight).Visible = msoFalse
            Next rw

            For Each clm In .Columns
                clm.Cells.Borders(ppBorderTop).Visible = msoFalse
                clm.Cells.Borders(ppBorderBottom).Visible = msoFalse
                clm.Cells.Borders(ppBorderLeft).Visible = msoFalse
                clm.Cells.Borders(ppBorderRight).Visible = msoFalse


                For Each Cll In clm.Cells
                    'Add white borders around each cell

                    With Cll.Borders(ppBorderTop)
                        .ForeColor.RGB = RGB(255, 255, 255)
                        .Weight = 1
                    End With
                    With Cll.Borders(ppBorderBottom)
                        .ForeColor.RGB = RGB(255, 255, 255)
                        .Weight = 1
                    End With
                    With Cll.Borders(ppBorderLeft)
                        .ForeColor.RGB = RGB(255, 255, 255)
                        .Weight = 1
                    End With
                    With Cll.Borders(ppBorderRight)
                        .ForeColor.RGB = RGB(255, 255, 255)
                        .Weight = 1
                    End With


                    'Fill in cells with Accent6 and set font color to Black
                    With Cll.Shape

                        .Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent6
                        .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)

                    End With

                Next Cll

            Next clm


            'Format 1st Column Shading
            For Each Cll In .Columns(1).Cells

                Cll.Shape.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent5

            Next Cll

            'FIRST ROW
            'Format first row to be orange title of Table, make small
            .Rows(1).Cells(1).Merge(MergeTo:= .Cell(1, .Rows(1).Cells.Count))

            With .Rows(1).Cells(1)
                With .Shape
                    .Fill.Visible = msoFalse

                    With .TextFrame
                        .TextRange.Font.Color.ObjectThemeColor = msoThemeColorAccent1
                        .TextRange.Font.Size = 11
                        .TextRange.Font.Bold = msoTrue
                        .MarginBottom = 0
                        .MarginTop = 0
                        .MarginLeft = 7.2
                        .TextRange.Text = Replace(.TextRange.Text, Chr(13), "")
                        .TextRange.Text = Replace(.TextRange.Text, Chr(11), "")
                    End With

                End With
                .Borders(ppBorderBottom).Visible = msoFalse
                .Borders(ppBorderBottom).Weight = 0

            End With
            .Rows(1).Height = 0
            .Rows(1).Cells(1).Shape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle
            .Rows(1).Cells(1).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft


            'LAST ROW
            'if last row is true then format it the same way as 2nd row
            If .LastRow = True Then
                For Each Cll In .Rows(.Rows.Count).Cells
                    With Cll.Shape

                        .Fill.Visible = msoTrue
                        .Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent2

                        With .TextFrame.TextRange.Font
                            .Color.ObjectThemeColor = msoThemeColorAccent5
                            .Bold = msoTrue
                        End With

                    End With
                Next Cll
            End If

            'SECOND ROW
            'Format 2nd Row
            For Each Cll In .Rows(2).Cells
                With Cll.Shape

                    .Fill.Visible = msoTrue
                    .Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent2

                    With .TextFrame.TextRange.Font
                        .Color.ObjectThemeColor = msoThemeColorAccent5
                        .Bold = msoTrue
                    End With

                End With

            Next Cll

            'Realign Position of Table
            If reFormatCheck = True Then
                .Parent.Top = TableTop - .Rows(1).Height
            Else
                .Parent.Top = TableTop
            End If


        End With



    End Sub

    Sub CopyTablePositionFormatting()

        With Application.ActiveWindow.Selection.ShapeRange(1).Table
            .Parent.Top = mTable.Parent.Top
            .Parent.Left = mTable.Parent.Left

            Dim clmIndex As Integer
            clmIndex = 0
            For Each clm In .Columns
                If clmIndex < mTable.Columns.Count Then
                    clmIndex = clmIndex + 1
                End If

                clm.Width = mTable.Columns(clmIndex).Width



                Dim cllIndex As Integer = 0
                For Each Cll In clm.Cells
                    If cllIndex < clm.Cells.Count Then
                        cllIndex = cllIndex + 1
                    End If

                    With Cll.Shape.TextFrame
                        .MarginLeft = mTable.Columns(clmIndex).Cells(cllIndex).Shape.TextFrame.MarginLeft
                        .MarginRight = mTable.Columns(clmIndex).Cells(cllIndex).Shape.TextFrame.MarginRight
                    End With
                Next Cll

            Next clm

            Dim rwIndex As Integer

            rwIndex = 0
            For Each rw In .Rows
                If rwIndex < mTable.Rows.Count Then
                    rwIndex = rwIndex + 1
                End If

                rw.Height = mTable.Rows(rwIndex).Height

                Dim cllIndex As Integer = 0
                For Each Cll In rw.Cells
                    If cllIndex < rw.Cells.Count Then
                        cllIndex = cllIndex + 1
                    End If

                    With Cll.Shape.TextFrame
                        .MarginBottom = mTable.Rows(rwIndex).Cells(cllIndex).Shape.TextFrame.MarginBottom
                        .MarginTop = mTable.Rows(rwIndex).Cells(cllIndex).Shape.TextFrame.MarginTop

                        With .TextRange.Font
                            .Name = mTable.Rows(rwIndex).Cells(cllIndex).Shape.TextFrame.TextRange.Font.Name
                            .Size = mTable.Rows(rwIndex).Cells(cllIndex).Shape.TextFrame.TextRange.Font.Size
                            .Color = mTable.Rows(rwIndex).Cells(cllIndex).Shape.TextFrame.TextRange.Font.Color
                        End With

                    End With

                Next Cll

            Next rw
        End With

    End Sub

    Sub AutoSizeTableHeight()


        Dim shp As PowerPoint.Shape
        Dim rw As PowerPoint.Row

        For Each shp In Application.ActiveWindow.Selection.ShapeRange

            For Each rw In shp.Table.Rows

                rw.Height = 0

            Next rw
        Next shp

    End Sub

    Sub SetTop()

        Application.ActiveWindow.Selection.ShapeRange(1).Top = tMaster.Top

    End Sub

    Sub SetLeft()

        Application.ActiveWindow.Selection.ShapeRange(1).Left = tMaster.Left

    End Sub

    Sub SetSizeShape()

        With Application.ActiveWindow.Selection.ShapeRange(1)
            .Top = tMaster.Top
            .Left = tMaster.Left
            .Width = tMaster.Width
        End With

    End Sub

    Sub SetColumnWidths()

        Dim myTable As Object

        myTable = Application.ActiveWindow.Selection.ShapeRange(1)

        On Error Resume Next
        With myTable

            For Each clm In .Columns

                clm.Width = .tMaster.Columns(clm.Index).Width

            Next clm
        End With

    End Sub

    Sub SetTBMargins(TBMargin As Double)

        Dim myTable As PowerPoint.Table
        Dim myShape As PowerPoint.Shape

        myShape = Application.ActiveWindow.Selection.ShapeRange(1)
        myTable = Application.ActiveWindow.Selection.ShapeRange(1).Table


        If myShape.HasTable = MsoTriState.msoFalse Then

            With myShape
                .TextFrame.MarginTop = TBMargin
                .TextFrame.MarginBottom = TBMargin
            End With

        ElseIf myShape.HasTable = MsoTriState.msoTrue Then

            With myTable
                Dim i As Integer
                Dim j As Integer

                For i = 1 To .Rows.Count

                    For j = 1 To .Columns.Count

                        With .Cell(i, j)

                            If .Selected Then
                                With .Shape.TextFrame
                                    .MarginBottom = TBMargin
                                    .MarginTop = TBMargin
                                End With
                            End If
                        End With
                    Next j
                Next i
            End With
        End If

    End Sub

    Sub SetLRMargins(LRMargin As Double)

        Dim myTable As PowerPoint.Table
        Dim myShape As PowerPoint.Shape

        myShape = Application.ActiveWindow.Selection.ShapeRange(1)
        myTable = Application.ActiveWindow.Selection.ShapeRange(1).Table




        If myShape.HasTable = MsoTriState.msoFalse Then

            With myShape
                .TextFrame.MarginLeft = LRMargin
                .TextFrame.MarginRight = LRMargin
            End With

        ElseIf myShape.HasTable = MsoTriState.msoTrue Then

            With myTable
                Dim i As Integer
                Dim j As Integer

                For i = 1 To .Rows.Count

                    For j = 1 To .Columns.Count

                        With .Cell(i, j)

                            If .Selected Then
                                With .Shape.TextFrame
                                    .MarginLeft = LRMargin
                                    .MarginRight = LRMargin
                                End With
                            End If
                        End With
                    Next j
                Next i
            End With
        End If


    End Sub

    Sub InsertWingdings(ByVal charNumber As Integer)
        Dim i As Integer
        Dim myRange As TextRange
        Dim wdNumber As Integer

        myRange = Application.ActiveWindow.Selection.TextRange

        If charNumber < 10 Then

            myRange.InsertSymbol("+Body", 10102 + charNumber, msoTrue)

        ElseIf charNumber = 10 Then

            For i = 0 To 9

                wdNumber = 10102 + i

                Application.ActiveWindow.Selection.TextRange.InsertSymbol("+Body", wdNumber, msoTrue)

            Next i

        End If



    End Sub

    Sub UpdateSelection()
        On Error Resume Next
        shpRange = Application.ActiveWindow.Selection.ShapeRange
        sldRange = Application.ActiveWindow.Selection.SlideRange

    End Sub

    Sub InsertElement(ByVal fName As String)
        Dim fileName As String
        Dim myPresentation As Presentation
        Dim currentSlide As Slide

        fileName = "C:\Users\adrak\AppData\Local\Efficient Elements\Efficient Elements for presentations\User\MyElements\1ea8a3dd-ba1c-40c2-b498-e49ebe7af483.pptx"

        currentSlide = Application.ActiveWindow.View.Slide

        myPresentation = Application.Presentations.Open(fileName, msoTrue, msoTrue, msoFalse)


        Dim shp As PowerPoint.Shapes
        shp = myPresentation.Slides(1).Shapes

        Dim i As Integer

        For i = 1 To shp.Count

            If IsError(shp(i).PlaceholderFormat.Type) Then
                shp(i).Copy()
                currentSlide.Shapes.Paste()
            End If
        Next i
    End Sub

    Private Sub Application_PresentationNewSlide(Sld As Slide) Handles Application.PresentationNewSlide

        'MsgBox("Success")

    End Sub

    Private Sub Application_AfterDragandDrop(Sld As Slide, x As Single, y As Single) Handles Application.AfterDragDropOnSlide

        MsgBox("AfterDragDrop")

    End Sub

    Private Sub Application_AfterShapeSizeChange(Sld As Slide) Handles Application.AfterShapeSizeChange
        MsgBox("AfterShapeChange")
    End Sub


End Class
