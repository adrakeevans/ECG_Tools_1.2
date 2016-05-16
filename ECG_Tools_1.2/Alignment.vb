Imports Microsoft.Office.Interop.PowerPoint

Public Module Alignment
    Public myPPT As New Application

    Sub aln_Direction(ByVal sDirection As String)
        Dim myShapeRange As PowerPoint.ShapeRange
        Dim alignPosition As Double
        Dim shp As Shape

        myShapeRange = myPPT.ActiveWindow.Selection.ShapeRange

        Select Case sDirection
            Case "Left"
                alignPosition = myShapeRange(myShapeRange.Count).Left

                For Each shp In myShapeRange
                    shp.Left = alignPosition
                Next shp

            Case "Right"
                alignPosition = myShapeRange(myShapeRange.Count).Left + myShapeRange(myShapeRange.Count).Width

                For Each shp In myShapeRange
                    shp.Left = alignPosition + shp.Width
                Next shp

            Case "Top"
                alignPosition = myShapeRange(myShapeRange.Count).Top

                For Each shp In myShapeRange
                    shp.Top = alignPosition
                Next shp

            Case "Bottom"
                alignPosition = myShapeRange(myShapeRange.Count).Top + myShapeRange(myShapeRange.Count).Height

                For Each shp In myShapeRange
                    shp.Top = alignPosition + shp.Height
                Next shp

        End Select


    End Sub

    Sub stretch_Direction(ByVal sDirection As String)
        Dim myShapeRange As ShapeRange
        Dim stretchPosition As Double
        Dim shp As Shape

        myShapeRange = myPPT.ActiveWindow.Selection.ShapeRange

        Select Case sDirection
            Case "Left"
                stretchPosition = myShapeRange(myShapeRange.Count).Left

                For Each shp In myShapeRange
                    shp.Width = shp.Width + (shp.Left - stretchPosition)
                    shp.Left = stretchPosition
                Next shp

            Case "Right"
                stretchPosition = myShapeRange(myShapeRange.Count).Left + myShapeRange(myShapeRange.Count).Width

                For Each shp In myShapeRange
                    shp.Width = stretchPosition - shp.Left
                Next shp

            Case "Top"
                stretchPosition = myShapeRange(myShapeRange.Count).Top

                For Each shp In myShapeRange
                    shp.Height = shp.Height + stretchPosition - shp.Top
                Next shp

            Case "Bottom"
                stretchPosition = myShapeRange(myShapeRange.Count).Top + myShapeRange(myShapeRange.Count).Height

                For Each shp In myShapeRange
                    shp.Height = stretchPosition - shp.Top
                Next shp

        End Select

    End Sub


End Module
