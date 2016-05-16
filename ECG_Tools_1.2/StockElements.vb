Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Core

Module StockElements
    Public myPPT As New Application

    Sub InsertElement(elementName As String)
        Dim fpath As String = "C:\Users\adrak\AppData\Local\Efficient Elements\Efficient Elements for presentations\User\MyElements"
        Dim fname As String = xmlLoader("C:\Users\adrak\AppData\Local\Efficient Elements\Efficient Elements for presentations\User\MyElements.xml", "element", "id", "name", elementName)

        fpath = fpath & "\" & fname & ".pptx"

        Dim currentslide As Slide = myPPT.ActiveWindow.View.Slide

        Dim myPresentation As Presentation = myPPT.Presentations.Open(fpath, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse)

        Dim shp As PowerPoint.Shapes = myPresentation.Slides(1).Shapes

        Dim i As Integer
        For i = 1 To shp.Count
            On Error Resume Next
            If IsError(shp(i).PlaceholderFormat.Type) Then
                shp(i).Copy()
                currentslide.Shapes.Paste()
            End If
        Next i
    End Sub

End Module
