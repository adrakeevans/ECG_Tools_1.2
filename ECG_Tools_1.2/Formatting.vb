
Imports System.Xml



Module Formatting


    Function xmlLoader(ByVal fPath As String, ByVal nodeName As String, ByVal returnAttributeName As String, ByVal Optional conditionalAttributeName As String = Nothing, ByVal Optional conditionalAttribute As String = Nothing) As String

        ' Create an XmlReader
        Dim reader As XmlReader = XmlReader.Create(fPath)
        Dim output As String
        While reader.Read
            If reader.Name = nodeName And reader.HasAttributes Then
                reader.MoveToAttribute(conditionalAttributeName)

                If reader.Value = conditionalAttribute Then
                    reader.MoveToAttribute(returnAttributeName)
                    output = reader.Value
                End If
            End If

        End While

        Return output

    End Function


End Module
