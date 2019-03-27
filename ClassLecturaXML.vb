Imports System
Imports System.IO
Imports System.Xml
Imports System.Linq


Public Class ClassLecturaXML
    Public objXmlDocumento As New XmlDocument

    Public Sub New(ByVal rutaXML As String)
        Try
            objXmlDocumento.Load(rutaXML)
            If objXmlDocumento Is Nothing Then
                'Si el archivo no abre, entonces no se pueden cargar los parametros.
                Throw New Exception("No se pudo leer el archivo de configuración XML, dirección: " & rutaXML)
            End If
        Catch err As Exception
            'se graba el error
            objXmlDocumento = Nothing
            Throw New Exception("NEW:" & err.Message)
        End Try
    End Sub


    Public Function leerValorNodo(ByVal strRutaNodo As String, ByVal esRequerido As Boolean) As String
        'Retorna el valor de un nodo hijo a partir de la ruta entregada.
        'la ruta es de la forma "nodo/nodoNivel2/nodoNivel3", en el ejemplo leeria el valor del xml
        'entrando hasta el nivel 3. EJEMPLO:
        '<nodo><nodoNivel2><nodoNivel3>VALOR</nodoNivel2></nodoNivel2></nodo>
        Try
            'se procede a leer el valor del nodo
            leerValorNodo = objXmlDocumento.SelectSingleNode(strRutaNodo).InnerText

            If esRequerido And Len(leerValorNodo) = 0 Then
                Throw New Exception("El innerText de la variable '" & strRutaNodo & "' no puede ser vacio")
            End If

        Catch err As Exception
            leerValorNodo = ""
            Throw New Exception("leerValorNono: " & err.Message)
        End Try
    End Function

    Public Function retornarXmlChildNodesFromNode(ByVal objXMLNODE As XmlNode, ByVal strNombreNodo As String) As XmlNodeList
        Try
            Return objXMLNODE.SelectNodes(strNombreNodo)
        Catch ex As Exception
            Throw New Exception("retornarXmlChildNodesFromNode: " & ex.Message)
        End Try

    End Function

    Public Function retornarNodo(ByVal strRutaNodo As String) As XmlNode
        'Retorna el un nodo hijo a partir de la ruta entregada.
        'la ruta es de la forma "nodo/nodoNivel2/nodoNivel3", en el ejemplo leeria el valor del xml
        'entrando hasta el nivel 3. EJEMPLO:
        '<nodo><nodoNivel2><nodoNivel3>VALOR</nodoNivel2></nodoNivel2></nodo>

        Dim objXMLNode As XmlNode

        Try
            Dim contador As Integer
            Dim arregloDeNodos() As String
            arregloDeNodos = Split(strRutaNodo, "/")
            objXMLNode = objXmlDocumento.SelectSingleNode(arregloDeNodos(0))
            contador = 1
            While contador <= UBound(arregloDeNodos)
                'se recorren los nodos hasta llegar al nodo que se pide
                objXMLNode = objXMLNode.SelectSingleNode(arregloDeNodos(contador))
                contador = contador + 1
            End While

            retornarNodo = objXMLNode

        Catch err As Exception
            retornarNodo = Nothing
            Throw New Exception("retornarNodo: " & err.Message)
        End Try
    End Function




    Public Sub close()
        objXmlDocumento = Nothing
    End Sub

End Class
