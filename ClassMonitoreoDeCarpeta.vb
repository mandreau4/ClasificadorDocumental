Imports IDMObjects
Imports System
Imports System.IO
Imports System.Xml

Public Class ClassMonitoreoDeCarpeta

    Public Sub monitorearCarpetaEcontent(ByRef objSingleton As ClassSingleton)

        Try
            'se procede a conectarse a eContent.
            objSingleton.objLog.grabarLogEnArchivoPlano("monitorearCarpetaEcontent", "Se realizo la conexion a eContent.", ClassGeneradorLog.tipoLog.LogEjecucion)
            objSingleton.objLog.grabarLogEnArchivoPlano("monitorearCarpetaEcontent", "Usuario: " + objSingleton.objLibreria.ActiveUser.ID, ClassGeneradorLog.tipoLog.LogEjecucion)

            Dim objIDM As Object
            Dim objCarpeta As New IDMObjects.Folder
            Dim objDocumento As New IDMObjects.Document
            Dim objArrDocumento As New IDMObjects.ObjectSet
            Dim objxmlLista As XmlNodeList
            Dim objXmlNode As XmlNode
            Dim strIdCarpetaDestino As String

            objIDM = objSingleton.objLibreria.GetObject(idmObjectType.idmObjTypeFolder, objSingleton.strIdCarpetaEcontent)
            objCarpeta = CType(objIDM, IDMObjects.IFnFolderDual)
            'se captura el listado de documentos en la carpeta
            objArrDocumento = objCarpeta.GetContents(idmFolderContent.idmFolderContentDocument)
            objSingleton.objLog.grabarLogEnArchivoPlano("monitorearCarpetaEcontent", "Se realizo la carga de la carpeta a ser monitoreada.", ClassGeneradorLog.tipoLog.LogEjecucion)

            Dim objGeneralEContent As New ClassGeneralEContent

            For Each objIDM In objArrDocumento
                objDocumento = CType(objIDM, IDMObjects.IFnDocumentDual)
                ' disable LocalDB on WS to fix .NET GA issue.
                objDocumento.TrackInLocalDb = False

                Dim documentID As String
                documentID = objDocumento.ID

                objSingleton.objLog.grabarLogEnArchivoPlano("monitorearCarpetaEcontent", "Se procede a clasificar el docuemento '" & documentID & "'.", ClassGeneradorLog.tipoLog.LogEjecucion)
                'se ejecuta la clasificacion del documento en la clase de documentos.
                '*******************************************************************

                strIdCarpetaDestino = traerIdFinalCarpetaAlmacenar(objDocumento, objGeneralEContent, objSingleton)
                If Not objGeneralEContent.moverDocumentoACarpeta(objDocumento, strIdCarpetaDestino, objSingleton.strIdCarpetaEcontent, objSingleton) Then
                    objSingleton.objLog.grabarLogEnArchivoPlano("monitorearCarpetaEcontent", "No se pudo mover el documento al folder destino '" & strIdCarpetaDestino & "', se intentara mover a la carpeta de error.", ClassGeneradorLog.tipoLog.LogErrores)
                    If Not objGeneralEContent.moverDocumentoACarpeta(objDocumento, objSingleton.strIdCarpetaErrorEcontent, objSingleton.strIdCarpetaEcontent, objSingleton) Then
                        objSingleton.objLog.grabarLogEnArchivoPlano("monitorearCarpetaEcontent", "No se pudo mover el documento al folder de error.", ClassGeneradorLog.tipoLog.LogErrores)
                    End If
                End If

                objDocumento = Nothing
                objSingleton.objLog.grabarLogEnArchivoPlano("monitorearCarpetaEcontent", "Se finalizo la clasificacion del documento", ClassGeneradorLog.tipoLog.LogEjecucion)

            Next

            objSingleton.objLog.grabarLogEnArchivoPlano("monitorearCarpetaEcontent", "Se finaliza el monitoreo de la carpeta.", ClassGeneradorLog.tipoLog.LogEjecucion)
            ' Intentar buscar carpeta de errores
            objSingleton.objLog.grabarLogEnArchivoPlano("monitorearCarpetaEcontent", "Monitoreo de la carpeta de errores", ClassGeneradorLog.tipoLog.LogEjecucion)
            objIDM = objSingleton.objLibreria.GetObject(idmObjectType.idmObjTypeFolder, objSingleton.strIdCarpetaErrorEcontent)
            objCarpeta = CType(objIDM, IDMObjects.IFnFolderDual)
            'se captura el listado de documentos en la carpeta
            objArrDocumento = objCarpeta.GetContents(idmFolderContent.idmFolderContentDocument)
            objSingleton.objLog.grabarLogEnArchivoPlano("monitorearCarpetaEcontent", "Documentos en la carpeta de errores, intento de clasificar.", ClassGeneradorLog.tipoLog.LogEjecucion)

            For Each objIDM In objArrDocumento
                objDocumento = CType(objIDM, IDMObjects.IFnDocumentDual)
                ' disable LocalDB on WS to fix .NET GA issue.
                objDocumento.TrackInLocalDb = False

                Dim documentID As String
                documentID = objDocumento.ID

                objSingleton.objLog.grabarLogEnArchivoPlano("monitorearCarpetaEcontent", "Se procede a clasificar el documento '" & documentID & "'.", ClassGeneradorLog.tipoLog.LogEjecucion)
                strIdCarpetaDestino = traerIdFinalCarpetaAlmacenar(objDocumento, objGeneralEContent, objSingleton)
                If Not objGeneralEContent.moverDocumentoACarpeta(objDocumento, strIdCarpetaDestino, objSingleton.strIdCarpetaErrorEcontent, objSingleton) Then
                    objSingleton.objLog.grabarLogEnArchivoPlano("monitorearCarpetaEcontent", "No se pudo mover el documento al folder destino '" & strIdCarpetaDestino & "', se intentara mover a la carpeta de error.", ClassGeneradorLog.tipoLog.LogErrores)
                End If
                objDocumento = Nothing
                objSingleton.objLog.grabarLogEnArchivoPlano("monitorearCarpetaEcontent", "Se finalizo la clasificacion del documento", ClassGeneradorLog.tipoLog.LogEjecucion)
            Next

            objSingleton.objLog.grabarLogEnArchivoPlano("monitorearCarpetaEcontent", "Se finaliza el monitoreo de la carpeta de errores.", ClassGeneradorLog.tipoLog.LogEjecucion)
            'End buscar carpeta de errores

            objSingleton.objLog.grabarLogEnArchivoPlano("monitorearCarpetaEcontent", "Cerrando sesion", ClassGeneradorLog.tipoLog.LogEjecucion)
            objSingleton.objLibreria.Logoff()
            objIDM = Nothing
            objCarpeta = Nothing
            objDocumento = Nothing
            objArrDocumento = Nothing

        Catch ex As Exception
            'grabar en event log
            Try
                objSingleton.objLog.grabarLogEnArchivoPlano("monitorearCarpetaEcontent", ex.Message, ClassGeneradorLog.tipoLog.LogErrores)
            Catch ex2 As Exception
                'el error se presento antyes de cargar el archivo de log de errores
            End Try
            Throw New Exception("monitorearCarpetaEcontent:" & ex.Message)
        End Try
    End Sub



    Public Function traerIdFinalCarpetaAlmacenar(ByVal objDocumento As IDMObjects.Document, ByVal objGeneralEContent As ClassGeneralEContent, ByRef objSingleton As ClassSingleton) As String
        Dim objXmlNode As XmlNode
        Dim objxmlLista As XmlNodeList
        Dim xmlItem As XmlNode
        Dim xmlItemInterno As XmlNode
        Dim objXmlItemsInternos As XmlNodeList
        Dim objXmlListSeguridad As XmlNodeList
        Dim objItemSeguridad As XmlNode

        Dim strNombrePropiedad As String
        Dim strValorPropiedad As String
        Dim strValorEsperadoPropiedad As String
        Dim strCarpetaPadre As String
        Dim carpetaMes As String

        Dim strNombreGrupoSeguridad As String
        Dim strNivelAcceso As String
        Dim intNivelAcceso As Integer

        Dim objPermisosGrupo As IDMObjects.Permission

        'cadena de caracteres que tiene el valor "TRUE" cuando la carpeta debe ser creada con grupos de seguridad y se le debe asignar al documento dicha seguridad
        Dim tieneSeguridad As String
        'cadena de caracteres separada por coma "," que indica que id de carpetas deberan ser asignados como seguridad al documento.
        Dim strVectorIdCarpetasConSeguridad As String
        Dim vectorIdCarpetas
        Dim contador As Integer

        Dim valorPropiedadFecha As String
        Dim nombrePropiedadFecha As String
        Dim dateFecha As Date


        Try


            objXmlNode = objSingleton.objXMLConfig.retornarNodo("MonitorDeCarpetas/estructuraAlamacenamientoEcontent")
            objxmlLista = objSingleton.objXMLConfig.retornarXmlChildNodesFromNode(objXmlNode, "documentClass")

            objSingleton.objLog.grabarLogEnArchivoPlano("traerIdFinalCarpetaAlmacenar", "Se procede a organizar en econtent la estructura para el archivo.", ClassGeneradorLog.tipoLog.LogEjecucion)
            'se lee la carpeta padre (nivel 1) en la cual se va a estructurar el archiovo de econtent en procesamiento.
            strValorPropiedad = objGeneralEContent.leerParametroDeDocumento(objDocumento, "idmDocType", objSingleton)
            'strValorPropiedad = objGeneralEContent.eliminarTildes(strValorPropiedad)

            'se asigna la carpeta padre GENERAL de almacenamiento definida en el XML
            strCarpetaPadre = objSingleton.strIdCarpetaPadreAlmacenamiento


            For Each xmlItem In objxmlLista
                strValorEsperadoPropiedad = xmlItem.Attributes("valor").Value
                'se compara sin tildes pero se generan carpetas con tildes
                If Trim(objGeneralEContent.eliminarTildes(UCase(strValorPropiedad))) = Trim(objGeneralEContent.eliminarTildes(UCase(strValorEsperadoPropiedad))) Then
                    'Como son iguales es porque se localizo la clase documental a ser movido el documento.

                    'Se procede a leer el identificador de la propiedad de la fecha, el cual indica el nombre final de las carpetas por fecha.
                    objSingleton.objLog.grabarLogEnArchivoPlano("traerIdFinalCarpetaAlmacenar", "Se procede a capturar la propiedad fecha (nombrePropiedadFecha) en el xml para la clase documental encontrada.", ClassGeneradorLog.tipoLog.LogEjecucion)
                    nombrePropiedadFecha = xmlItem.Attributes("idmDocCustomFecha").Value
                    valorPropiedadFecha = objGeneralEContent.leerParametroDeDocumento(objDocumento, nombrePropiedadFecha, objSingleton)
                    If Not IsDate(valorPropiedadFecha) Then
                        Throw New Exception("La fecha capturada de eContent para la clase documental no es valida. Fecha capturada: '" & valorPropiedadFecha & "', propiedad econtent: '" & nombrePropiedadFecha & "'")
                    End If

                    'se procede a crear la carpeta del primer nivel
                    objSingleton.objLog.grabarLogEnArchivoPlano("traerIdFinalCarpetaAlmacenar", "Se procede crear la carpeta '" & strValorPropiedad & "'.", ClassGeneradorLog.tipoLog.LogEjecucion)
                    strCarpetaPadre = objGeneralEContent.crearCarpetaSiNoExiste(Trim(strValorPropiedad), strCarpetaPadre, objSingleton)

                    'Siempre la primera carpeta tiene seguridad que debera ser asignada al documento creado.
                    strVectorIdCarpetasConSeguridad = strCarpetaPadre

                    'se procede a navegar en la estructura del nivel
                    objXmlItemsInternos = objSingleton.objXMLConfig.retornarXmlChildNodesFromNode(xmlItem, "propiedad")
                    For Each xmlItemInterno In objXmlItemsInternos
                        strValorEsperadoPropiedad = xmlItemInterno.Attributes("idmDocCustom").Value
                        strValorPropiedad = Trim(objGeneralEContent.leerParametroDeDocumento(objDocumento, strValorEsperadoPropiedad, objSingleton))
                        'se procede a crear la carpeta para el nivel interno
                        objSingleton.objLog.grabarLogEnArchivoPlano("traerIdFinalCarpetaAlmacenar", "Se procede crear la carpeta '" & strValorPropiedad & "'.", ClassGeneradorLog.tipoLog.LogEjecucion)
                        strCarpetaPadre = objGeneralEContent.crearCarpetaSiNoExiste(Trim(strValorPropiedad), strCarpetaPadre, objSingleton)

                        'se verifica si la carpeta tiene seguridad a ser asignada
                        tieneSeguridad = xmlItemInterno.Attributes("tieneSeguridad").Value
                        If UCase(tieneSeguridad) = "TRUE" Then
                            objSingleton.objLog.grabarLogEnArchivoPlano("traerIdFinalCarpetaAlmacenar", "Se detecta que la carpeta tiene seguridad y se procede a crear los grupos.", ClassGeneradorLog.tipoLog.LogEjecucion)
                            'hay que crearle los grupos de seguridad y adicionar la carpeta al vector de carpetas con seguridad
                            'para asi ser asignada al final al documento.
                            strVectorIdCarpetasConSeguridad = strVectorIdCarpetasConSeguridad & "," & strCarpetaPadre

                            'crear grupos de seguridad
                            objXmlListSeguridad = objSingleton.objXMLConfig.retornarXmlChildNodesFromNode(xmlItemInterno, "seguridad")
                            For Each objItemSeguridad In objXmlListSeguridad
                                strNombreGrupoSeguridad = objItemSeguridad.Attributes("nombreGrupo").Value
                                strNombreGrupoSeguridad = strNombreGrupoSeguridad & strValorPropiedad
                                strNivelAcceso = objItemSeguridad.Attributes("acceso").Value
                                If Not IsNumeric(strNivelAcceso) Then
                                    objSingleton.objLog.grabarLogEnArchivoPlano("traerIdFinalCarpetaAlmacenar", "El nivel de acceso capturado del xml para el grupo no es valido, debe ser numerico, se asignara viewer por defecto.", ClassGeneradorLog.tipoLog.LogErrores)
                                    intNivelAcceso = 1
                                Else
                                    intNivelAcceso = CInt(strNivelAcceso)
                                End If
                                objSingleton.objLog.grabarLogEnArchivoPlano("traerIdFinalCarpetaAlmacenar", "Se va a crear el grupo '" & strNombreGrupoSeguridad & "' con un nivel de seguridad " & strNivelAcceso & " (de ser necesario) para la carpeta.", ClassGeneradorLog.tipoLog.LogEjecucion)
                                objGeneralEContent.asignarNuevosGruposACarpeta(strCarpetaPadre, strNombreGrupoSeguridad, intNivelAcceso, objSingleton)
                            Next

                            objSingleton.objLog.grabarLogEnArchivoPlano("traerIdFinalCarpetaAlmacenar", "Se finaliza la creacion o ingreso a la carpeta.", ClassGeneradorLog.tipoLog.LogEjecucion)
                        End If
                    Next
                    'ya se entro hasta el ultimo nivel, entonces se procede a crear la estructura interna de almacena miento (por fechas)
                    'AAAA/MM
                    objSingleton.objLog.grabarLogEnArchivoPlano("traerIdFinalCarpetaAlmacenar", "Se van a crear las carpetas para la fecha capturada de la propiedad de eContent.", ClassGeneradorLog.tipoLog.LogEjecucion)
                    dateFecha = CDate(valorPropiedadFecha)
                    strCarpetaPadre = objGeneralEContent.crearCarpetaSiNoExiste(CStr(Year(dateFecha)), strCarpetaPadre, objSingleton)

                    If CInt(Month(dateFecha)) < 10 Then
                        carpetaMes = "0" & CStr(Month(dateFecha))
                    Else
                        carpetaMes = CStr(Month(dateFecha))
                    End If
                    strCarpetaPadre = objGeneralEContent.crearCarpetaSiNoExiste(carpetaMes, strCarpetaPadre, objSingleton)

                    vectorIdCarpetas = Split(strVectorIdCarpetasConSeguridad, ",")

                    contador = 0
                    While contador <= UBound(vectorIdCarpetas)
                        'se asigna una a una la seguridad de las carpetas quen se quiere
                        objSingleton.objLog.grabarLogEnArchivoPlano("traerIdFinalCarpetaAlmacenar", "Se va a asignar la seguridad de la carpeta '" & vectorIdCarpetas(contador) & "'", ClassGeneradorLog.tipoLog.LogEjecucion)
                        objGeneralEContent.asignarSeguridadADocumento(objDocumento, vectorIdCarpetas(contador), objSingleton)
                        contador = contador + 1
                    End While

                    'se retorna el id de la carpeta destino donde debera quedar alamcenado el documento.
                    traerIdFinalCarpetaAlmacenar = strCarpetaPadre

                    'ya entro a la estructura que le sirve de ayuda, ahora ya puede salir de la funcion para mover el documento.
                    Exit Function

                End If
            Next

            Throw New Exception("No se encontro ninguna estructura de almacenamiento definida para la clase documental del documento '" & objDocumento.ID & "'.")
        Catch ex As Exception
            'ojo no lanzar una nueva "exeption" en esta parte porque no enviaria el documento a una carpeta de error.
            objSingleton.objLog.grabarLogEnArchivoPlano("TraerIdFinalCarpetaAlmacenar:", ex.Message, ClassGeneradorLog.tipoLog.LogErrores)
        End Try
    End Function
End Class
