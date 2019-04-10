Imports IDMObjects
Imports AdminAPI
Imports AdminAPI.SPI_STS_CODES

Public Class ClassGeneralEContent

    Public Function moverDocumentoACarpeta(ByRef objDocument As IDMObjects.Document, ByVal idCarpetaDestino As String, ByVal idCarpetaOrigen As String, ByRef objSinlgeton As ClassSingleton) As Boolean
        'se crean los objetos
        Dim aIdmFolderDestino As IDMObjects.Folder
        Dim aIdmFolderOrigen As IDMObjects.Folder
        Dim aFolderDestino As Object
        Dim FolderAlmacenado As Object
        Dim aFolderOrigen As Object
        Try
            'Instancia del folder Destino
            aFolderDestino = objSinlgeton.objLibreria.GetObject(idmObjectType.idmObjTypeFolder, idCarpetaDestino)
            aIdmFolderDestino = CType(aFolderDestino, IDMObjects.IFnFolderDual)
            aIdmFolderDestino.File(objDocument)

            For Each FolderAlmacenado In objDocument.FoldersFiledIn
                If (FolderAlmacenado.id = idCarpetaOrigen) Then
                    'Instancia del folder Origen
                    aFolderOrigen = objSinlgeton.objLibreria.GetObject(idmObjectType.idmObjTypeFolder, idCarpetaOrigen)
                    aIdmFolderOrigen = CType(aFolderOrigen, IDMObjects.IFnFolderDual)
                    aIdmFolderOrigen.Unfile(objDocument)
                End If
            Next

            aIdmFolderDestino = Nothing
            aIdmFolderOrigen = Nothing
            aFolderDestino = Nothing
            FolderAlmacenado = Nothing
            aFolderOrigen = Nothing
            moverDocumentoACarpeta = True
        Catch ex As Exception
            objSinlgeton.objLog.grabarLogEnArchivoPlano("moverDocumentoACarpeta", ex.Message, ClassGeneradorLog.tipoLog.LogErrores)
            aIdmFolderDestino = Nothing
            aIdmFolderOrigen = Nothing
            aFolderDestino = Nothing
            FolderAlmacenado = Nothing
            aFolderOrigen = Nothing
            moverDocumentoACarpeta = False
            'OJO: no generar un nuevo log de errores para no interferir con el proceso de mover documento a cola de error
            'Throw New Exception("MoverDocumentoACarpeta:" & ex.Message)
        End Try
    End Function

    Public Function crearCarpetaSiNoExiste(ByVal strNombreCarpeta As String, ByVal strIdCarpetaPadre As String, ByVal objSingleton As ClassSingleton) As String
        'Esta funcion verifica si exite una carpeta con el nombre entregado en la propiedad strNombreCarpeta, dentro de la carpeta idCarpetaPAdre,
        'si la carpeta no existe la crea dentro de la carpeta padre. Retorno de la funcion es el id del folder creado o capturado.
        Dim aIdmFolderPadre As IDMObjects.Folder
        Dim objFolder As IDMObjects.Folder
        Dim strIdFolderFinal As String 'id de la carpeta capturada o creada "strNombreCarpeta".
        Try
            If Not existeCarpeta(Trim(strNombreCarpeta), strIdCarpetaPadre, strIdFolderFinal, objSingleton) Then
                If Len(strIdCarpetaPadre) = 0 Then
                    'se almacena en la raiz
                    objFolder = objSingleton.objLibreria.CreateObject(idmObjectType.idmObjTypeFolder, "")
                Else
                    aIdmFolderPadre = objSingleton.objLibreria.GetObject(idmObjectType.idmObjTypeFolder, strIdCarpetaPadre)
                    aIdmFolderPadre = CType(aIdmFolderPadre, IDMObjects.IFnFolderDual)
                    objFolder = aIdmFolderPadre.CreateSubFolder(idmObjectType.idmObjTypeFolder, "")
                End If
                ' Set the name for the new folder
                objFolder.Name = Trim(strNombreCarpeta)
                ' Save to the library.  
                objFolder.SaveNew()
                'se captura el id del folder
                strIdFolderFinal = objFolder.ID
                objFolder = Nothing
            End If
            'se retorna el id del folder capturado (o creado si fue necesario)
            Return strIdFolderFinal
            aIdmFolderPadre = Nothing
            objFolder = Nothing
        Catch ex As Exception
            objSingleton.objLog.grabarLogEnArchivoPlano("crearCarpetaSiNoExiste", ex.Message, ClassGeneradorLog.tipoLog.LogErrores)
            aIdmFolderPadre = Nothing
            objFolder = Nothing
            Throw New Exception("crearCarpetaSiNoExiste:" & ex.Message)
        End Try
    End Function



    Public Function existeCarpeta(ByVal strNombreCarpeta As String, ByVal strIdFolderPadre As String, ByRef strOutIDCarpeta As String, ByVal objSingleton As ClassSingleton) As Boolean
        'Esta funcion verifica si un folder existe dentro de la carpeta padre.
        'PArametros de entrada:
        'strNombreCarpeta: Nombre de la carpeta a buscar si existe o no
        'strIdFolderPadre: Id de la carpeta en donde se va a buscar el folder. Si este id esta vacio, entonces se busca en la raiz (objLibrary)
        Try
            Dim objFolderPadre As IDMObjects.Folder
            Dim objFoldersColl As IDMObjects.ObjectSet
            Dim objSubFolder As IDMObjects.Folder

            If Len(strIdFolderPadre) = 0 Then
                'como no se paso el id Del padre entonces se busca en la raiz de la libreria
                objFoldersColl = objSingleton.objLibreria.TopFolders()

            Else
                'se procede a cargar la carpeta padre
                objFolderPadre = objSingleton.objLibreria.GetObject(idmObjectType.idmObjTypeFolder, strIdFolderPadre)
                objFolderPadre = CType(objFolderPadre, IDMObjects.IFnFolderDual)
                objFoldersColl = objFolderPadre.SubFolders
            End If

            'SE RECORRE LA LISTA DE CARPETAS ENCONTRADAS EN LA RAIZ O EN EL FOLDER PADRE
            If objFoldersColl.Count > 0 Then
                For Each objSubFolder In objFoldersColl
                    If objSubFolder.Label = Trim(strNombreCarpeta) Then
                        'si el folder encontrado es el mismo que el que se esta buscando entonces se sale de la funcion.
                        strOutIDCarpeta = objSubFolder.ID
                        objSubFolder = Nothing
                        objFoldersColl = Nothing
                        objFolderPadre = Nothing
                        Return True
                        Exit Function
                    End If
                Next
            End If

            'si se llega hasta este punto entonces no existe el folder
            Return False
        Catch ex As Exception
            objSingleton.objLog.grabarLogEnArchivoPlano("existeCarpeta", ex.Message, ClassGeneradorLog.tipoLog.LogErrores)
            Throw New Exception("existeCarpeta:" & ex.Message)
        End Try
    End Function

    Public Function eliminarTildes(ByVal strCadena As String) As String
        Dim temporal As String
        temporal = Replace(strCadena, "á", "a")
        temporal = Replace(temporal, "é", "e")
        temporal = Replace(temporal, "í", "i")
        temporal = Replace(temporal, "ó", "o")
        temporal = Replace(temporal, "ú", "u")
        temporal = Replace(temporal, "Á", "A")
        temporal = Replace(temporal, "É", "E")
        temporal = Replace(temporal, "Í", "I")
        temporal = Replace(temporal, "Ó", "O")
        temporal = Replace(temporal, "Ú", "U")
        eliminarTildes = temporal
    End Function

    Public Function leerParametroDeDocumento(ByVal objDocumento As IDMObjects.Document, ByVal strParametro As String, ByVal objSingleton As ClassSingleton) As String
        Try
            Dim objPropiedad As IDMObjects.Property
            Dim valorPropiedad As String
            objPropiedad = objDocumento.Properties.Item(strParametro)
            valorPropiedad = CStr(objPropiedad.Value)
            leerParametroDeDocumento = valorPropiedad
        Catch ex As Exception
            objSingleton.objLog.grabarLogEnArchivoPlano("leerParametroDeDocumento", "Propiedad:'" & strParametro & "', error: " & ex.Message, ClassGeneradorLog.tipoLog.LogErrores)
            Throw New Exception("leerParametroDeDocumento:" & ex.Message)
        End Try
    End Function

    Public Function asignarSeguridadADocumento(ByRef objDocumento As IDMObjects.Document, ByVal strIdCarpeta As String, ByRef objSingleton As ClassSingleton)
        Dim aIdmFolder As IDMObjects.Folder
        Dim idmPermisos As IDMObjects.Permissions
        Dim permiso As Permission

        Try
            aIdmFolder = objSingleton.objLibreria.GetObject(idmObjectType.idmObjTypeFolder, strIdCarpeta)
            aIdmFolder = CType(aIdmFolder, IDMObjects.IFnFolderDual)

            idmPermisos = aIdmFolder.Permissions

            For Each permiso In idmPermisos
                objDocumento.Permissions.AddByName(permiso.GranteeName, permiso.GranteeType, permiso.Access)
            Next
            objDocumento.Save()

            aIdmFolder = Nothing
            idmPermisos = Nothing
            permiso = Nothing

        Catch ex As Exception
            objSingleton.objLog.grabarLogEnArchivoPlano("asignarSeguridadADocumento", ex.Message, ClassGeneradorLog.tipoLog.LogErrores)
            aIdmFolder = Nothing
            idmPermisos = Nothing
            permiso = Nothing
            Throw New Exception("asignarSeguridadADocumento:" & ex.Message)
        End Try
    End Function


    Public Function asignarNuevosGruposACarpeta(ByVal strIdCarpeta As String, ByVal strNombreGrupo As String, ByVal nivelAcceso As Integer, ByRef objSingleton As ClassSingleton)
        Dim aIdmFolder As IDMObjects.Folder
        Try
            m_CreateGroup(strNombreGrupo, objSingleton)
            aIdmFolder = objSingleton.objLibreria.GetObject(idmObjectType.idmObjTypeFolder, strIdCarpeta)
            aIdmFolder = CType(aIdmFolder, IDMObjects.IFnFolderDual)

            aIdmFolder.Permissions.AddByName(strNombreGrupo, idmObjectType.idmObjTypeGroup, nivelAcceso)
            aIdmFolder.Save()

            aIdmFolder = Nothing
        Catch ex As Exception
            objSingleton.objLog.grabarLogEnArchivoPlano("asignarNuevosGruposACarpeta", ex.Message, ClassGeneradorLog.tipoLog.LogErrores)
            aIdmFolder = Nothing
            Throw New Exception("asignarNuevosGruposACarpeta:" & ex.Message)
        End Try
    End Function



#Region "Declaracion de constantes"
    Private Const SPI_OBJ_SYSTEM As Integer = 1280
    Private Const SPI_SYS_GROUPS As Integer = 1302
    Private Const SPI_OBJ_GROUP As Integer = 1536
    Private Const SPI_GROUP_NAME As Integer = 1537
#End Region



    Protected Function m_Logon(ByVal objSingleton As ClassSingleton) As AdminAPI.mezSession
        Try
            Dim iStatus As Integer
            Dim m_mzoSession As AdminAPI.mezSession
            m_mzoSession = New AdminAPI.mezSession
            iStatus = m_mzoSession.Login(objSingleton.strUsuario, objSingleton.strPassword, objSingleton.strSystem, objSingleton.strServer)
            If iStatus = AdminAPI.SPI_STS_CODES.SPI_STS_SUCCESS Then
                Return m_mzoSession
            Else
                Return Nothing
                Throw New Exception("Se presento un error realizando el login a la api. iStatus: " & iStatus)
            End If
        Catch ex As Exception
            Throw New Exception("m_logon: " & ex.Message)
        End Try
    End Function


    Protected Function m_Logoff(ByVal objSession As mezSession)
        Try
            If Not objSession Is Nothing Then
                objSession.Logout()
                objSession = Nothing
            End If
        Catch ex As Exception
        End Try
    End Function


    Protected Function m_CreateGroup(ByVal strNombreGrupo As String, ByVal objSingleton As ClassSingleton)
        'Declaraciones para la obtencion de coleccion de grupos 
        'Y creacion de nuevo elemento de grupo
        Dim mzoProperty As mezProperty
        Dim mzoNewGroup As mezObject
        Dim iStatus As Integer
        Dim mzoUtility As mezUtility

        'Declaraciones para manipulacion de propiedades del grupo como tal
        Dim iIndex As Integer
        Dim mzoPropertyGroup As mezProperty

        Dim sErrInfo As String
        Dim iNumMsgs As Integer

        Dim objSession As mezSession

        Try

            objSingleton.objLog.grabarLogEnArchivoPlano("m_CreateGroup", "se procede a crear acceso api del csexplorer", ClassGeneradorLog.tipoLog.LogEjecucion)

            objSession = m_Logon(objSingleton)


            If strNombreGrupo = "" Then
                Throw New Exception("Un nombre para el grupo debe ser indicado")
            End If

            mzoUtility = New mezUtility

            'Get de la coleccion de grupos.

            objSingleton.objLog.grabarLogEnArchivoPlano("m_CreateGroup", "se continua con la creacion del grupo '" & strNombreGrupo & "' (si no existe)", ClassGeneradorLog.tipoLog.LogEjecucion)

            iStatus = mzoUtility.GetMezProperty(objSession.hSctx, SPI_OBJ_SYSTEM, SPI_SYS_GROUPS, objSession.sSystemName, mzoProperty)
            If iStatus <> SPI_STS_SUCCESS Then
                Throw New Exception("Error tratando de obtener el objeto de grupos: (" & CStr(iStatus) & ")")
            Else
                'Crea un nuevo elemento en la coleccion de grupos.
                iStatus = mzoProperty.NewChild(mzoNewGroup)
                If iStatus <> SPI_STS_SUCCESS Then
                    Throw New Exception("Error creando un nuevo elemento de grupo: (" & CStr(iStatus) & ")")
                Else
                    'Aqui iteremos todas las propiedades del grupo, buscando el nombre de grupo
                    'para signarlo
                    For iIndex = 1 To mzoNewGroup.iPropertyCount
                        'Get de la propiedad especifica del grupo.
                        iStatus = mzoNewGroup.GetPropertyByIndex(iIndex, mzoPropertyGroup)
                        If iStatus <> SPI_STS_SUCCESS Then
                            Throw New Exception("Error obteniendo la propiedad: (" & CStr(iStatus) & ")")
                        End If
                        'Aqui se modifica el nombre del grupo    
                        If mzoPropertyGroup.sPropertyName = "Group Name" Then
                            mzoPropertyGroup.SetPropertyData(Trim(strNombreGrupo))
                        End If
                    Next

                    'Por ultimo se actualiza el grupo con el comando update. esta isntrucciones es 
                    'la que en realidad crea el grupo(siempre que este tenga asignado al menos un nombre
                    iStatus = mzoNewGroup.Update(sErrInfo, iNumMsgs)
                    If iStatus <> SPI_STS_SUCCESS Then
                        'este error no se reporta porque probablemente es que el grupo ya existe.
                        'Throw New Exception("El grupo no pudo ser creado, se presento el siguiente error: codigo" & CStr(iStatus))
                        objSingleton.objLog.grabarLogEnArchivoPlano("m_CreateGroup", "se presento un error creando el grupo, puede ser que el grupo ya existe o que se presento un error en la escritura. iStatus " & iStatus, ClassGeneradorLog.tipoLog.LogEjecucion)
                    Else
                        objSingleton.objLog.grabarLogEnArchivoPlano("m_CreateGroup", "Se creo el grupo satisfactoriamente.", ClassGeneradorLog.tipoLog.LogEjecucion)
                        'Return "El grupo " & strNombreGrupo & " fue creado exitosamente"
                    End If
                End If
            End If

            m_Logoff(objSession)

        Catch excep As Exception
            Try
                m_Logoff(objSession)
            Catch ex As Exception
            End Try
            objSingleton.objLog.grabarLogEnArchivoPlano("m_CreateGroup", excep.Message, ClassGeneradorLog.tipoLog.LogErrores)
        End Try
    End Function

End Class

