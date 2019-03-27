Imports IDMObjects

Public Class ClassSingleton
    Public objLog As ClassGeneradorLog
    Public objXMLConfig As ClassLecturaXML
    'Public objLibreria As New Library
    Public strIdCarpetaEcontent As String
    Public strIdCarpetaErrorEcontent As String
    Public strIdCarpetaPadreAlmacenamiento As String

    Public strSystem As String
    Public strServer As String
    Public strUsuario As String
    Public strPassword As String

    'Private gSessionManager As New IDMObjects.SessionManager

    Public Sub New(ByVal rutaArchivoConfigXML As String, ByVal strRutaArchivoLog As String, ByVal strRutaArchivoLogErrores As String, ByVal strEcontentServer As String, ByVal strEcontentSystem As String, ByVal strEcontentUsuario As String, ByVal strEcontentPSW As String, ByVal strIdCarpeta As String, ByVal strCarpetaError As String, ByVal idCarpetaPadreAlmacenamiento As String)
        Try
            'crear los objetos que viajan por todo el sistema en el esquema singleton
            objLog = New ClassGeneradorLog(strRutaArchivoLog, strRutaArchivoLogErrores)
            objXMLConfig = New ClassLecturaXML(rutaArchivoConfigXML)

            'carpeta de eco0ntent a ser monitoreada
            strIdCarpetaEcontent = strIdCarpeta
            strIdCarpetaErrorEcontent = strCarpetaError
            strIdCarpetaPadreAlmacenamiento = idCarpetaPadreAlmacenamiento

            'se realiza la conexion a eContent
            'objLibreria.SystemType = IDMObjects.idmSysTypeOptions.idmSysTypeDS
            'objLibreria.Name = strEcontentSystem & "^" & strEcontentServer
            'objLibreria.SessionManager = gSessionManager
            'objLibreria.Logon(strEcontentUsuario, strEcontentPSW, "", IDMObjects.idmLibraryLogon.idmLogonOptServerNoUI)

            strSystem = strEcontentSystem
            strServer = strEcontentServer
            strUsuario = strEcontentUsuario
            strPassword = strEcontentPSW

        Catch ex As Exception
            'objLibreria = Nothing
            objXMLConfig = Nothing
            'objLog = Nothing
            Throw New Exception("ClassSingleton.NEW:" & ex.Message)
        End Try
    End Sub

    Public Sub close()
        'cerrar todos los objetos
        Try
            'objLibreria = Nothing
            objLog = Nothing
            objXMLConfig = Nothing
            'objLibreria.Logoff()

        Catch ex As Exception
            'objLibreria = Nothing
            objLog = Nothing
            objXMLConfig = Nothing
        End Try
    End Sub

End Class
