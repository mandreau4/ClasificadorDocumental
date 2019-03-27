Imports System
Imports System.IO

Public Class ClassGeneradorLog
    Private rutaArchivoLog As String
    Private rutaArchivoLogErrores As String

    Public Enum tipoLog
        LogErrores = 1
        LogEjecucion = 2
    End Enum

    Public Sub New(ByVal strRutaArchivoLog As String, ByVal strRutaArchivoLogErrores As String)
        rutaArchivoLog = strRutaArchivoLog
        rutaArchivoLogErrores = strRutaArchivoLogErrores
    End Sub

    Public Sub grabarLogEnArchivoPlano(ByVal strFuncion As String, ByVal strDescripcion As String, ByVal tipoLog As tipoLog)
        Try
            If tipoLog = tipoLog.LogErrores Then
                'se graban el log de errores.
                grabarLineaArchivoPlano("Funcion: " & strFuncion & ". Error: " & strDescripcion, rutaArchivoLogErrores)
                'sepre se intenta dejar rastro en el log de ejecucion.
                'grabarLineaArchivoPlano("Funcion: " & strFuncion & ". Error: " & strDescripcion, rutaArchivoLog)

            ElseIf tipoLog = tipoLog.LogEjecucion Then
                'grabar log de ejcucion, siempre y cuando la ruta sea valida.
                grabarLineaArchivoPlano("Funcion: " & strFuncion & ". Log: " & strDescripcion, rutaArchivoLog)
            End If
        Catch ex As Exception
            Throw New Exception("grabarLogEnArchivoPlano:" & ex.Message)
        End Try
    End Sub

    Private Sub grabarLineaArchivoPlano(ByVal strMensaje As String, ByVal rutaArchivo As String)
        Try
            If Len(rutaArchivo) <> 0 Then
                Dim file As New System.IO.StreamWriter(rutaArchivo & "_" & Year(Now) & Month(Now) & Day(Now) & ".log", True)
                file.WriteLine(CStr(Now) & ":" & strMensaje)
                file.Close()
                file = Nothing
            End If
        Catch ex As Exception
            'Throw New Exception("grabarLineaArchivoPlano:" & ex.Message)
        End Try
    End Sub

    Public Function close()
        rutaArchivoLog = Nothing
        rutaArchivoLogErrores = Nothing
    End Function

End Class
