Imports System.ServiceProcess
Imports ComponenteClasificadorDocumentos

Public Class INTeNT_ClasificadorDocumental
    Inherits System.ServiceProcess.ServiceBase
    Public activoServicio As Boolean = False

#Region " Component Designer generated code "

    Public Sub New()
        MyBase.New()

        ' This call is required by the Component Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call
        If Not objEventLog.SourceExists("INTeNT_ClasificadorDocumental") Then
            objEventLog.CreateEventSource("INTeNT_ClasificadorDocumental", "Application")
        End If
        objEventLog.Source = "INTeNT_ClasificadorDocumental"
        objEventLog.Log = "Application"
    End Sub

    'UserService overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' The main entry point for the process
    <MTAThread()> _
    Shared Sub Main()
        Dim ServicesToRun() As System.ServiceProcess.ServiceBase

        ' More than one NT Service may run within the same process. To add
        ' another service to this process, change the following line to
        ' create a second service object. For example,
        '
        '   ServicesToRun = New System.ServiceProcess.ServiceBase () {New Service1, New MySecondUserService}
        '
        ServicesToRun = New System.ServiceProcess.ServiceBase() {New INTeNT_ClasificadorDocumental}

        System.ServiceProcess.ServiceBase.Run(ServicesToRun)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    ' NOTE: The following procedure is required by the Component Designer
    ' It can be modified using the Component Designer.  
    ' Do not modify it using the code editor.
    Friend WithEvents objEventLog As System.Diagnostics.EventLog
    Friend WithEvents temporizadorServicio As System.Timers.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.objEventLog = New System.Diagnostics.EventLog
        Me.temporizadorServicio = New System.Timers.Timer
        CType(Me.objEventLog, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.temporizadorServicio, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'temporizadorServicio
        '
        Me.temporizadorServicio.AutoReset = False
        Me.temporizadorServicio.Enabled = True
        Me.temporizadorServicio.Interval = 10000
        '
        'INTeNT_ClasificadorDocumental
        '
        Me.ServiceName = "INTeNT_ClasificadorDocumental"
        CType(Me.objEventLog, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.temporizadorServicio, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub

#End Region

    Protected Overrides Sub OnStart(ByVal args() As String)
        Try
            objEventLog.WriteEntry("V. 26/01/2017 Se esta iniciando el servicio de clasificador documental.")
            temporizadorServicio.Enabled = True
            temporizadorServicio.Stop()

            leerParametrosInicialesConfiguracion()

            temporizadorServicio.Interval = tiempoEsperaServicio
            temporizadorServicio.Start()

            objEventLog.WriteEntry("Se inicial el servicio satisfactoriamente con un intervalo de: " & CStr(tiempoEsperaServicio) & " milisegundos.")

        Catch err As Exception
            objEventLog.WriteEntry("Funcion:Start. " & err.Message, EventLogEntryType.Error)
        End Try


    End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
        objEventLog.WriteEntry("Se esta detuvo el servicio de clasificador documental.")
    End Sub

    Private Sub temporizadorServicio_Elapsed(ByVal sender As System.Object, ByVal e As System.Timers.ElapsedEventArgs) Handles temporizadorServicio.Elapsed
        Dim objSingleton As ClassSingleton
        Dim objMonitoreo As New ClassMonitoreoDeCarpeta
        Try
            temporizadorServicio.Stop()
            'ya se inicio correctamente el servicio
            If activoServicio = True Then
                'se crea el objetos con las varaibles globales
                Dim rutaArchivoConfig As String
                rutaArchivoConfig = System.Configuration.ConfigurationSettings.AppSettings.Get("config.xml")
                'objEventLog.WriteEntry("Creando classSingleton ", EventLogEntryType.Information)
                objSingleton = New ClassSingleton(rutaArchivoConfig, rutaCompletaArchivoLog, rutaCompletaArchivoLogErrores, strEcontentServidor, strEcontentlibreria, strEcontentUser, strEcontentPWd, strIdCarpeta, strIdCarpetaError, idCarpetaPadreAlmacenamiento)
                'objEventLog.WriteEntry("Done classSingleton ", EventLogEntryType.Information)

                'objSingleton.objLog.grabarLogEnArchivoPlano("monitorearColaEcontent", "Se comienza a monitorear la carpeta.", ClassGeneradorLog.tipoLog.LogEjecucion)
                'se comienza la ejecucion del monitoreo de carpetas, la cual esta totalmente asignada al compoentne de ejecucion
                objMonitoreo.monitorearCarpetaEcontent(objSingleton)
                'objSingleton.objLog.grabarLogEnArchivoPlano("monitorearColaEcontent", "Proximo evento en: " + CStr(temporizadorServicio.Interval) + "ms", ClassGeneradorLog.tipoLog.LogEjecucion)
                objSingleton.close()
            Else
                activoServicio = True
            End If

            objSingleton = Nothing
            objMonitoreo = Nothing
            temporizadorServicio.Start()
        Catch ex As Exception
            objEventLog.WriteEntry("Error en el evento timer: " & ex.Message, EventLogEntryType.Error)
            Try
                objSingleton.objLog.grabarLogEnArchivoPlano("monitorearColaEcontent", ex.Message, ClassGeneradorLog.tipoLog.LogErrores)
                objSingleton.close()
            Catch ex2 As Exception
                'el error se presento antyes de cargar el archivo de log de errores
            End Try
            objSingleton = Nothing
            objMonitoreo = Nothing
            temporizadorServicio.Start()
        End Try
    End Sub



#Region " Codigo de desarrollo para el SERVICIO de monitoreo de carpetas como tal. "
    'variables del servicio
    Private tiempoEsperaServicio As Double
    Private rutaCompletaArchivoLogErrores As String
    Private rutaCompletaArchivoLog As String ' si la ruta es vacia no se graba log

    'varaibles de econtent
    Private strEcontentServidor As String
    Private strEcontentlibreria As String
    Private strEcontentUser As String
    Private strEcontentPWd As String

    'variables del monitoreo de carpetas
    Private strIdCarpeta As String
    Private strIdCarpetaError As String
    Private idCarpetaPadreAlmacenamiento As String 'carpeta a partir de la cual se comenzara a crear la estructura de almacenamiento para la clase documental

    Private Sub leerParametrosInicialesConfiguracion()
        'en esta funcion se leen las variables globales del servicio, directamete desde el xml de configuracion
        Try
            Dim rutaArchivoConfig As String
            rutaArchivoConfig = System.Configuration.ConfigurationSettings.AppSettings.Get("config.xml")
            'objEventLog.WriteEntry("UPDATED: rutaArchivoConfig" & rutaArchivoConfig)
            Dim objXML As New ClassLecturaXML(rutaArchivoConfig)

            'se capturan los parametros generales del monitoreo de carpetas de econtent
            strEcontentServidor = CStr(objXML.leerValorNodo("MonitorDeCarpetas/Econtent/EcontentServidor", True))
            'objEventLog.WriteEntry("UPDATED: strEcontentServidor" & strEcontentServidor)

            strEcontentlibreria = CStr(objXML.leerValorNodo("MonitorDeCarpetas/Econtent/Econtentlibreria", True))
            'objEventLog.WriteEntry("UPDATED: strEcontentlibreria" & strEcontentlibreria)

            strEcontentUser = CStr(objXML.leerValorNodo("MonitorDeCarpetas/Econtent/EcontentUser", True))
            'objEventLog.WriteEntry("UPDATED: strEcontentUser" & strEcontentUser)

            strEcontentPWd = CStr(objXML.leerValorNodo("MonitorDeCarpetas/Econtent/EcontentPWd", False))
            
            strIdCarpeta = CStr(objXML.leerValorNodo("MonitorDeCarpetas/Econtent/idCarpeta", True))
            'objEventLog.WriteEntry("UPDATED: strIdCarpeta" & strIdCarpeta)

            strIdCarpetaError = CStr(objXML.leerValorNodo("MonitorDeCarpetas/Econtent/idCarpetaError", True))
            'objEventLog.WriteEntry("UPDATED: strIdCarpetaError" & strIdCarpetaError)

            idCarpetaPadreAlmacenamiento = CStr(objXML.leerValorNodo("MonitorDeCarpetas/Econtent/idCarpetaPadreAlmacenamiento", False))
            'objEventLog.WriteEntry("UPDATED: idCarpetaPadreAlmacenamiento" & idCarpetaPadreAlmacenamiento)

            'variable de tiempo de ejecucion
            tiempoEsperaServicio = CDbl(objXML.leerValorNodo("MonitorDeCarpetas/Servicio/Intervalo", True))
            If tiempoEsperaServicio < 15000 Then
                tiempoEsperaServicio = 15000
                'grabar event log
            End If

            rutaCompletaArchivoLogErrores = CStr(objXML.leerValorNodo("MonitorDeCarpetas/Servicio/rutaCompletaArchivoLogErrores", False))
            objEventLog.WriteEntry("UPDATED: rutaCompletaArchivoLogErrores" & rutaCompletaArchivoLogErrores)

            rutaCompletaArchivoLog = CStr(objXML.leerValorNodo("MonitorDeCarpetas/Servicio/rutaCompletaArchivoLog", False))
            objEventLog.WriteEntry("UPDATED: rutaCompletaArchivoLog" & rutaCompletaArchivoLog)

            objXML.close()
            objXML = Nothing

        Catch ex As Exception
            'grabar en event log
            objEventLog.WriteEntry(ex.Message)
            Throw New Exception("leerParametrosInicialesConfiguracion:" & ex.Message)
        End Try
    End Sub

#End Region


End Class
