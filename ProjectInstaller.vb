Imports System.ComponentModel
Imports System.Configuration.Install

<RunInstaller(True)> Public Class ProjectInstaller
    Inherits System.Configuration.Install.Installer

#Region " Component Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Installer overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents ServiceInstaller1 As System.ServiceProcess.ServiceInstaller
    Friend WithEvents INTeNT_ClasificadorDocumental As System.ServiceProcess.ServiceInstaller
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.INTeNT_ClasificadorDocumental = New System.ServiceProcess.ServiceInstaller()
        Me.ServiceInstaller1 = New System.ServiceProcess.ServiceInstaller()
        '
        'INTeNT_ClasificadorDocumental
        '
        Me.INTeNT_ClasificadorDocumental.DisplayName = "INTeNT_ClasificadorDocumental"
        Me.INTeNT_ClasificadorDocumental.ServiceName = "INTeNT_ClasificadorDocumentalV1"
        Me.INTeNT_ClasificadorDocumental.StartType = System.ServiceProcess.ServiceStartMode.Automatic
        '
        'ServiceInstaller1
        '
        Me.ServiceInstaller1.Description = "Clasificador documental service"
        Me.ServiceInstaller1.ServiceName = "INTeNT_ClasificadorDocumentalv1"
        '
        'ProjectInstaller
        '
        Me.Installers.AddRange(New System.Configuration.Install.Installer() {Me.INTeNT_ClasificadorDocumental, Me.ServiceInstaller1})

    End Sub

#End Region

End Class
