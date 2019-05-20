<System.ComponentModel.RunInstaller(True)> Partial Class InstaladorDelServicioDeCarga
    Inherits System.Configuration.Install.Installer

    'Installer reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de componentes
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de componentes requiere el siguiente procedimiento
    'Se puede modificar usando el Diseñador de componentes.
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.InstaladorDelServicio = New System.ServiceProcess.ServiceProcessInstaller()
        Me.ServicioDeCargaDelReloj = New System.ServiceProcess.ServiceInstaller()
        '
        'InstaladorDelServicio
        '
        Me.InstaladorDelServicio.Account = System.ServiceProcess.ServiceAccount.LocalSystem
        Me.InstaladorDelServicio.Password = Nothing
        Me.InstaladorDelServicio.Username = Nothing
        '
        'ServicioDeCargaDelReloj
        '
        Me.ServicioDeCargaDelReloj.Description = "Es el servicio encargado de mantener la base de datos en el servidor actualizada"
        Me.ServicioDeCargaDelReloj.DisplayName = "Servicio de carga de los relojes a la base de datos"
        Me.ServicioDeCargaDelReloj.ServiceName = "Servicio de carga de los relojes checadores"
        Me.ServicioDeCargaDelReloj.StartType = System.ServiceProcess.ServiceStartMode.Automatic
        '
        'InstaladorDelServicioDeCarga
        '
        Me.Installers.AddRange(New System.Configuration.Install.Installer() {Me.InstaladorDelServicio, Me.ServicioDeCargaDelReloj})

    End Sub

    Friend WithEvents InstaladorDelServicio As ServiceProcess.ServiceProcessInstaller
    Friend WithEvents ServicioDeCargaDelReloj As ServiceProcess.ServiceInstaller
End Class
