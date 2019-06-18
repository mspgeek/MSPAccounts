Public Class iSync2
    Implements LabTech.Interfaces.ISync2
    Private m_Host As LabTech.Interfaces.IControlCenter
    Dim g_Name As String

    Public Sub Initialize(ByVal Host As LabTech.Interfaces.IControlCenter) Implements LabTech.Interfaces.ISync2.Initialize
        m_Host = Host
    End Sub

    Public Sub Decommision() Implements LabTech.Interfaces.ISync2.Decommision
        m_Host = Nothing
    End Sub

    Public ReadOnly Property Name As String Implements LabTech.Interfaces.ISync2.Name
        Get
            Return g_Name & "_ISync2"
        End Get
    End Property


    Public Sub Syncronize() Implements LabTech.Interfaces.ISync2.Syncronize
        '// run every hour
        If Minute(Now) < 6 Then
            If m_Host.GetSQL("Select Local_Service_Account from plugin_itsc_msp_accounts_settings") Then
                ServiceManagement.LocalServiceLoader(m_Host, "Add")
            End If
        End If
    End Sub


End Class
