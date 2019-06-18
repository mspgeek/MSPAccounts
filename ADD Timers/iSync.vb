Public Class clsSync
    Implements LabTech.Interfaces.ISync
    Private m_Host As LabTech.Interfaces.IControlCenter
    Dim g_Name As String

    Public Sub Initialize(ByVal Host As LabTech.Interfaces.IControlCenter) Implements LabTech.Interfaces.ISync.Initialize
        m_Host = Host
    End Sub

    Public Sub Decommision() Implements LabTech.Interfaces.ISync.Decommision
        m_Host = Nothing
    End Sub

    Public ReadOnly Property Name As String Implements LabTech.Interfaces.ISync.Name
        Get
            Return g_Name & "_ISync"
        End Get
    End Property

    Public Sub Syncronize() Implements LabTech.Interfaces.ISync.Syncronize
        '// Perform Operations and Sync the Results with the Database using the Host object.
        Dim changeInterval As Integer = m_Host.GetSQL("SELECT Password_Change_Days FROM `plugin_itsc_msp_accounts_settings`")
        Dim strServiceName As String = m_Host.GetSQL("SELECT Service_Account FROM `plugin_itsc_msp_accounts_settings`")
        Dim ds As DataSet = m_Host.GetDataSet("SELECT Username FROM plugin_itsc_msp_accounts_users WHERE AutoChangePassword IS TRUE AND CURDATE() > DATE_ADD(STR_TO_DATE(AutoChangeDate, '%Y-%m-%d'), INTERVAL " & changeInterval & " DAY)")
        For Each DataRow As DataRow In ds.Tables(0).Rows
            Dim User_Name As String = DataRow("Username").ToString
            m_Host.SetSQL("UPDATE plugin_itsc_msp_accounts_users SET `AutoChangeDate` = CURDATE() WHERE `Username` = '" & User_Name & "'")
            Dim autoPassword As String = "Random"
            PasswordManagement.LoopedChangePassword(m_Host, User_Name, autoPassword)
            If User_Name = strServiceName Then
                If m_Host.GetSQL("Select Local_Service_Account from plugin_itsc_msp_accounts_settings") Then
                    ServiceManagement.LocalServiceLoader(m_Host, "Change")
                End If
            End If
        Next
    End Sub


End Class