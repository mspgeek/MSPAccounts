Public Class Reporting

    Public Shared Sub log_Reporting(m_host As LabTech.Interfaces.IControlCenter, passedUserName As String)
        On Error GoTo errorHandler
        Dim sqlstr_Domains As String = "SELECT CONCAT(Computers.ComputerID, '|', TRIM(LEADING 'DC:' FROM computers.Domain)) FROM computers WHERE computers.Computerid IN (SELECT Computers.ComputerID FROM Computers, v_detectedroles WHERE v_detectedroles.ComputerID = Computers.ComputerID AND v_detectedroles.RoleName = 'AD PDC Emulator')"
        Dim arrComputerDomains = m_host.GetValues(sqlstr_Domains)
        Dim dicComputerDomains As New Dictionary(Of String, String)
        For Each strDomain As String In arrComputerDomains
            Dim mytmpComputerID As String = strDomain.ToString.Split("|")(0)
            Dim mytmpDomainName As String = strDomain.ToString.Split("|")(1)
            dicComputerDomains.Add(mytmpComputerID, mytmpDomainName)
        Next
        Dim strConcatDCids As String = m_host.GetSQL("SELECT ClientDCids FROM plugin_itsc_msp_accounts_userstatus WHERE Username = '" & passedUserName & "'")
        Dim arrDCids() As String = strConcatDCids.Split(New Char() {"|"c})
        Dim myLogResult As String = "Change log for " & passedUserName & ".<br/><br/>"
        Dim strEachDC As String = ""
        For Each strEachDC In arrDCids
            Dim tmpColumnName As String = dicComputerDomains.Item(strEachDC)
            Dim tmpResult As String = m_host.GetSQL("SELECT `" & tmpColumnName & "` FROM plugin_itsc_msp_accounts_userstatus WHERE `Username` = '" & passedUserName & "'")
            Dim tmpDate As DateTime = Now
            If IsDate(tmpResult.ToString.Split("|")(0)) Then
                tmpDate = CDate(tmpResult.ToString.Split("|")(0))
            End If
            Dim tmpMessage As String = ""
            If String.IsNullOrEmpty(tmpResult) Or tmpResult.Contains("-9999") Then
                myLogResult &= tmpColumnName & ": The requested action failed to log anything for this client.<br/>" & vbCrLf
            Else
                tmpMessage = tmpResult.ToString.Split("|")(1)
                If Now < tmpDate Then
                    myLogResult &= tmpColumnName & ": The requested action failed to log anything for this client. (debug|" & tmpDate & "|" & strEachDC & ")<br/>" & vbCrLf
                ElseIf (Not tmpResult.Contains("succeeded")) Then
                    Dim tmpRecordText As String = tmpResult.ToString.Split("|")(1)
                    myLogResult &= tmpColumnName & ": " & tmpRecordText & "<br/>" & vbCrLf
                End If
            End If
            tmpMessage = tmpDate.ToString("yyyy-MM-dd HH:mm:ss") & "|" & tmpMessage
            m_host.SetSQL("UPDATE plugin_itsc_msp_accounts_userstatus SET `" & tmpColumnName & "` = '" & tmpMessage & "' WHERE Username = '" & passedUserName & "'")
        Next
        If myLogResult = "Change log for " & passedUserName & ".<br/><br/>" Then
            myLogResult &= "Processed successful."
        End If
        GoTo EmailSender
errorHandler:
        myLogResult &= "Error: " & strEachDC & " : " & Err.Description
EmailSender:
        Dim strUserEmail As String = m_host.GetSQL("SELECT PluginUserEmail FROM plugin_itsc_msp_accounts_userstatus WHERE `Username` = '" & passedUserName & "'")
        Dim logHost As LabTech.Interfaces.ICoreFunctionality
        logHost = DirectCast(m_host, LabTech.Interfaces.ICoreFunctionality)
        logHost.SendEmail(body:=myLogResult, subject:="MSP Accounts Log", from:=strUserEmail, toEmail:=strUserEmail)
        logHost = Nothing
    End Sub

End Class
