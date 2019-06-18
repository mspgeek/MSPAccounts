Imports LabTech.Interfaces


Public Class UserManagement


    Public Shared Sub createUserPrep(ByVal obj0 As Object, ByRef str1 As String, ByRef str2 As String, ByRef str3 As String)
        Dim subHost As LabTech.Interfaces.IControlCenter = obj0
        Dim currentComputerID As Integer = CInt(str1)
        Dim strMSPName As String = str2
        Dim strDomainName As String = str3
        Dim cmdOutput As String = ""
        Dim mspOU As String = strMSPName & " User Accounts"
        ' get Domain Admins group LDAP
        Dim cmdID As Integer = subHost.SendCommand(currentComputerID, 2, "cmd!!!/C dsquery group domainroot -name " & Chr(34) & "Domain Admins" & Chr(34))
        Do While CInt(subHost.GetSQL("Select Status from commands where cmdid=" & cmdID)) < 3
            Threading.Thread.Sleep(5000)
        Loop
        cmdOutput = subHost.GetCommandOutput(cmdID)
        cmdID = Nothing
        Dim domainAdminsLDAP As String = cmdOutput.Replace(Chr(34), "").Replace(vbCr, "").Replace(vbLf, "")
        ' get base LDAP
        Dim LDAPbase As String = ""
        Dim arrLDAP As String() = strDomainName.Split(New Char() {"."c})
        For Each partLDAP In arrLDAP
            LDAPbase &= "DC=" & partLDAP & ","
        Next
        LDAPbase = LDAPbase.TrimEnd(",")
        Dim mspLDAP As String = "OU=" & mspOU & "," & LDAPbase
        ' does the MSP OU exist
        Dim cmdID1 As Integer = subHost.SendCommand(currentComputerID, 2, "cmd!!!/C dsquery ou -scope onelevel -name " & Chr(34) & mspOU & Chr(34))
        Do While CInt(subHost.GetSQL("Select Status from commands where cmdid=" & cmdID1)) < 3
            Threading.Thread.Sleep(5000)
        Loop
        cmdOutput = subHost.GetCommandOutput(cmdID1)
        ' if necessary create MSP OU
        If Not cmdOutput.Contains(mspOU) Then
            Dim cmdID2 As Integer = subHost.SendCommand(currentComputerID, 2, "cmd!!!/C dsadd ou " & Chr(34) & mspLDAP & Chr(34))
            Do While CInt(subHost.GetSQL("Select Status from commands where cmdid=" & cmdID2)) < 3
                Threading.Thread.Sleep(5000)
            Loop
            cmdOutput = subHost.GetCommandOutput(cmdID2)
        End If
        str1 = mspLDAP
        str2 = domainAdminsLDAP
    End Sub


    Public Shared Sub UserCreate(ByVal objParams As Object)
        Dim subHost As LabTech.Interfaces.IControlCenter = objParams(0)
        Dim currentComputerID As Integer = objParams(1)
        Dim User_Name As String = objParams(2)
        Dim strMSPName As String = objParams(3)
        Dim strMSPPrefix As String = objParams(4)
        Dim userPassword As String = objParams(5)
        Dim strDomainName As String = objParams(6)
        Dim runSendInfo As String = objParams(9)
        On Error GoTo errorHandler
        If runSendInfo = "yes" Then
            Dim obj0 As Object = subHost
            Dim str1 As String = currentComputerID
            Dim str2 As String = strMSPName
            Dim str3 As String = strDomainName
            UserManagement.createUserPrep(obj0, str1, str2, str3)
            objParams(7) = str1
            objParams(8) = str2
        End If
        Dim mspLDAP As String = objParams(7)
        Dim domainAdminsLDAP As String = objParams(8)
        If User_Name = "Select User" Then Exit Sub
        Dim cmdOutput As String = ""
        Dim cmdgrpOutput As String = ""
        Dim strLog As String = ""
        Dim userName As String = strMSPPrefix & User_Name
        Dim userDisplayName As String = strMSPName & " " & User_Name
        Dim userDescription As String = strMSPName & " Support Account"
        Dim mspOU As String = strMSPName & " User Accounts"
        Dim cmdID As Integer = subHost.SendCommand(currentComputerID, 2, "cmd!!!/C dsadd user " & Chr(34) & "CN=" & userName & "," & mspLDAP & Chr(34) & " -samid " & Chr(34) & userName & Chr(34) & " -pwd " & Chr(34) & userPassword & Chr(34) & " -display " & Chr(34) & userDisplayName & Chr(34) & " -desc " & Chr(34) & userDescription & Chr(34) & " -memberof " & Chr(34) & domainAdminsLDAP & Chr(34))
        Do While CInt(subHost.GetSQL("Select Status from commands where cmdid=" & cmdID)) < 3
            Threading.Thread.Sleep(5000)
        Loop
        cmdOutput = subHost.GetCommandOutput(cmdID)
        cmdID = Nothing
        If cmdOutput.Contains("created successfully but there was an error") Then
            Threading.Thread.Sleep(30000)
            Dim cmdID1 As Integer = subHost.SendCommand(currentComputerID, 2, "net group " & Chr(34) & "Domain Admins" & Chr(34) & " " & userName & " /add")
            Do While CInt(subHost.GetSQL("Select Status from commands where cmdid=" & cmdID1)) < 3
                Threading.Thread.Sleep(5000)
            Loop
            cmdgrpOutput = subHost.GetCommandOutput(cmdID1)
        End If
        Threading.Thread.Sleep(30000)
        Dim cmdID2 As Integer = subHost.SendCommand(currentComputerID, 2, "cmd!!!/C net user " & userName)
        Do While CInt(subHost.GetSQL("Select Status from commands where cmdid=" & cmdID2)) < 3
            Threading.Thread.Sleep(5000)
        Loop
        Dim cmdchkOutput As String = subHost.GetCommandOutput(cmdID2)
        If cmdchkOutput.Contains("user name could not be found") Then
            strLog = "Creation of user failed."
        ElseIf Not cmdchkOutput.Contains("Domain Admins") Then
            strLog = "User created but failed to add to Domain Admins."
        Else
            Dim tempData As String = subHost.GetSQL("SELECT CONVERT(CONCAT(p.Service_Account,'|',p.Min_Password_Length,'|',c.ClientID,'|',c.LocationID,'|',p.Min_Password_Upper,'|',p.Min_Password_Lower,'|',p.Min_Password_Number,'|',p.Min_Password_Special) USING UTF8) FROM computers AS c, plugin_itsc_msp_accounts_settings AS p WHERE c.ComputerID = " & currentComputerID)
            Dim strServiceName As String = tempData.ToString.Split("|")(0)
            Dim strMinPassLength As String = tempData.ToString.Split("|")(1)
            Dim strClientID As String = tempData.ToString.Split("|")(2)
            Dim strLocationID As String = tempData.ToString.Split("|")(3)
            Dim strMinPassUpper As String = tempData.ToString.Split("|")(4)
            Dim strMinPassLower As String = tempData.ToString.Split("|")(5)
            Dim strMinPassNumber As String = tempData.ToString.Split("|")(6)
            Dim strMinPassSpecial As String = tempData.ToString.Split("|")(7)
            If User_Name = strServiceName Then
                Dim strPasswordKey As String = " " & CStr(strClientID + 1)
                Dim strPassword As String = PasswordManagement.randomPassword(strMinPassLength, strMinPassUpper, strMinPassLower, strMinPassNumber, strMinPassSpecial)
                Dim strTitle As String = strMSPName & " Service Account - " & strDomainName
                Dim strPasswordID As Integer = subHost.GetSQL("SELECT PasswordID FROM passwords WHERE Title = '" & strTitle & "'")
                If strPasswordID < 1 Then
                    subHost.SetSQL("INSERT INTO passwords (ClientID, LocationID, Title, UserName, Password) VALUES (" & strClientID & "," & strLocationID & ",'" & strTitle & "','" & strDomainName & "\\" & strMSPPrefix & strServiceName & "',AES_ENCRYPT(CONVERT('" & userPassword & "' USING utf8),SHA('" & strPasswordKey & "')))")
                    strPasswordID = subHost.GetSQL("SELECT PasswordID FROM passwords WHERE Title = '" & strTitle & "'")
                Else
                    subHost.SetSQL("UPDATE passwords SET Password = AES_ENCRYPT(CONVERT('" & strPassword & "' USING utf8),SHA('" & strPasswordKey & "')) WHERE Title = '" & strTitle & "'")
                End If
                subHost.SetSQL("UPDATE locations SET PasswordID = " & strPasswordID & " WHERE LocationID = " & strLocationID)
            End If
            strLog = "Creation of user succeeded."
        End If
        If runSendInfo = "yes" Then
            Threading.Thread.Sleep(180000)
            Dim cmdIDhardinfo As Integer = subHost.SendCommand(currentComputerID, 17, currentComputerID) ' SendHardwareInfo (users)
            Do While CInt(subHost.GetSQL("Select Status from commands where cmdid=" & cmdIDhardinfo)) < 3
                Threading.Thread.Sleep(5000)
            Loop
            Dim cmdIDsysinfo As Integer = subHost.SendCommand(currentComputerID, 123, currentComputerID) ' SendSystemInfo
        End If
        strLog = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "|" & strLog
        subHost.SetSQL("UPDATE plugin_itsc_msp_accounts_userstatus SET `" & strDomainName & "` = '" & strLog & "' WHERE `Username` = '" & User_Name & "'")
        Exit Sub
errorHandler:
        Dim strLogErr As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "|"
        subHost.SetSQL("UPDATE plugin_itsc_msp_accounts_userstatus SET `" & strDomainName & "` = '" & strLogErr & Err.Description & " while creating user.' WHERE `Username` = '" & User_Name & "'")
    End Sub


    Public Shared Sub UserDelete(ByVal objParams As Object)
        Dim subHost As LabTech.Interfaces.IControlCenter = objParams(0)
        Dim currentComputerID As Integer = objParams(1)
        Dim User_Name As String = objParams(2)
        Dim strMSPName As String = objParams(3)
        Dim strMSPPrefix As String = objParams(4)
        Dim strDomainName As String = objParams(6)
        Dim runSendInfo As String = objParams(9)
        On Error GoTo errorHandler
        Dim cmdOutput As String = ""
        Dim strLog As String = ""
        Dim userName As String = strMSPPrefix & User_Name
        Dim mspOU As String = strMSPName & " User Accounts"
        ' get base LDAP
        Dim LDAPbase As String = ""
        Dim arrLDAP As String() = strDomainName.Split(New Char() {"."c})
        For Each partLDAP In arrLDAP
            LDAPbase &= "DC=" & partLDAP & ","
        Next
        LDAPbase = LDAPbase.TrimEnd(",")
        Dim mspLDAP As String = "OU=" & mspOU & "," & LDAPbase
        ' delete user
        Dim cmdID As Integer = subHost.SendCommand(currentComputerID, 2, "cmd!!!/C dsrm " & Chr(34) & "CN=" & userName & "," & mspLDAP & Chr(34) & " -noprompt")
        Do While CInt(subHost.GetSQL("Select Status from commands where cmdid=" & cmdID)) < 3
            Threading.Thread.Sleep(5000)
        Loop
        cmdOutput = subHost.GetCommandOutput(cmdID)
        cmdID = Nothing
        If Not cmdOutput.Contains("succeeded") Then
            Dim cmdID1 As Integer = subHost.SendCommand(currentComputerID, 2, "cmd!!!/C net user " & userName & " /delete")
            Do While CInt(subHost.GetSQL("Select Status from commands where cmdid=" & cmdID1)) < 3
                Threading.Thread.Sleep(5000)
            Loop
            cmdOutput = subHost.GetCommandOutput(cmdID1)
        End If
        Threading.Thread.Sleep(30000)
        Dim cmdID2 As Integer = subHost.SendCommand(currentComputerID, 2, "cmd!!!/C net user " & userName)
        Do While CInt(subHost.GetSQL("Select Status from commands where cmdid=" & cmdID2)) < 3
            Threading.Thread.Sleep(5000)
        Loop
        Dim cmdchkOutput As String = subHost.GetCommandOutput(cmdID2)
        If cmdchkOutput.Contains("user name could not be found") Then
            strLog = "Deletion of user succeeded."
        Else
            strLog = "Deletion of user failed."
        End If
        If runSendInfo = "yes" Then
            Threading.Thread.Sleep(180000)
            Dim cmdIDhardinfo As Integer = subHost.SendCommand(currentComputerID, 17, currentComputerID) ' SendHardwareInfo (users)
            Do While CInt(subHost.GetSQL("Select Status from commands where cmdid=" & cmdIDhardinfo)) < 3
                Threading.Thread.Sleep(5000)
            Loop
            Dim cmdIDsysinfo As Integer = subHost.SendCommand(currentComputerID, 123, currentComputerID) ' SendSystemInfo
        End If
        strLog = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "|" & strLog
        subHost.SetSQL("UPDATE plugin_itsc_msp_accounts_userstatus SET `" & strDomainName & "` = '" & strLog & "' WHERE `Username` = '" & User_Name & "'")
        Exit Sub
errorHandler:
        Dim strLogErr As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "|"
        subHost.SetSQL("UPDATE plugin_itsc_msp_accounts_userstatus SET `" & strDomainName & "` = '" & strLogErr & Err.Description & " while creating user.') WHERE `Username` = '" & User_Name & "'")
    End Sub


    Public Shared Sub verifyAccount(ByVal objParams As Object, ByRef cmdOutput As String)
        Dim subHost As LabTech.Interfaces.IControlCenter = objParams(0)
        Dim currentComputerID As Integer = objParams(1)
        Dim User_Name As String = objParams(2)
        Dim strMSPName As String = objParams(3)
        Dim strMSPPrefix As String = objParams(4)
        Dim userPassword As String = objParams(5)
        Dim strDomainName As String = objParams(6)
        Dim userName As String = strMSPPrefix & User_Name
        Dim sqlPassword As String = "shinybrowncoat"
        Dim concatData As String = subHost.GetSQL("SELECT CONCAT(clientID, '|', locationID) FROM computers WHERE computerID = " & currentComputerID)
        Dim currentClientID As String = concatData.ToString.Split("|")(0)
        Dim currentLocationID As String = concatData.ToString.Split("|")(1)
        Dim passwordKey As String = " " & CStr(currentClientID + 1)
        Dim actualServiceAccountID As Integer = subHost.GetSQL("SELECT PasswordID FROM locations WHERE LocationID = " & currentLocationID)
        If actualServiceAccountID < 1 Then actualServiceAccountID = 0
        subHost.SetSQL("INSERT INTO passwords (ClientID, LocationID, Title, UserName, Password) VALUES (" & currentClientID & "," & currentLocationID & ",'Testing Account','" & strDomainName & "\\" & userName & "',AES_ENCRYPT(CONVERT('" & userPassword & "' USING utf8),SHA('" & passwordKey & "')))")
        Dim tempServiceAccountID As Integer = subHost.GetSQL("SELECT PasswordID FROM passwords WHERE Title = 'Testing Account' and LocationID = " & currentLocationID)
        subHost.SetSQL("UPDATE locations SET PasswordID = " & tempServiceAccountID & " WHERE LocationID = " & currentLocationID)
        Dim cmdID As String = subHost.SendCommand(12, 120, "cmd /c whoami")
        Do While CInt(subHost.GetSQL("Select Status from commands where cmdid=" & cmdID)) < 3
            Threading.Thread.Sleep(5000)
        Loop
        cmdOutput = subHost.GetCommandOutput(cmdID)
        subHost.SetSQL("UPDATE locations SET PasswordID = " & actualServiceAccountID & " WHERE LocationID = " & currentLocationID)
        subHost.SetSQL("DELETE FROM passwords WHERE Title = 'Testing Account' and LocationID = " & currentLocationID)

    End Sub
End Class
