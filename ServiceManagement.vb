Imports LabTech.Interfaces

Public Class ServiceManagement


    Public Shared Sub LocalServiceLoader(subHost As LabTech.Interfaces.IControlCenter, passedAction As String)
        On Error GoTo errorHandler
        Dim tempData As String = subHost.GetSQL("SELECT CONVERT(CONCAT(MSP_Name,'|',User_Prefix,'|',Service_Account,'|',Min_Password_Length,'|',Local_Service_Account_Exclude,'|',Min_Password_Upper,'|',Min_Password_Lower,'|',Min_Password_Number,'|',Min_Password_Special) USING UTF8) FROM plugin_itsc_msp_accounts_settings")
        Dim strMSPName As String = tempData.ToString.Split("|")(0)
        Dim strMSPPrefix As String = tempData.ToString.Split("|")(1)
        Dim strServiceAccount As String = tempData.ToString.Split("|")(2)
        Dim strMinPassLength As String = tempData.ToString.Split("|")(3)
        Dim strLocalAccountsExclude As String = tempData.ToString.Split("|")(4)
        Dim strMinPassUpper As String = tempData.ToString.Split("|")(5)
        Dim strMinPassLower As String = tempData.ToString.Split("|")(6)
        Dim strMinPassNumber As String = tempData.ToString.Split("|")(7)
        Dim strMinPassSpecial As String = tempData.ToString.Split("|")(8)
        Dim strFullServiceAccountName As String = strMSPPrefix & strServiceAccount
        If passedAction = "Delete" Then
            subHost.SetSQL("UPDATE locations SET PasswordID = 0 WHERE LocationID IN (SELECT LocationID FROM passwords WHERE Title LIKE '" & strMSPName & " Local Service Account - %')")
            subHost.SetSQL("DELETE FROM passwords WHERE Title LIKE '" & strMSPName & " Local Service Account - %'")
        End If
        Dim strSQLLocationSearch As String = "SELECT DISTINCT(c.LocationID), c.ClientID FROM computers AS c WHERE c.Name = SUBSTRING_INDEX(c.Username,'\\',1) AND c.Domain = SUBSTRING_INDEX(c.Domain,'\.',1) AND TIMESTAMPDIFF(HOUR,c.LastContact,NOW()) < 1"
        'Dim ds As DataSet = subHost.GetDataSet(strSQLLocationSearch)
        Dim ds As DataSet
        If strLocalAccountsExclude = True Then
            ds = subHost.GetDataSet(strSQLLocationSearch)
        Else
            ds = subHost.GetDataSet(strSQLLocationSearch & " AND FIND_IN_SET(c.LocationID, (SELECT Exclude_Locations FROM plugin_itsc_msp_accounts_settings)) = 0")
        End If
        For Each DataRow As DataRow In ds.Tables(0).Rows
            Threading.Thread.Sleep(1000)
            Dim strClientID As String = DataRow("ClientID").ToString
            Dim strLocationID As String = DataRow("LocationID").ToString
            Dim strPasswordKey As String = " " & CStr(DataRow("ClientID") + 1)
            Dim strPassword As String = PasswordManagement.randomPassword(strMinPassLength, strMinPassUpper, strMinPassLower, strMinPassNumber, strMinPassSpecial)
            Dim strTitle As String = strMSPName & " Local Service Account - " & strLocationID
            Dim strPasswordID As Integer = subHost.GetSQL("SELECT PasswordID FROM passwords WHERE Title = '" & strTitle & "'")
            Dim strSQLComputerSearch As String = "SELECT c.ComputerID, c.Name FROM computers AS c WHERE c.Name = SUBSTRING_INDEX(c.Username,'\\',1) AND c.Domain = SUBSTRING_INDEX(c.Domain,'\.',1) AND TIMESTAMPDIFF(HOUR,c.LastContact,NOW()) < 1"
            If passedAction = "Add" Then
                ' Create Location Account
                If strPasswordID < 1 Then
                    subHost.SetSQL("INSERT INTO passwords (ClientID, LocationID, Title, UserName, Password) VALUES (" & strClientID & "," & strLocationID & ",'" & strTitle & "','.\\" & strFullServiceAccountName & "',AES_ENCRYPT(CONVERT('" & strPassword & "' USING utf8),SHA('" & strPasswordKey & "')))")
                    strPasswordID = subHost.GetSQL("SELECT PasswordID FROM passwords WHERE Title = '" & strTitle & "'")
                Else
                    strPassword = subHost.GetSQL("SELECT AES_DECRYPT(p.Password,SHA('" & strPasswordKey & "')) FROM passwords AS p WHERE p.Title = '" & strTitle & "'")
                End If
                If subHost.GetSQL("SELECT PasswordID FROM locations WHERE LocationID = " & strLocationID) < 1 Then
                    subHost.SetSQL("UPDATE locations SET PasswordID = " & strPasswordID & " WHERE LocationID = " & strLocationID)
                End If
                strSQLComputerSearch = strSQLComputerSearch & " AND c.UserAccounts NOT LIKE '%:" & strFullServiceAccountName & ":%' AND c.LocationID = " & strLocationID
            ElseIf passedAction = "Delete" Then
                strSQLComputerSearch = strSQLComputerSearch & " AND c.UserAccounts LIKE '%:" & strFullServiceAccountName & ":%' AND c.LocationID = " & strLocationID
            ElseIf passedAction = "Change" Then
                strSQLComputerSearch = strSQLComputerSearch & " AND c.UserAccounts LIKE '%:" & strFullServiceAccountName & ":%' AND c.LocationID = " & strLocationID
                subHost.SetSQL("UPDATE passwords SET Password = AES_ENCRYPT(CONVERT('" & strPassword & "' USING utf8),SHA('" & strPasswordKey & "')) WHERE Title = '" & strTitle & "'")
            Else
                Continue For
            End If
            Dim objParams(10) As Object
            objParams(0) = subHost
            objParams(1) = passedAction
            objParams(2) = strSQLComputerSearch
            objParams(3) = strFullServiceAccountName
            objParams(4) = strTitle
            objParams(5) = strMSPName & " " & strServiceAccount
            objParams(6) = strPassword
            System.Threading.ThreadPool.QueueUserWorkItem(New System.Threading.WaitCallback(AddressOf ServiceManagement.LocalComputerLoader), objParams)
        Next
        Exit Sub
errorHandler:
        Windows.Forms.MessageBox.Show(Err.Number & " : " & Err.Description & " in LocalServiceLoader")
    End Sub


    Public Shared Sub LocalComputerLoader(ByVal objParams As Object)
        On Error GoTo errorHandler
        Dim subHost As LabTech.Interfaces.IControlCenter = objParams(0)
        Dim passedAction As String = objParams(1)
        Dim strSQLComputerSearch As String = objParams(2)
        Dim strFullServiceAccountName As String = objParams(3)
        Dim strComment As String = objParams(4)
        Dim strFullName As String = objParams(5)
        Dim strPassword As String = objParams(6)
        Dim strNetCommand As String = ""
        Dim ds As DataSet = subHost.GetDataSet(strSQLComputerSearch)
        For Each DataRow As DataRow In ds.Tables(0).Rows
            Threading.Thread.Sleep(1000)
            Dim strComputerID As String = DataRow("ComputerID").ToString
            If passedAction = "Add" Then
                strNetCommand = "cmd!!!/C net user /add /fullname:" & Chr(34) & strFullName & Chr(34) & " /comment:" & Chr(34) & strComment & Chr(34) & " " & Chr(34) & strFullServiceAccountName & Chr(34) & " " & Chr(34) & strPassword & Chr(34) & " && net localgroup Administrators " & strFullServiceAccountName & " /add"
                Dim myReg = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts\Userlist:" & strFullServiceAccountName & ":0"
                subHost.SendCommand(strComputerID, 24, myReg)
            ElseIf passedAction = "Delete" Then
                strNetCommand = "cmd!!!/C net user /delete " & Chr(34) & strFullServiceAccountName & Chr(34)
            ElseIf passedAction = "Change" Then
                strNetCommand = "cmd!!!/C net user " & Chr(34) & strFullServiceAccountName & Chr(34) & " " & Chr(34) & strPassword & Chr(34)
            Else
                Continue For
            End If
            Dim objParams2(10) As Object
            objParams2(0) = subHost
            objParams2(1) = strComputerID
            objParams2(2) = strNetCommand
            System.Threading.ThreadPool.QueueUserWorkItem(New System.Threading.WaitCallback(AddressOf ServiceManagement.LocalComputerThread), objParams2)
        Next
        Exit Sub
errorHandler:
        Windows.Forms.MessageBox.Show(Err.Number & " : " & Err.Description & " in LocalComputerLoader")
    End Sub


    Public Shared Sub LocalComputerThread(ByVal objParams2 As Object)
        On Error GoTo errorHandler
        Dim subHost As LabTech.Interfaces.IControlCenter = objParams2(0)
        Dim strComputerID As String = objParams2(1)
        Dim strNetCommand As String = objParams2(2)
        Dim cmdID As Integer = subHost.SendCommand(strComputerID, 2, strNetCommand)
        Do While CInt(subHost.GetSQL("Select Status from commands where cmdid=" & cmdID)) < 3
            Threading.Thread.Sleep(5000)
        Loop
        Threading.Thread.Sleep(30000)
        cmdID = subHost.SendCommand(strComputerID, 123, strComputerID)
        Do While CInt(subHost.GetSQL("Select Status from commands where cmdid=" & cmdID)) < 3
            Threading.Thread.Sleep(5000)
        Loop
        cmdID = Nothing
        Exit Sub
errorHandler:
        Windows.Forms.MessageBox.Show(Err.Number & " : " & Err.Description & " : in LocalComputerThread")
    End Sub


End Class
