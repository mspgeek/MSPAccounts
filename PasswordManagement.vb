Imports LabTech.Interfaces

Public Class PasswordManagement


    Public Shared Function ValidatePassword(ByVal pwd As String, ByVal minLength As Integer, ByVal numUpper As Integer, ByVal numLower As Integer, ByVal numNumber As Integer, ByVal numSpecial As Integer)
        ' Replace [A-Z] with \p{Lu}, to allow for Unicode uppercase letters.
        Dim upper As New System.Text.RegularExpressions.Regex("[A-Z]")
        Dim lower As New System.Text.RegularExpressions.Regex("[a-z]")
        Dim number As New System.Text.RegularExpressions.Regex("[0-9]")
        Dim special As New System.Text.RegularExpressions.Regex("[^A-Za-z0-9]")
        ' Check the length.
        If Len(pwd) < minLength Then Return False
        ' Check for minimum number of occurrences.
        If upper.Matches(pwd).Count < numUpper Then Return False
        If lower.Matches(pwd).Count < numLower Then Return False
        If number.Matches(pwd).Count < numNumber Then Return False
        If special.Matches(pwd).Count < numSpecial Then Return False
        ' mine
        If System.Text.RegularExpressions.Regex.IsMatch(pwd, "[\|\&\%\'\`\""\;\\\s]") Then Return False
        If pwd.ToUpper().Contains("12345") Or pwd.ToUpper().Contains("PASSWORD") Or pwd.ToUpper().Contains("QWERTY") Or pwd.ToUpper().Contains("ABC123") Or pwd.ToUpper().Contains("11111") Or pwd.ToUpper().Contains("ILOVEYOU") Or pwd.ToUpper().Contains("123123") Or pwd.ToUpper().Contains("LETMEIN") Or pwd.ToUpper().Contains("BASEBALL") Or pwd.ToUpper().Contains("DRAGON") Then
            Return False
        End If
        ' Passed all checks.
        Return True
    End Function


    Public Shared Function randomPassword(passedLength As Integer, passedUpper As Integer, passedLower As Integer, passedNumber As Integer, passedSpecial As Integer)
        Dim myLowercase As String = "abcdefghijklmnopqrstuvwxyz"
        Dim myUppercase As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Dim myNumbers As String = "0123456789"
        Dim mySymbols As String = "~!@#$^*()_-+={}[]:<>,.?/"
        Dim myAllChars As String = myLowercase & myUppercase & myNumbers & mySymbols
        Dim myRandom As New System.Random
        Dim myPassword As String = ""
        Dim passedMinimums As Integer = passedUpper + passedLower + passedNumber + passedSpecial
        If passedLength < passedMinimums Then
            passedLength = passedMinimums
        End If
        For i As Integer = 1 To passedUpper
            myPassword = myPassword & myUppercase(myRandom.Next(0, myUppercase.Length))
        Next
        For i As Integer = 1 To passedLower
            myPassword = myPassword & myLowercase(myRandom.Next(0, myLowercase.Length))
        Next
        For i As Integer = 1 To passedNumber
            myPassword = myPassword & myNumbers(myRandom.Next(0, myNumbers.Length))
        Next
        For i As Integer = 1 To passedSpecial
            myPassword = myPassword & mySymbols(myRandom.Next(0, mySymbols.Length))
        Next
        If myPassword.Length < passedLength Then
            For i As Integer = 1 To (passedLength - myPassword.Length)
                myPassword = myPassword & myAllChars(myRandom.Next(0, myAllChars.Length))
            Next
        End If
        Dim strInput As String = myPassword
        Dim strOutput As String = ""
        Dim rand As New System.Random
        Dim intPlace As Integer
        While strInput.Length > 0
            intPlace = rand.Next(0, strInput.Length)
            strOutput += strInput.Substring(intPlace, 1)
            strInput = strInput.Remove(intPlace, 1)
        End While
        myPassword = strOutput
        Return myPassword
    End Function


    Public Shared Sub LoopedChangePassword(subHost As LabTech.Interfaces.IControlCenter, passedUserName As String, passedPassword As String)
        Dim myPassword As String = ""
        Dim myConcatClientDCids As String = ""
        Dim sqlPassword As String = "shinybrowncoat"
        Dim concatData As String = subHost.GetSQL("SELECT CONVERT(CONCAT(MSP_Name,'|',User_Prefix,'|',Exclude_Locations,'|',Service_Account,'|',Min_Password_Length,'|',Min_Password_Upper,'|',Min_Password_Lower,'|',Min_Password_Number,'|',Min_Password_Special) USING UTF8) FROM plugin_itsc_msp_accounts_settings")
        Dim strMSPName As String = concatData.ToString.Split("|")(0)
        Dim strMSPPrefix As String = concatData.ToString.Split("|")(1)
        Dim strServiceName As String = concatData.ToString.Split("|")(3)
        Dim strMinPassLength As String = concatData.ToString.Split("|")(4)
        Dim strMinPassUpper As String = concatData.ToString.Split("|")(5)
        Dim strMinPassLower As String = concatData.ToString.Split("|")(6)
        Dim strMinPassNumber As String = concatData.ToString.Split("|")(7)
        Dim strMinPassSpecial As String = concatData.ToString.Split("|")(8)
        If passedPassword = "Random" Then
            myPassword = randomPassword(strMinPassLength, strMinPassUpper, strMinPassLower, strMinPassNumber, strMinPassSpecial)
        Else
            myPassword = passedPassword
        End If
        subHost.SetSQL("UPDATE plugin_itsc_msp_accounts_users SET Password = AES_ENCRYPT(CONVERT('" & myPassword & "' USING utf8),SHA('" & sqlPassword & "')) WHERE Username = '" & passedUserName & "'")
        Dim arrThreads As New List(Of System.Threading.Thread)
        subHost.SetSQL("UPDATE plugin_itsc_msp_accounts_users SET `AutoChangeDate` = CURDATE() WHERE `Username` = '" & passedUserName & "'")
        Dim mySQL As String = "SELECT computers.ClientID, computers.LocationID, computers.ComputerID, TRIM(LEADING 'DC:' FROM computers.Domain) AS Domain FROM Computers, v_detectedroles WHERE v_detectedroles.ComputerID = Computers.ComputerID AND v_detectedroles.RoleName = 'AD PDC Emulator' AND computers.UserAccounts LIKE '%:" & strMSPPrefix & passedUserName & "%'"
        Dim ds As DataSet = subHost.GetDataSet(mySQL)
        For Each DataRow As DataRow In ds.Tables(0).Rows
            Threading.Thread.Sleep(1000)
            Dim currentClientID As String = DataRow("ClientID").ToString
            Dim currentLocationID As String = DataRow("LocationID").ToString
            Dim currentComputerID As String = DataRow("ComputerID").ToString
            Dim currentDomain As String = DataRow("Domain").ToString
            Convert.ToInt32(currentClientID)
            Convert.ToInt32(currentLocationID)
            If passedUserName = strServiceName Then
                Dim passwordKey As String = " " & CStr(currentClientID + 1)
                If passedPassword = "Random" Then
                    myPassword = randomPassword(strMinPassLength, strMinPassUpper, strMinPassLower, strMinPassNumber, strMinPassSpecial)
                End If
                subHost.SetSQL("UPDATE passwords SET Password = AES_ENCRYPT(CONVERT('" & myPassword & "' USING utf8),SHA('" & passwordKey & "')) WHERE Title = '" & strMSPName & " Service Account - " & currentDomain & "'")
            End If
            myConcatClientDCids &= "|" & currentComputerID
            Dim objParams(10) As Object
            objParams(0) = subHost
            objParams(1) = currentComputerID
            objParams(2) = passedUserName
            objParams(3) = strMSPName
            objParams(4) = strMSPPrefix
            objParams(5) = myPassword
            objParams(6) = currentDomain
            Dim t As New System.Threading.Thread(AddressOf PasswordChange)
            arrThreads.Add(t)
            t.Start(objParams)
        Next
        For Each t As System.Threading.Thread In arrThreads
            t.Join(180000)
        Next
        myConcatClientDCids = myConcatClientDCids.TrimStart("|")
        Dim userEmailAddress As String = subHost.GetSQL("SELECT `Email` FROM users WHERE `Name` = '" & passedUserName & "'")
        subHost.SetSQL("INSERT INTO plugin_itsc_msp_accounts_userstatus (`Username`, `PluginUserEmail`, `TimeStamp`, `ClientDCids`) VALUES ('" & passedUserName & "', '" & userEmailAddress & "', NOW(), '" & myConcatClientDCids & "') ON DUPLICATE KEY UPDATE `PluginUserEmail` = '" & userEmailAddress & "', `TimeStamp` = NOW(), `ClientDCids` = '" & myConcatClientDCids & "';")
        Reporting.log_Reporting(subHost, passedUserName)
    End Sub


    Public Shared Sub PasswordChange(ByVal objParams As Object)
        Dim subHost As LabTech.Interfaces.IControlCenter = objParams(0)
        Dim currentComputerID As Integer = objParams(1)
        Dim User_Name As String = objParams(2)
        Dim strMSPName As String = objParams(3)
        Dim strMSPPrefix As String = objParams(4)
        Dim userPassword As String = objParams(5)
        Dim strDomainName As String = objParams(6)
        Dim cmdOutput As String = ""
        Dim strLog As String = ""
        Dim userName As String = strMSPPrefix & User_Name
        Dim mspOU As String = strMSPName & " User Accounts"
        '
        Dim LDAPbase As String = ""
        Dim arrLDAP As String() = strDomainName.Split(New Char() {"."c})
        For Each partLDAP In arrLDAP
            LDAPbase &= "DC=" & partLDAP & ","
        Next
        LDAPbase = LDAPbase.TrimEnd(",")
        '
        Dim strCMD As String = "cmd!!!/C dsmod user " & Chr(34) & "CN=" & userName & ",OU=" & mspOU & "," & LDAPbase & Chr(34) & " -pwd " & Chr(34) & userPassword & Chr(34) & " -pwdneverexpires no"
        Dim cmdID As Integer = subHost.SendCommand(currentComputerID, 2, strCMD)
        Do While CInt(subHost.GetSQL("Select Status from commands where cmdid=" & cmdID)) < 3
            Threading.Thread.Sleep(5000)
        Loop
        cmdOutput = subHost.GetCommandOutput(cmdID)
        cmdID = Nothing
        If Not cmdOutput.Contains("succeeded") Then
            Dim cmdID1 As Integer = subHost.SendCommand(currentComputerID, 2, "cmd!!!/C net user " & userName & " " & Chr(34) & userPassword & Chr(34))
            Do While CInt(subHost.GetSQL("Select Status from commands where cmdid=" & cmdID1)) < 3
                Threading.Thread.Sleep(5000)
            Loop
            cmdOutput = subHost.GetCommandOutput(cmdID1)
        End If
        '
        Threading.Thread.Sleep(30000)
        Dim cmdID2 As Integer = subHost.SendCommand(currentComputerID, 2, "cmd!!!/C dsquery user     -name " & userName & "     -u " & userName & "     -p " & Chr(34) & userPassword & Chr(34))
        Do While CInt(subHost.GetSQL("Select Status from commands where cmdid=" & cmdID2)) < 3
            Threading.Thread.Sleep(5000)
        Loop
        Dim cmdchkOutput As String = subHost.GetCommandOutput(cmdID2)
        If cmdchkOutput.Contains("failed") Then
            strLog = "Password change failed."
        Else
            strLog = "Password change succeeded."
        End If
        strLog = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "|" & strLog
        subHost.SetSQL("UPDATE plugin_itsc_msp_accounts_userstatus SET `" & strDomainName & "` = '" & strLog & "' WHERE `Username` = '" & User_Name & "'")
        'System.Windows.Forms.MessageBox.Show("UPDATE plugin_itsc_msp_accounts_userstatus SET `" & strDomainName & "` = '" & strLog & "' WHERE `Username` = '" & User_Name & "'")
    End Sub


End Class