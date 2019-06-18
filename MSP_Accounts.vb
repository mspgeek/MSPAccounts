Imports System.Windows.Forms
Imports System.Threading
Imports LabTech.Interfaces


Public Class MSP_Accounts_Form
    Private ID As Integer
    Private mComputerID As Integer
    Private mobjhost As IControlCenter
    Private LabtechUserSecurityAdmin As Boolean
    Private LabtechUserID As String
    Private LabtechUserName As String
    Private strMSPPrefix As String
    Private strMSPName As String
    Private strServiceName As String
    Private sqlstr_ClientDCids As String
    Private sqlstr_ClientDCids_Exclusions As String
    Private sqlstr_Domains As String
    Private sqlstr_LoadAllLocations As String
    Private sqlstr_LoadAllLocations_Exclusions As String
    Private sqlstr_LoadAllLocations_Sort As String
    Private sqlstr_msp_accounts_users As String
    Private loopThroughDCsAction As String
    Private labtechServerID As Integer
    Private pwdClass As New PasswordManagement


    Sub New(ByVal objhost As IControlCenter)
        InitializeComponent()
        mobjhost = objhost
        Set_Tab_Permissions()
        mComputerID = ID
        LabtechUserID = mobjhost.GetUser(0).ID
        LabtechUserName = mobjhost.GetUser(0).Name
        labtechServerID = mobjhost.GetSQL("SELECT ComputerID FROM computers WHERE Shares LIKE '%ltshare%' LIMIT 1")
        Dim tempData As String = mobjhost.GetSQL("SELECT CONVERT(CONCAT(MSP_Name,'|',User_Prefix) USING UTF8) FROM plugin_itsc_msp_accounts_settings")
        strMSPName = tempData.ToString.Split("|")(0)
        strMSPPrefix = tempData.ToString.Split("|")(1)
        sqlstr_ClientDCids = "SELECT Computers.ComputerID FROM Computers, v_detectedroles WHERE v_detectedroles.ComputerID = Computers.ComputerID AND v_detectedroles.RoleName = 'AD PDC Emulator'"
        sqlstr_ClientDCids_Exclusions = " AND FIND_IN_SET(Computers.LocationID, (SELECT Exclude_Locations FROM plugin_itsc_msp_accounts_settings)) = 0"
        sqlstr_Domains = "SELECT CONCAT(Computers.ComputerID, '|', TRIM(LEADING 'DC:' FROM computers.Domain)) FROM computers WHERE computers.Computerid IN (" & sqlstr_ClientDCids & ")"
        sqlstr_LoadAllLocations = "SELECT CONCAT(locations.LocationID) as sqlID, CONCAT(clients.Name,' > ',locations.Name) AS sqlName FROM locations, clients WHERE locations.ClientID = clients.ClientID"
        sqlstr_LoadAllLocations_Exclusions = " AND FIND_IN_SET(locations.LocationID, (SELECT Exclude_Locations FROM plugin_itsc_msp_accounts_settings)) = 0"
        sqlstr_LoadAllLocations_Sort = " ORDER BY sqlName"
        sqlstr_msp_accounts_users = "SELECT Username, AutoChangePassword FROM plugin_itsc_msp_accounts_users ORDER BY Username"
        ' SQL Updates Begin
        Dim strColumnExist1 As String = mobjhost.GetSQL("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE table_name = 'plugin_itsc_msp_accounts_settings' AND column_name = 'Local_Service_Account'")
        If strColumnExist1 <> "Local_Service_Account" Then
            mobjhost.SetSQL("ALTER TABLE plugin_itsc_msp_accounts_settings ADD `Local_Service_Account` BINARY(1) DEFAULT 0")
        End If
        Dim strColumnExist2 As String = mobjhost.GetSQL("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE table_name = 'plugin_itsc_msp_accounts_settings' AND column_name = 'Local_Service_Account_Exclude'")
        If strColumnExist2 <> "Local_Service_Account_Exclude" Then
            mobjhost.SetSQL("ALTER TABLE plugin_itsc_msp_accounts_settings ADD `Local_Service_Account_Exclude` BINARY(1) DEFAULT 1")
        End If
        ' SQL Updates End
    End Sub


    Private Sub Set_Tab_Permissions()
        Dim objHost2 As LabTech.Interfaces.IControlCenter2
        objHost2 = CType(mobjhost, LabTech.Interfaces.IControlCenter2)
        LabtechUserSecurityAdmin = objHost2.UserSystemAccess(UserPermissions.ClassConfig)
        objHost2 = Nothing
    End Sub


    Private Sub Reliability_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged
        TabControl1.TabPages.Remove(tabManageUsers)
        TabControl1.TabPages.Remove(tabManageLocations)
        TabControl1.TabPages.Remove(tabServiceAccount)
        TabControl1.TabPages.Remove(tabSettings)
        If LabtechUserSecurityAdmin Then
            TabControl1.TabPages.Add(tabManageUsers)
            TabControl1.TabPages.Add(tabManageLocations)
            TabControl1.TabPages.Add(tabServiceAccount)
            TabControl1.TabPages.Add(tabSettings)
            loadsettings()
        End If
        loadmyuser()
    End Sub


    Private Sub TabControl1_Selected(sender As Object, e As TabControlEventArgs) Handles TabControl1.Selected
        If TabControl1.SelectedTab.Name = "tabManageUsers" Then
            LabelPrefix.Text = strMSPPrefix
            loadManageUsers()
        ElseIf TabControl1.SelectedTab.Name = "tabManageLocations" Then
            loadManageLocations()
        ElseIf TabControl1.SelectedTab.Name = "tabServiceAccount" Then
            loadServiceAccount()
        ElseIf TabControl1.SelectedTab.Name = "tabSettings" Then
            loadsettings()
        ElseIf TabControl1.SelectedTab.Name = "tabMyAccount" Then
            loadmyuser()
        End If
    End Sub


    Sub loadServiceAccount()
        ' load service user
        Dim dsServiceUser As DataSet = mobjhost.GetDataSet(sqlstr_msp_accounts_users)
        dsServiceUser.Tables(0).Rows.Add("None")
        With cbxServiceAccount
            .DisplayMember = "Username"
            .DataSource = dsServiceUser.Tables(0)
        End With
        strServiceName = mobjhost.GetSQL("SELECT Service_Account FROM `plugin_itsc_msp_accounts_settings`")
        cbxServiceAccount.SelectedIndex = cbxServiceAccount.FindStringExact(strServiceName)
        chkLocalAccountsFlag.Checked = mobjhost.GetSQL("SELECT Local_Service_Account FROM `plugin_itsc_msp_accounts_settings`")
        chkLocalAccountsExclude.Checked = mobjhost.GetSQL("SELECT Local_Service_Account_Exclude FROM `plugin_itsc_msp_accounts_settings`")
        If strServiceName = "None" Then
            btnLocalAccountsFlag.Enabled = False
            chkLocalAccountsFlag.Enabled = False
            chkLocalAccountsExclude.Enabled = False
        Else
            btnLocalAccountsFlag.Enabled = True
            chkLocalAccountsFlag.Enabled = True
            chkLocalAccountsExclude.Enabled = True
        End If
    End Sub


    Sub loadsettings()
        On Error GoTo errorHandler
        Dim tempData As String = mobjhost.GetSQL("SELECT CONVERT(CONCAT(MSP_Name,'|',User_Prefix,'|',Exclude_Locations,'|',Service_Account,'|',Min_Password_Length,'|',Password_Change_Days,'|',Min_Password_Upper,'|',Min_Password_Lower,'|',Min_Password_Number,'|',Min_Password_Special) USING UTF8) FROM plugin_itsc_msp_accounts_settings")
        strMSPName = tempData.ToString.Split("|")(0)
        strMSPPrefix = tempData.ToString.Split("|")(1)
        strServiceName = tempData.ToString.Split("|")(3)
        txtMSPNameSettings.Text = strMSPName
        txtMSPPrefixSettings.Text = strMSPPrefix
        txtMinPassword.Text = tempData.ToString.Split("|")(4)
        txtChangeDays.Text = tempData.ToString.Split("|")(5)
        txtMinUpper.Text = tempData.ToString.Split("|")(6)
        txtMinLower.Text = tempData.ToString.Split("|")(7)
        txtMinNumber.Text = tempData.ToString.Split("|")(8)
        txtMinSpecial.Text = tempData.ToString.Split("|")(9)
        lblVersion.Text = "ver. " & System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString
        Exit Sub
errorHandler:
        MessageBox.Show(Err.Description & " while loading settings tab.")
    End Sub


    Private Sub loadManageUsers()
        On Error GoTo errorHandler
        DataGridView1.DataSource = Nothing
        DataGridView1.Columns.Clear()
        Dim ds As DataSet = mobjhost.GetDataSet(sqlstr_msp_accounts_users)
        DataGridView1.DataSource = ds
        With DataGridView1
            .RowHeadersVisible = False
            .AutoGenerateColumns = True
            .DataSource = ds.Tables(0)
        End With
        DataGridView1.AllowUserToAddRows = False
        DataGridView1.Columns("Username").ReadOnly = True
        If DataGridView1.Columns(0).HeaderText <> "Select" And Not DataGridView1.Columns.Contains("Delete") Then
            Dim chk As New DataGridViewCheckBoxColumn()
            DataGridView1.Columns.Insert(0, chk)
            chk.HeaderText = "Delete"
            chk.Name = "checkBoxColumn"
            chk.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
        End If
        Dim usernameColumn As DataGridViewColumn = DataGridView1.Columns("Username")
        usernameColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        Dim autochangeColumn As DataGridViewColumn = DataGridView1.Columns("AutoChangePassword")
        autochangeColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
        '
        Dim dsUserName As DataSet = mobjhost.GetDataSet("SELECT `Name` FROM `users` WHERE `Name` NOT IN (SELECT `Username` FROM `plugin_itsc_msp_accounts_users`) ORDER BY `Name`")
        dsUserName.Tables(0).Rows.Add("Select User")
        With cbUserName
            .ValueMember = "Name"
            .DisplayMember = "Name"
            .DataSource = dsUserName.Tables(0)
            .SelectedIndex = cbUserName.FindStringExact("Select User")
        End With
        txtPassword.Text = "Random"
        Exit Sub
errorHandler:
        MessageBox.Show(Err.Description & " while loading users tab.")
    End Sub


    Private Sub loadManageLocations()
        On Error GoTo errorHandler
        Cursor.Current = Cursors.WaitCursor
        ' load locations
        Dim dsAddLocation As DataSet = mobjhost.GetDataSet(sqlstr_LoadAllLocations & sqlstr_LoadAllLocations_Exclusions & sqlstr_LoadAllLocations_Sort)
        dsAddLocation.Tables(0).Rows.Add(0, "Select Location")
        With cbxTAClocations
            .ValueMember = "sqlID"
            .DisplayMember = "sqlName"
            .DataSource = dsAddLocation.Tables(0)
            .SelectedIndex = cbxTAClocations.FindStringExact("Select Location")
        End With
        ' load users
        Dim dsAddUser As DataSet = mobjhost.GetDataSet(sqlstr_msp_accounts_users)
        dsAddUser.Tables(0).Rows.Add("Select User")
        With cbxTACusers
            .DisplayMember = "Username"
            .DataSource = dsAddUser.Tables(0)
            .SelectedIndex = cbxTACusers.FindStringExact("Select User")
        End With
        cbxTAClocations.SelectedIndex = cbxTAClocations.FindStringExact("Select Location")
        cbxTACusers.SelectedIndex = cbxTACusers.FindStringExact("Select User")
        cbxTACactions.SelectedIndex = cbxTACactions.FindStringExact("Select Action")
        cbxTACxoverride.Checked = False
        cbxTACusers.Enabled = False
        cbxTACactions.Enabled = False
        btnAllLocationsExec.Enabled = False
        DataGridViewLocations.DataSource = Nothing
        DataGridViewLocations.Columns.Clear()
        DataGridViewLocations.AutoGenerateColumns = False
        DataGridViewLocations.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "ID", .Visible = False})
        DataGridViewLocations.Columns.Add(New DataGridViewCheckBoxColumn With {.DataPropertyName = "Exclude", .ThreeState = False, .FalseValue = 0, .TrueValue = 1})
        DataGridViewLocations.Columns.Add(New DataGridViewTextBoxColumn With {.DataPropertyName = "Location"})
        Dim ds As DataSet = mobjhost.GetDataSet("SELECT NOT(FIND_IN_SET(locations.LocationID, (SELECT Exclude_Locations FROM plugin_itsc_msp_accounts_settings)) = 0) AS Exclude, CONCAT(clients.Name,' > ',locations.Name) AS Location, LocationID AS ID FROM locations, clients WHERE locations.ClientID = clients.ClientID")
        DataGridViewLocations.DataSource = ds
        With DataGridViewLocations
            .RowHeadersVisible = False
            .AutoGenerateColumns = False
            .AllowUserToAddRows = False
            .DataSource = ds.Tables(0)
        End With
        DataGridViewLocations.Columns(0).Name = "ID"
        DataGridViewLocations.Columns(1).Name = "Exclude"
        DataGridViewLocations.Columns(1).HeaderText = "Exclude"
        DataGridViewLocations.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
        DataGridViewLocations.Columns(2).HeaderText = "Location"
        DataGridViewLocations.Columns(2).ReadOnly = True
        DataGridViewLocations.Columns(2).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        Cursor.Current = Cursors.Default
        Exit Sub
errorHandler:
        Cursor.Current = Cursors.Default
        MessageBox.Show(Err.Description & " while loading locations tab.")
    End Sub


    Private Sub loadmyuser()
        On Error GoTo errorHandler
        txtMyPassword.PasswordChar = "*"
        btnShowPassword.Text = "Show Password"
        lblMyPrefix.Text = strMSPPrefix
        txtMyUsername.Text = LabtechUserName
        Dim CheckString As String = mobjhost.GetSQL("SELECT COUNT(*) FROM plugin_itsc_msp_accounts_users WHERE Username = '" & LabtechUserName & "'")
        If CheckString = 1 Then
            Dim myPassword As String = mobjhost.GetSQL("SELECT AES_DECRYPT(p.Password,SHA('" & sqlPassword & "')) FROM `plugin_itsc_msp_accounts_users` AS p WHERE p.Username = '" & LabtechUserName & "'")
            txtMyPassword.Text = myPassword
            lblMyWarning.Visible = False
            btnMySavePassword.Visible = True
            txtNewPassword.Visible = True
            lblNewPassword.Visible = True
        Else
            lblMyWarning.Text = "Username Does Not Exists.  Please contact your supervisor."
            lblMyWarning.Visible = True
            btnMySavePassword.Visible = False
            txtNewPassword.Visible = False
            lblNewPassword.Visible = False
        End If
        Exit Sub
errorHandler:
        MessageBox.Show(Err.Description & " while loading account tab.")
    End Sub


    Private Sub btn_SaveSettings_Click(sender As System.Object, e As System.EventArgs) Handles btn_SaveSettings.Click
        On Error GoTo errorHandler
        strMSPName = txtMSPNameSettings.Text
        If Not inputValidation(strMSPName) Then Exit Sub
        strMSPPrefix = txtMSPPrefixSettings.Text
        If Not inputValidation(strMSPPrefix) Then Exit Sub
        Dim strMinPassword As Integer = CInt(txtMinPassword.Text)
        If strMinPassword < 10 Or strMinPassword > 99 Then strMinPassword = 14
        Dim strChangePassword As Integer = CInt(txtChangeDays.Text)
        If strChangePassword < 1 Or strMinPassword > 999 Then strMinPassword = 60
        Dim strMinUpper As Integer = CInt(txtMinUpper.Text)
        If strMinUpper < 0 Or strMinUpper > 99 Then strMinUpper = 2
        Dim strMinLower As Integer = CInt(txtMinLower.Text)
        If strMinLower < 0 Or strMinLower > 99 Then strMinLower = 2
        Dim strMinNumber As Integer = CInt(txtMinNumber.Text)
        If strMinNumber < 0 Or strMinNumber > 99 Then strMinNumber = 2
        Dim strMinSpecial As Integer = CInt(txtMinSpecial.Text)
        If strMinSpecial < 0 Or strMinSpecial > 99 Then strMinSpecial = 2
        Dim strSQL = "UPDATE `plugin_itsc_msp_accounts_settings` SET `MSP_Name` = '" & strMSPName & "', `User_Prefix` = '" & strMSPPrefix & "', `Min_Password_Length` = '" & strMinPassword & "', `Password_Change_Days` = '" & strChangePassword & "', `Min_Password_Upper` = '" & strMinUpper & "', `Min_Password_Lower` = '" & strMinLower & "', `Min_Password_Number` = '" & strMinNumber & "', `Min_Password_Special` = '" & strMinSpecial & "'"
        mobjhost.SetSQL(strSQL)
        loadsettings()
        Exit Sub
errorHandler:
        MessageBox.Show(Err.Description & " while saving settings.")
    End Sub


    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click
        On Error GoTo errorHandler
        Dim response = MsgBox("Confirm Delete?", MsgBoxStyle.YesNo)
        If response = MsgBoxResult.Yes Then
            Cursor.Current = Cursors.WaitCursor
            For Each oRow As DataGridViewRow In DataGridView1.Rows
                If oRow.Cells.Item("checkBoxColumn").Value = True Then
                    Dim myDeadUserName As String = oRow.Cells.Item("Username").Value
                    If myDeadUserName = strServiceName Then
                        MessageBox.Show(strServiceName & " is the current Service Account and as such cannot be deleted.")
                    Else
                        Dim objDCParams(3) As Object
                        objDCParams(0) = "action_UserDelete"
                        objDCParams(1) = myDeadUserName
                        objDCParams(2) = True
                        ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf loopThroughDCs), objDCParams)
                        mobjhost.SetSQL("DELETE FROM plugin_itsc_msp_accounts_users WHERE plugin_itsc_msp_accounts_users.UserName = '" & myDeadUserName & "'")
                    End If
                End If
            Next
            Cursor.Current = Cursors.Default
            MessageBox.Show("Changes can take up to 15 minutes to propagate.")
        End If
        Exit Sub
errorHandler:
        Cursor.Current = Cursors.Default
        MessageBox.Show(Err.Description & " while deleting user.")
    End Sub


    Private Sub btnAddUser_Click_1(sender As System.Object, e As System.EventArgs) Handles btnAddUser.Click
        On Error GoTo errorHandler
        If Not inputValidation(txtPassword.Text) Then Exit Sub
        If cbUserName.SelectedValue <> "Select User" Then
            If dbSavePassword(cbUserName.SelectedValue, txtPassword.Text) Then
                Cursor.Current = Cursors.WaitCursor
                Dim objDCParams(3) As Object
                objDCParams(0) = "action_UserCreate"
                objDCParams(1) = cbUserName.SelectedValue
                objDCParams(2) = True
                ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf loopThroughDCs), objDCParams)
                Cursor.Current = Cursors.Default
                MessageBox.Show("Changes can take up to 15 minutes to propagate.")
            End If
        End If
        Exit Sub
errorHandler:
        Cursor.Current = Cursors.Default
        MessageBox.Show(Err.Description & " while adding user.")
    End Sub


    Private Sub btnMySavePassword_Click(sender As System.Object, e As System.EventArgs) Handles btnMySavePassword.Click
        On Error GoTo errorHandler
        If Not inputValidation(txtNewPassword.Text) Then Exit Sub
        Dim userPassword As String = txtNewPassword.Text
        If dbSavePassword(LabtechUserName, userPassword) Then
            Cursor.Current = Cursors.WaitCursor
            PasswordManagement.LoopedChangePassword(mobjhost, LabtechUserName, userPassword)
            loadmyuser()
            Cursor.Current = Cursors.Default
            MessageBox.Show("You will receive an email when the changes have completed.")
        End If
        Exit Sub
errorHandler:
        Cursor.Current = Cursors.Default
        MessageBox.Show(Err.Description & " while changing password.")
    End Sub


    Private Sub loopThroughDCs(ByVal objDCParams As Object)
        On Error GoTo errorHandler
        Dim mySQL As String
        Dim loopThroughDCsAction = objDCParams(0)
        Dim passedUserName = objDCParams(1)
        Dim useExclusions = objDCParams(2)
        If useExclusions Then
            mySQL = sqlstr_ClientDCids & sqlstr_ClientDCids_Exclusions
        Else
            mySQL = sqlstr_ClientDCids
        End If
        If loopThroughDCsAction = "action_UserDelete" Or loopThroughDCsAction = "action_ChangePassword" Then
            mySQL &= " AND computers.UserAccounts LIKE '%:" & strMSPPrefix & passedUserName & "%'"
        End If
        Dim userPassword As String = mobjhost.GetSQL("SELECT AES_DECRYPT(p.Password,SHA('" & sqlPassword & "')) FROM plugin_itsc_msp_accounts_users AS p WHERE p.Username =  '" & passedUserName & "'")
        ' begin load status columns
        Dim arrStatusColumns = mobjhost.GetValues("SELECT column_name FROM information_schema.columns WHERE table_name='plugin_itsc_msp_accounts_userstatus'")
        Dim dicStatusColumns As New Dictionary(Of String, String)
        For Each strColumnName In arrStatusColumns
            dicStatusColumns.Add(strColumnName, "")
        Next
        Dim arrComputerDomains = mobjhost.GetValues(sqlstr_Domains)
        Dim dicComputerDomains As New Dictionary(Of String, String)
        For Each strDomain As String In arrComputerDomains
            Dim mytmpComputerID As String = strDomain.ToString.Split("|")(0)
            Dim mytmpDomainName As String = strDomain.ToString.Split("|")(1)
            dicComputerDomains.Add(mytmpComputerID, mytmpDomainName)
        Next
        Dim myConcatClientDCids As String = ""
        ' end load status columns
        Dim strThreadingStart As DateTime = Now
        Dim arrThreads As New List(Of Thread)
        Dim computerIDs = mobjhost.GetValues(mySQL)
        For Each computerID As String In computerIDs
            Threading.Thread.Sleep(1000)
            ' begin add status columns as needed
            Dim myComputerDomain As String = dicComputerDomains.Item(computerID)
            If Not dicStatusColumns.ContainsKey(myComputerDomain) Then
                mobjhost.SetSQL("ALTER TABLE plugin_itsc_msp_accounts_userstatus ADD `" & myComputerDomain & "` TEXT(500)")
                dicStatusColumns.Add(myComputerDomain, "")
            End If
            myConcatClientDCids &= "|" & computerID
            ' end add status columns as needed
            Dim objParams(10) As Object
            objParams(0) = mobjhost
            objParams(1) = computerID
            objParams(2) = passedUserName
            objParams(3) = strMSPName
            objParams(4) = strMSPPrefix
            objParams(5) = userPassword
            objParams(6) = myComputerDomain
            Select Case loopThroughDCsAction
                Case "action_UserCreate"
                    objParams(9) = "yes"
                    '                    ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf UserManagement.UserCreate), objParams)
                    Dim t As New Thread(AddressOf UserManagement.UserCreate)
                    arrThreads.Add(t)
                    t.Start(objParams)
                Case "action_UserDelete"
                    objParams(9) = "yes"
                    '                    ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf UserManagement.UserDelete), objParams)
                    Dim t As New Thread(AddressOf UserManagement.UserDelete)
                    arrThreads.Add(t)
                    t.Start(objParams)
                Case "action_ChangePassword"
                    '                    ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf PasswordManagement.PasswordChange), objParams)
                    Dim t As New Thread(AddressOf PasswordManagement.PasswordChange)
                    arrThreads.Add(t)
                    t.Start(objParams)
            End Select
        Next
        For Each t As Thread In arrThreads
            t.Join(180000)
        Next
        'MessageBox.Show("End processing " & strThreadingStart & " > " & Now)
        ' update status
        myConcatClientDCids = myConcatClientDCids.TrimStart("|")
        mobjhost.SetSQL("INSERT INTO plugin_itsc_msp_accounts_userstatus (`Username`, `PluginUserEmail`, `TimeStamp`, `ClientDCids`) VALUES ('" & passedUserName & "','" & mobjhost.GetUser(0).Email & "', NOW(),'" & myConcatClientDCids & "') ON DUPLICATE KEY UPDATE `TimeStamp` = NOW(), `ClientDCids` = '" & myConcatClientDCids & "', `PluginUserEmail` = '" & mobjhost.GetUser(0).Email & "'")
        mobjhost.SetSQL("UPDATE plugin_itsc_msp_accounts_users SET `AutoChangeDate` = CURDATE() WHERE `Username` = '" & passedUserName & "'")
        Reporting.log_Reporting(mobjhost, passedUserName)
        Exit Sub
errorHandler:
        MessageBox.Show(Err.Description & " while looping though PDCs. Computer: computerID, Domain: myComputerDomain")
    End Sub


    Private Sub btnShowPassword_Click(sender As System.Object, e As System.EventArgs) Handles btnShowPassword.Click
        If btnShowPassword.Text = "Show Password" Then
            txtMyPassword.PasswordChar = ControlChars.NullChar
            btnShowPassword.Text = "Hide Password"
        Else
            txtMyPassword.PasswordChar = "*"
            btnShowPassword.Text = "Show Password"
        End If
    End Sub


    Private Sub cbxTACxoverride_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles cbxTACxoverride.CheckedChanged
        On Error GoTo errorHandler
        Cursor.Current = Cursors.WaitCursor
        Dim dsTAC As DataSet
        If cbxTACxoverride.Checked = True Then
            dsTAC = mobjhost.GetDataSet(sqlstr_LoadAllLocations & sqlstr_LoadAllLocations_Sort)
        Else
            dsTAC = mobjhost.GetDataSet(sqlstr_LoadAllLocations & sqlstr_LoadAllLocations_Exclusions & sqlstr_LoadAllLocations_Sort)
        End If
        dsTAC.Tables(0).Rows.Add(0, "Select Location")
        cbxTAClocations.DataSource = dsTAC.Tables(0)
        cbxTAClocations.ValueMember = "sqlID"
        cbxTAClocations.DisplayMember = "sqlName"
        cbxTAClocations.SelectedIndex = cbxTAClocations.FindStringExact("Select Location")
        cbxTAClocations.Refresh()
        Cursor.Current = Cursors.Default
        Exit Sub
errorHandler:
        Cursor.Current = Cursors.Default
        MessageBox.Show(Err.Description & " on TACOverride.")
    End Sub


    Private Function inputValidation(passedString As String)
        If System.Text.RegularExpressions.Regex.IsMatch(passedString, "[\|\&\%\'\`\""\;\\\s]") Then
            MessageBox.Show("You cannot use the characters: ampersand, percent, quote, backslash, vertical bar, semicolon, or space.")
            Return False
        Else
            Return True
        End If
    End Function


    Private Function dbSavePassword(passedUserName As String, passedPassword As String)
        On Error GoTo errorHandler
        Dim finalPassword As String = ""
        Dim concatData As String = mobjhost.GetSQL("SELECT CONVERT(CONCAT(Min_Password_Length,'|',Min_Password_Upper,'|',Min_Password_Lower,'|',Min_Password_Number,'|',Min_Password_Special) USING UTF8) FROM plugin_itsc_msp_accounts_settings")
        Dim strMinPassLength As String = concatData.ToString.Split("|")(0)
        Dim strMinPassUpper As String = concatData.ToString.Split("|")(1)
        Dim strMinPassLower As String = concatData.ToString.Split("|")(2)
        Dim strMinPassNumber As String = concatData.ToString.Split("|")(3)
        Dim strMinPassSpecial As String = concatData.ToString.Split("|")(4)
        If passedPassword = "Random" Then
            finalPassword = PasswordManagement.randomPassword(strMinPassLength, strMinPassUpper, strMinPassLower, strMinPassNumber, strMinPassSpecial)
        ElseIf PasswordManagement.ValidatePassword(passedPassword, strMinPassLength, strMinPassUpper, strMinPassLower, strMinPassNumber, strMinPassSpecial) Then
            finalPassword = passedPassword
        Else
            MessageBox.Show("Passwords must meet a minimum length requirement of " & strMinPassLength & " characters, and contain " & strMinPassUpper & " upper & " & strMinPassLower & " lower case letters, " & strMinPassNumber & " numbers, and " & strMinPassSpecial & " symbols; or you tried to use one of the 10 most common passwords in your password.")
            Return False
        End If
        mobjhost.SetSQL("INSERT INTO `plugin_itsc_msp_accounts_users` (`Username`, `Password`) VALUES ('" & passedUserName & "', AES_ENCRYPT(CONVERT('" & finalPassword & "' USING utf8),SHA('" & sqlPassword & "'))) ON DUPLICATE KEY UPDATE Password = AES_ENCRYPT(CONVERT('" & finalPassword & "' USING utf8),SHA('" & sqlPassword & "'))")
        Return True
        Exit Function
errorHandler:
        MessageBox.Show(Err.Description & " while saving password.")
    End Function


    Private Sub cbxTACusers_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cbxTACusers.SelectedIndexChanged
        If cbxTACusers.Text = "Select User" Then
            btnAllLocationsExec.Enabled = False
        Else
            btnAllLocationsExec.Enabled = True
        End If
    End Sub


    Private Sub cbxTAClocations_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cbxTAClocations.SelectedIndexChanged
        If cbxTAClocations.Text = "Select Location" Then
            cbxTACusers.Text = "Select User"
            cbxTACusers.Enabled = False
            cbxTACactions.Text = "Select Action"
            cbxTACactions.Enabled = False
            btnAllLocationsExec.Enabled = False
        Else
            cbxTACactions.Enabled = True
            If cbxTACactions.Text.Contains("Individual User") Then
                cbxTACusers.Enabled = True
                If cbxTACusers.Text.Contains("Select User") Then
                    btnAllLocationsExec.Enabled = False
                Else
                    btnAllLocationsExec.Enabled = True
                End If
            Else
                cbxTACusers.Enabled = False
            End If
        End If
    End Sub


    Private Sub cbxTACactions_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cbxTACactions.SelectedIndexChanged
        Select Case cbxTACactions.Text
            Case "Add All Users"
                cbxTACusers.Enabled = False
                btnAllLocationsExec.Enabled = True
            Case "Delete All Users"
                cbxTACusers.Enabled = False
                btnAllLocationsExec.Enabled = True
            Case "Add Individual User"
                cbxTACusers.Enabled = True
                If cbxTACusers.Text.Contains("Select User") Then
                    btnAllLocationsExec.Enabled = False
                Else
                    btnAllLocationsExec.Enabled = True
                End If
            Case "Delete Individual User"
                cbxTACusers.Enabled = True
                If cbxTACusers.Text.Contains("Select User") Then
                    btnAllLocationsExec.Enabled = False
                Else
                    btnAllLocationsExec.Enabled = True
                End If
            Case Else
                cbxTACusers.Enabled = False
                btnAllLocationsExec.Enabled = False
        End Select
    End Sub


    Private Sub btnAllLocationsExec_Click(sender As System.Object, e As System.EventArgs) Handles btnAllLocationsExec.Click
        On Error GoTo errorHandler
        Dim myLocation As String = cbxTAClocations.SelectedValue
        Dim myUser As String = cbxTACusers.Text
        Dim myAction As String = cbxTACactions.Text
        If myLocation <> "0" Then
            Dim myComputerID As String = mobjhost.GetSQL("SELECT Computers.ComputerID FROM Computers, v_detectedroles WHERE v_detectedroles.ComputerID = Computers.ComputerID AND v_detectedroles.RoleName = 'AD PDC Emulator' and computers.LocationID = " & myLocation)
            Dim myComputerDomain As String = mobjhost.GetSQL("SELECT TRIM(LEADING 'DC:' FROM computers.Domain) FROM computers WHERE computers.Computerid = " & myComputerID)
            Dim myStatusColumnExist As Integer = mobjhost.GetSQL("SELECT COUNT(*) FROM information_schema.columns WHERE table_name = 'plugin_itsc_msp_accounts_userstatus' AND column_name = '" & myComputerDomain & "'")
            If myStatusColumnExist < 1 Then
                mobjhost.SetSQL("ALTER TABLE plugin_itsc_msp_accounts_userstatus ADD `" & myComputerDomain & "` TEXT(500)")
            End If
            Cursor.Current = Cursors.WaitCursor
            Dim objParams(10) As Object
            objParams(0) = mobjhost
            objParams(1) = myComputerID
            objParams(2) = myUser
            objParams(3) = strMSPName
            objParams(4) = strMSPPrefix
            objParams(6) = myComputerDomain
            Select Case myAction
                Case "Add All Users"
                    ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf location_add), objParams)
                    MessageBox.Show("Changes can take up to 10 minutes to propagate.")
                Case "Delete All Users"
                    ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf location_delete), objParams)
                    MessageBox.Show("Changes can take up to 10 minutes to propagate.")
                Case "Add Individual User"
                    If myUser <> "Select User" Then
                        mobjhost.SetSQL("INSERT INTO plugin_itsc_msp_accounts_userstatus (`Username`, `PluginUserEmail`) VALUES ('" & myUser & "','" & mobjhost.GetUser(0).Email & "') ON DUPLICATE KEY UPDATE `PluginUserEmail` = '" & mobjhost.GetUser(0).Email & "'")
                        Dim myPassword As String = mobjhost.GetSQL("SELECT AES_DECRYPT(p.Password,SHA('" & sqlPassword & "')) FROM `plugin_itsc_msp_accounts_users` AS p WHERE p.Username =  '" & myUser & "'")
                        objParams(5) = myPassword
                        objParams(9) = "yes"
                        ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf UserManagement.UserCreate), objParams)
                        MessageBox.Show("Add user command issued.")
                    Else
                        MessageBox.Show("Please select a Username.")
                    End If
                Case "Delete Individual User"
                    If myUser <> "Select User" Then
                        mobjhost.SetSQL("INSERT INTO plugin_itsc_msp_accounts_userstatus (`Username`, `PluginUserEmail`, `TimeStamp`, `ClientDCids`) VALUES ('" & myUser & "','" & mobjhost.GetUser(0).Email & "', NOW(), '|" & myComputerID & "')) ON DUPLICATE KEY UPDATE `TimeStamp` = NOW(), `ClientDCids` = '" & myComputerID & "', `PluginUserEmail` = '" & mobjhost.GetUser(0).Email & "'")
                        objParams(9) = "yes"
                        ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf UserManagement.UserDelete), objParams)
                        MessageBox.Show("Delete command issued.")
                    Else
                        MessageBox.Show("Please select a Username.")
                    End If
                Case Else
            End Select
            Cursor.Current = Cursors.Default
        Else
            MessageBox.Show("Please select a location.")
        End If
        cbxTAClocations.Text = "Select Location"
        cbxTACusers.Text = "Select User"
        cbxTACusers.Enabled = False
        cbxTACactions.Text = "Select Action"
        cbxTACactions.Enabled = False
        Exit Sub
errorHandler:
        Cursor.Current = Cursors.Default
        MessageBox.Show(Err.Description & " while managing locations.")
    End Sub


    Private Sub btnServiceAccount_Click(sender As System.Object, e As System.EventArgs) Handles btnServiceAccount.Click
        On Error GoTo errorHandler
        Dim response = MsgBox("Confirm Change All Service Accounts?", MsgBoxStyle.YesNo)
        If response = MsgBoxResult.No Then
            Exit Sub
        End If
        strServiceName = cbxServiceAccount.Text
        Dim oldServiceAccount As String = mobjhost.GetSQL("SELECT Service_Account FROM plugin_itsc_msp_accounts_settings")
        If strServiceName = oldServiceAccount Then
            Exit Sub
        End If
        Cursor.Current = Cursors.WaitCursor
        mobjhost.SetSQL("UPDATE plugin_itsc_msp_accounts_settings SET `Service_Account` = '" & strServiceName & "'")
        If oldServiceAccount <> "None" Then
            mobjhost.SetSQL("UPDATE locations SET PasswordID = 0 WHERE LocationID IN (SELECT LocationID FROM passwords WHERE Title LIKE '" & strMSPName & " Service Account - %')")
            mobjhost.SetSQL("DELETE FROM passwords WHERE Title LIKE '" & strMSPName & " Service Account - %'")
            ServiceManagement.LocalServiceLoader(mobjhost, "Delete")
        End If
        If strServiceName <> "None" Then
            mobjhost.SetSQL("UPDATE plugin_itsc_msp_accounts_users SET `AutoChangePassword` = TRUE WHERE `Username` = '" & strServiceName & "'")
            Dim userPassword As String = "Random"
            Dim mySQL As String = "SELECT computers.ClientID, computers.LocationID, TRIM(LEADING 'DC:' FROM computers.Domain) AS Domain FROM Computers, v_detectedroles WHERE v_detectedroles.ComputerID = Computers.ComputerID AND v_detectedroles.RoleName = 'AD PDC Emulator'"
            Dim ds As DataSet = mobjhost.GetDataSet(mySQL)
            For Each DataRow As DataRow In ds.Tables(0).Rows
                Dim currentClientID As String = DataRow("ClientID").ToString
                Dim currentLocationID As String = DataRow("LocationID").ToString
                Dim currentDomain As String = DataRow("Domain").ToString
                Convert.ToInt32(currentClientID)
                Convert.ToInt32(currentLocationID)
                Dim passwordKey As String = " " & CStr(currentClientID + 1)
                mobjhost.SetSQL("INSERT INTO passwords (ClientID, LocationID, Title, UserName, Password) VALUES (" & currentClientID & "," & currentLocationID & ",'" & strMSPName & " Service Account - " & currentDomain & "','" & currentDomain & "\\" & strMSPPrefix & strServiceName & "',AES_ENCRYPT(CONVERT('" & userPassword & "' USING utf8),SHA('" & passwordKey & "')))")
                mobjhost.SetSQL("UPDATE locations SET PasswordID = (SELECT PasswordID FROM passwords WHERE Title = '" & strMSPName & " Service Account - " & currentDomain & "') WHERE LocationID = " & currentLocationID)
            Next
            PasswordManagement.LoopedChangePassword(mobjhost, strServiceName, userPassword)
            btnLocalAccountsFlag.Enabled = True
            chkLocalAccountsFlag.Enabled = True
            chkLocalAccountsExclude.Enabled = True
        Else
            btnLocalAccountsFlag.Enabled = False
            chkLocalAccountsFlag.Enabled = False
            chkLocalAccountsExclude.Enabled = False
            mobjhost.SetSQL("UPDATE plugin_itsc_msp_accounts_settings SET `Local_Service_Account` = FALSE")
        End If
        Cursor.Current = Cursors.Default
        MessageBox.Show("Changes can take up to 15 minutes to propagate.")
        Exit Sub
errorHandler:
        Cursor.Current = Cursors.Default
        MessageBox.Show(Err.Description & " while managing service account.")
    End Sub


    Private Sub btnAutoChange_Click(sender As System.Object, e As System.EventArgs)
        Cursor.Current = Cursors.WaitCursor
        On Error GoTo errorHandler
        For Each thisRow As DataGridViewRow In DataGridView1.Rows
            mobjhost.SetSQL("UPDATE plugin_itsc_msp_accounts_users SET `AutoChangePassword` = " & thisRow.Cells.Item("AutoChangePassword").Value & ", `Password_Change_Days` = CURDATE() WHERE `Username` = '" & thisRow.Cells.Item("Username").Value & "'")
        Next
        Cursor.Current = Cursors.Default
        MessageBox.Show("Auto Change Settings Saved")
        Exit Sub
errorHandler:
        Cursor.Current = Cursors.Default
        MessageBox.Show(Err.Description & " while updating users.")
    End Sub


    Private Sub dataGridView1_CellClick(ByVal sender As Object, ByVal e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        If e.ColumnIndex = 2 Then
            DataGridView1.EndEdit()
            DataGridView1.Rows(e.RowIndex).Cells(2).Value = Not DataGridView1.Rows(e.RowIndex).Cells(2).Value
            DataGridView1.EndEdit()
            Dim changeValue As Boolean = DataGridView1.Rows(e.RowIndex).Cells(2).Value
            Dim thisUsername As String = DataGridView1.Rows(e.RowIndex).Cells(1).Value
            mobjhost.SetSQL("UPDATE plugin_itsc_msp_accounts_users SET `AutoChangePassword` = " & changeValue & ", `AutoChangeDate` = CURDATE() WHERE `Username` = '" & thisUsername & "'")
            MessageBox.Show("Auto Change setting saved for " & thisUsername)
        End If
    End Sub


    Private Sub location_add(ByVal objLocParams As Object)
        Dim myEmail As String = mobjhost.GetUser(0).Email
        Dim myComputerID As String = objLocParams(1)
        Dim myComputerDomain As String = objLocParams(6)
        Dim obj0 As Object = mobjhost
        Dim str1 As String = myComputerID
        Dim str2 As String = strMSPName
        Dim str3 As String = myComputerDomain
        UserManagement.createUserPrep(obj0, str1, str2, str3)
        objLocParams(7) = str1
        objLocParams(8) = str2
        For Each loopedUsername As String In mobjhost.GetValues(sqlstr_msp_accounts_users)
            Threading.Thread.Sleep(1000)
            mobjhost.SetSQL("INSERT INTO plugin_itsc_msp_accounts_userstatus (`Username`, `PluginUserEmail`, `TimeStamp`, `ClientDCids`) VALUES ('" & loopedUsername & "','" & myEmail & "', NOW(), '" & myComputerID & "') ON DUPLICATE KEY UPDATE `TimeStamp` = NOW(), `ClientDCids` = '" & myComputerID & "', `PluginUserEmail` = '" & myEmail & "'")
            Dim myPassword As String = mobjhost.GetSQL("SELECT AES_DECRYPT(p.Password,SHA('" & sqlPassword & "')) FROM `plugin_itsc_msp_accounts_users` AS p WHERE p.Username =  '" & loopedUsername & "'")
            objLocParams(2) = loopedUsername
            objLocParams(5) = myPassword
            ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf UserManagement.UserCreate), objLocParams)
        Next
        Threading.Thread.Sleep(180000)
        Dim cmdIDhardinfo As Integer = mobjhost.SendCommand(myComputerID, 17, myComputerID) ' SendHardwareInfo (users)
        Do While CInt(mobjhost.GetSQL("Select Status from commands where cmdid=" & cmdIDhardinfo)) < 3
            Threading.Thread.Sleep(5000)
        Loop
        Dim cmdIDsysinfo As Integer = mobjhost.SendCommand(myComputerID, 123, myComputerID) ' SendSystemInfo
    End Sub


    Private Sub location_delete(ByVal objLocParams As Object)
        Dim myEmail As String = mobjhost.GetUser(0).Email
        Dim myComputerID As String = objLocParams(1)
        For Each loopedUsername As String In mobjhost.GetValues(sqlstr_msp_accounts_users)
            Threading.Thread.Sleep(1000)
            mobjhost.SetSQL("INSERT INTO plugin_itsc_msp_accounts_userstatus (`Username`, `PluginUserEmail`, `TimeStamp`, `ClientDCids`) VALUES ('" & loopedUsername & "','" & myEmail & "', NOW(), '" & myComputerID & "') ON DUPLICATE KEY UPDATE `TimeStamp` = NOW(), `ClientDCids` = '" & myComputerID & "', `PluginUserEmail` = '" & myEmail & "'")
            objLocParams(2) = loopedUsername
            ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf UserManagement.UserDelete), objLocParams)
        Next
        Threading.Thread.Sleep(180000)
        Dim cmdIDhardinfo As Integer = mobjhost.SendCommand(myComputerID, 17, myComputerID) ' SendHardwareInfo (users)
        Do While CInt(mobjhost.GetSQL("Select Status from commands where cmdid=" & cmdIDhardinfo)) < 3
            Threading.Thread.Sleep(5000)
        Loop
        Dim cmdIDsysinfo As Integer = mobjhost.SendCommand(myComputerID, 123, myComputerID) ' SendSystemInfo
    End Sub


    Private Sub btnTest_Click(sender As System.Object, e As System.EventArgs)
        Cursor.Current = Cursors.WaitCursor
        Cursor.Current = Cursors.Default
    End Sub


    Private Sub btnLocalAccountsFlag_Click(sender As System.Object, e As System.EventArgs) Handles btnLocalAccountsFlag.Click
        mobjhost.SetSQL("UPDATE plugin_itsc_msp_accounts_settings SET `Local_Service_Account` = " & chkLocalAccountsFlag.Checked & ", `Local_Service_Account_Exclude` = " & chkLocalAccountsExclude.Checked)
        MsgBox("Settings Saved")
    End Sub


    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        MsgBox("Button1_Click")
        Cursor.Current = Cursors.WaitCursor
        ServiceManagement.LocalServiceLoader(mobjhost, "Add")
        Cursor.Current = Cursors.Default
    End Sub


    Private Sub btnSaveExclusions_Click(sender As System.Object, e As System.EventArgs) Handles btnSaveExclusions.Click
        Cursor.Current = Cursors.WaitCursor
        On Error GoTo errorHandler
        Dim strExcludeLocations As String = ""
        For Each thisRow As DataGridViewRow In DataGridViewLocations.Rows

            If thisRow.Cells.Item("Exclude").Value Then
                strExcludeLocations &= thisRow.Cells.Item("ID").Value & ","
            End If
        Next
        strExcludeLocations = strExcludeLocations.TrimEnd(",")
        mobjhost.SetSQL("UPDATE `plugin_itsc_msp_accounts_settings` SET `Exclude_Locations` = '" & strExcludeLocations & "'")
        MsgBox("Setting Saved")
        Exit Sub
errorHandler:
        Cursor.Current = Cursors.Default
        MessageBox.Show(Err.Description & " while saving exclusions.")
    End Sub


    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles btnTest.Click
        Dim CheckString As String = mobjhost.GetSQL("SHOW COLUMNS FROM plugin_itsc_msp_accounts_userstatus LIKE 'TimeStamp';")
        If CheckString IsNot Nothing AndAlso CheckString = "-9999" Then
            MessageBox.Show("inside:" & CheckString)
        Else
            MessageBox.Show("outside:" & CheckString)
        End If
    End Sub


    Private Sub btnServicePass_Click(sender As System.Object, e As System.EventArgs) Handles btnServicePass.Click
        On Error GoTo errorHandler
        If Not inputValidation(txtServicePass.Text) Then Exit Sub
        Dim strServiceName As String = mobjhost.GetSQL("SELECT Service_Account FROM `plugin_itsc_msp_accounts_settings`")
        Dim strServicePassword As String = txtServicePass.Text
        If dbSavePassword(strServiceName, strServicePassword) Then
            Cursor.Current = Cursors.WaitCursor
            mobjhost.SetSQL("UPDATE plugin_itsc_msp_accounts_users SET `AutoChangeDate` = CURDATE() WHERE `Username` = '" & strServiceName & "'")
            PasswordManagement.LoopedChangePassword(mobjhost, strServiceName, strServicePassword)
            Cursor.Current = Cursors.Default
            MessageBox.Show("Changes can take up to 15 minutes to propagate.")
        End If
        Exit Sub
errorHandler:
        Cursor.Current = Cursors.Default
        MessageBox.Show(Err.Description & " while changing password.")
    End Sub


    Private Sub ChangeAll_Click(sender As System.Object, e As System.EventArgs) Handles ChangeAll.Click
        If MessageBox.Show("Are you sure you want to change them all?", "Global Password Change", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Dim strServiceName As String = mobjhost.GetSQL("SELECT Service_Account FROM `plugin_itsc_msp_accounts_settings`")
            Dim ds As DataSet = mobjhost.GetDataSet("SELECT Username FROM plugin_itsc_msp_accounts_users WHERE AutoChangePassword IS TRUE")
            Cursor.Current = Cursors.WaitCursor
            For Each DataRow As DataRow In ds.Tables(0).Rows
                Dim User_Name As String = DataRow("Username").ToString
                Dim user_password As String = "Random"
                PasswordManagement.LoopedChangePassword(mobjhost, User_Name, user_password)
                If User_Name = strServiceName Then
                    If mobjhost.GetSQL("Select Local_Service_Account from plugin_itsc_msp_accounts_settings") Then
                        ServiceManagement.LocalServiceLoader(mobjhost, "Change")
                    End If
                End If
            Next
            Cursor.Current = Cursors.Default
        End If
    End Sub


End Class

