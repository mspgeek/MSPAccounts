Public Class clsPermissions
    Implements LabTech.Interfaces.IPermissions
    'This class is loaded by the DBAgent. There will only ever be one instance of this class.

    Private objHost As LabTech.Interfaces.IControlCenter

    
    Private Sub DoInitialSetup()
        Try
            'Create the Tables if they do not exist
            'Check if the System is setup by checking for the existance of the table
            Dim sql As String
            Dim CheckString As String
            CheckString = objHost.GetSQL("SHOW TABLES LIKE 'plugin_itsc_msp_accounts_settings'")
            If CheckString IsNot Nothing AndAlso CheckString = "-9999" Then
                sql = "CREATE TABLE `plugin_itsc_msp_accounts_settings` (`MSP_Name` varchar(50) NOT NULL, `User_Prefix` varchar(50) NOT NULL, `Exclude_Locations` varchar(2000) NOT NULL, `Service_Account` varchar(50) DEFAULT NULL, `Min_Password_Length` tinyint(2) NOT NULL, `Password_Change_Days` tinyint(3) NOT NULL, `Local_Service_Account` BINARY(1) DEFAULT 0, `Local_Service_Account_Exclude` BINARY(1) DEFAULT 1, `Min_Password_Upper` tinyint(2) NOT NULL, `Min_Password_Lower` tinyint(2) NOT NULL, `Min_Password_Number` tinyint(2) NOT NULL, `Min_Password_Special` tinyint(2) NOT NULL, PRIMARY KEY (`MSP_Name`)) ENGINE=InnoDB DEFAULT CHARSET=utf8"
                objHost.SetSQL(sql)
                sql = "INSERT INTO plugin_itsc_msp_accounts_settings (`MSP_Name`, `User_Prefix`, `Exclude_Locations`, `Service_Account`, `Min_Password_Length`, `Password_Change_Days`, `Min_Password_Upper`, `Min_Password_Lower`, `Min_Password_Number`, `Min_Password_Special`) VALUES ('Managed_Service_Provider', 'MSP_', '0,1', 'None', 14, 59, 2, 2, 2, 2)"
                objHost.SetSQL(sql)
            End If
            CheckString = objHost.GetSQL("SHOW TABLES LIKE 'plugin_itsc_msp_accounts_users'")
            If CheckString IsNot Nothing AndAlso CheckString = "-9999" Then
                sql = "CREATE TABLE plugin_itsc_msp_accounts_users (`Username` varchar(50) NOT NULL, `Password` blob NOT NULL, `AutoChangePassword` tinyint(1) NOT NULL, `AutoChangeDate` date NOT NULL, PRIMARY KEY (`Username`)) ENGINE=InnoDB DEFAULT CHARSET=utf8"
                objHost.SetSQL(sql)
            End If
            CheckString = objHost.GetSQL("SHOW TABLES LIKE 'plugin_itsc_msp_accounts_userstatus'")
            If CheckString IsNot Nothing AndAlso CheckString = "-9999" Then
                sql = "CREATE TABLE plugin_itsc_msp_accounts_userstatus (`Username` varchar(50) NOT NULL, `PluginUserEmail` varchar(50) DEFAULT NULL, `TimeStamp` DATETIME DEFAULT NULL, `ClientDCids` varchar(5000) DEFAULT NULL, PRIMARY KEY (`Username`)) ENGINE=InnoDB DEFAULT CHARSET=utf8"
                objHost.SetSQL(sql)
                sql = "INSERT INTO plugin_itsc_msp_accounts_userstatus (`Username`) VALUES ('debug')"
                objHost.SetSQL(sql)
            End If
            ' repair previous to 2.160531
            CheckString = objHost.GetSQL("SELECT character_maximum_length FROM information_schema.columns WHERE table_name = 'plugin_itsc_msp_accounts_settings' AND column_name = 'Exclude_Locations';")
            If CheckString < 2000 Then
                sql = "ALTER TABLE plugin_itsc_msp_accounts_settings MODIFY Exclude_Locations VARCHAR(2000)"
                objHost.SetSQL(sql)
            End If
            ' modify previous to 2.171129
            CheckString = objHost.GetSQL("SHOW COLUMNS FROM plugin_itsc_msp_accounts_settings LIKE 'Min_Password_Upper';")
            If CheckString IsNot Nothing AndAlso CheckString = "-9999" Then
                sql = "ALTER TABLE plugin_itsc_msp_accounts_settings ADD COLUMN `Min_Password_Upper` tinyint(2) NOT NULL, ADD COLUMN `Min_Password_Lower` tinyint(2) NOT NULL, ADD COLUMN `Min_Password_Number` tinyint(2) NOT NULL, ADD COLUMN `Min_Password_Special` tinyint(2) NOT NULL;"
                objHost.SetSQL(sql)
                sql = "UPDATE plugin_itsc_msp_accounts_settings SET `Min_Password_Upper` = 2, `Min_Password_Lower` = 2, `Min_Password_Number` = 2, `Min_Password_Special` = 2;"
                objHost.SetSQL(sql)
            End If
            CheckString = objHost.GetSQL("SHOW COLUMNS FROM plugin_itsc_msp_accounts_userstatus LIKE 'TimeStamp';")
            If CheckString IsNot Nothing AndAlso CheckString = "-9999" Then
                sql = "ALTER TABLE plugin_itsc_msp_accounts_userstatus ADD COLUMN `TimeStamp` DATETIME DEFAULT NULL AFTER `PluginUserEmail`;"
                objHost.SetSQL(sql)
            End If
        Catch ex As Exception
            objHost.LogMessage("DoInitialSetup Error:" & ex.Message)
        End Try
    End Sub



    Public Function GetPermissionSet(ByVal UserID As Integer, ByVal IsSuperAdmin As Boolean, ByVal UserClasses As String) As System.Collections.Hashtable Implements LabTech.Interfaces.IPermissions.GetPermissionSet
        Dim ht As New Hashtable
        Try
            'This will be called for every user on the system. The USer Classes is a Comma seperated string of userclasses the user is a member of
            'You can use the UserID to determin what access to grant to the user.
            'Super admins typically have unlimited access to the system

            'Add all Tables you have added to the database to this hash table so the users can have access.

            ht.Add("plugin_itsc_msp_accounts_settings", "SELECT,INSERT,UPDATE,DELETE")
            ht.Add("plugin_itsc_msp_accounts_users", "SELECT,INSERT,UPDATE,DELETE")
            ht.Add("plugin_itsc_msp_accounts_userstatus", "SELECT,INSERT,UPDATE,DELETE,ALTER")

        Catch ex As Exception
            objHost.LogMessage("GetPermissionSet Error:" & ex.Message)
        End Try
        Return ht
    End Function



    Public Sub Initialize(ByVal Host As LabTech.Interfaces.IControlCenter) Implements LabTech.Interfaces.IPermissions.Initialize
        objHost = Host
        DoInitialSetup()
    End Sub


    Public Sub Decommision() Implements LabTech.Interfaces.IPermissions.Decommision
        objHost = Nothing
    End Sub

    Public ReadOnly Property Name() As String Implements LabTech.Interfaces.IPermissions.Name
        Get
            Return PluginName & " Permissions v" & mVersion
        End Get
    End Property

End Class