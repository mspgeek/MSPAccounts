Imports System.Windows.Forms

Public Class clsMenus
    Implements LabTech.Interfaces.IMenu
    'This Class is created in the control center and is only ever run ONE time.


    Private objHost As LabTech.Interfaces.IControlCenter

    Dim F As MSP_Accounts_Form 'Used to maintain a reference to the created form.

    Public Function CreateMainMenu() As System.Windows.Forms.MenuItem() Implements LabTech.Interfaces.IMenu.CreateMainMenu
        '   add button to toolbar
        'Try
        'Dim toolstrip1 As System.Windows.Forms.ToolStrip = DirectCast(Me.objHost.FormMain().Controls.Item("tsMain"), System.Windows.Forms.ToolStrip)
        'toolstrip1.Items.Add("MSP Accounts", My.Resources.msp_button, New EventHandler(AddressOf ShowApprovalForm))
        'toolstrip1.Items.Item((toolstrip1.Items.Count - 1)).ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.SizeToFit
        'toolstrip1 = Nothing
        'Catch exception1 As Exception
        'End Try

        Dim MNUs(0) As System.Windows.Forms.MenuItem
        Dim m As New System.Windows.Forms.MenuItem("MSP Accounts", AddressOf ShowApprovalForm)
        MNUs(0) = m
        Return MNUs
        'IF you are not wanting menus in this item then return NOTHING
        'Return Nothing
    End Function

    Public Function CreateToolsMenu() As System.Windows.Forms.MenuItem() Implements LabTech.Interfaces.IMenu.CreateToolsMenu
        'IF you are not wanting menus in this item then return NOTHING
        Return Nothing
    End Function

    Sub ShowApprovalForm(sender As Object, e As EventArgs)
        If F IsNot Nothing AndAlso Not F.IsDisposed Then
            F.BringToFront()
        Else
            F = New MSP_Accounts_Form(objHost)
            F.Text = "MSP Accounts"
            F.Show()
        End If
    End Sub

    Public Function CreateViewMenu() As System.Windows.Forms.MenuItem() Implements LabTech.Interfaces.IMenu.CreateViewMenu
        'you can have a permissions check here to see if they are to see this menu.
        'You can return Multipule menus in the Array
        'Dim MNUs(0) As System.Windows.Forms.MenuItem
        'Dim m As New System.Windows.Forms.MenuItem("MSP Accounts", AddressOf ShowApprovalForm)
        'MNUs(0) = m
        'Return MNUs
        Return Nothing
    End Function

    Public Sub Decommision() Implements LabTech.Interfaces.IMenu.Decommision
        'See we dispose of the form on closing of the control center
        If F IsNot Nothing AndAlso Not F.IsDisposed Then
            'F.Close()
            F.Dispose()
        End If
        objHost = Nothing
    End Sub

    Public Sub Initialize(Host As LabTech.Interfaces.IControlCenter) Implements LabTech.Interfaces.IMenu.Initialize
        objHost = Host
    End Sub

    Public ReadOnly Property Name As String Implements LabTech.Interfaces.IMenu.Name
        Get
            Return PluginName & " Menu v" & mVersion
        End Get
    End Property

    
End Class
