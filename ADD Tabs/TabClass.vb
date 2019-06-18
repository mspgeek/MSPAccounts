Imports system.Windows.Forms
Imports LabTech.Interfaces
Imports System.Reflection
Imports System.IO
Imports System.xml
Imports System.Xml.Serialization


Public Class TabClass
    Implements LabTech.Interfaces.ITabs
    'This Class is created by the Control Center. It is only created ONE time.
    ' BUT some items can have multipule. For every computer screen we will call the ComputerINIT for EACH ONE.

    Private objHost As LabTech.Interfaces.IControlCenter

    Public Sub Initialize(ByVal Host As LabTech.Interfaces.IControlCenter) Implements LabTech.Interfaces.ITabs.Initialize
        objHost = Host
        'Load Properties you might need to configure stuff
    End Sub


    Public Sub Decommision() Implements LabTech.Interfaces.ITabs.Decommision
        objHost = Nothing
    End Sub

    Public ReadOnly Property Name() As String Implements LabTech.Interfaces.ITabs.Name
        Get
            Return PluginName & " Config v" & mVersion
        End Get
    End Property


    'The Dashboard Intergrations Tab. NO ID, only ONE instance of this one.
    Public Function ConfigInit() As System.Windows.Forms.TabPage Implements LabTech.Interfaces.ITabs.ConfigInit
        Return Nothing
    End Function

    Public Sub ConfigClose() Implements LabTech.Interfaces.ITabs.ConfigClose
    End Sub


    'The Client Screen, the ID is the ClientID, Multi instance possible
    Public Function ClientInit(ByVal ID As Integer) As System.Windows.Forms.TabPage Implements LabTech.Interfaces.ITabs.ClientInit
        Return Nothing
    End Function

    Public Sub ClientClose(ByVal ID As Integer) Implements LabTech.Interfaces.ITabs.ClientClose

    End Sub

    'The Location Screen, The ID is the Location ID, Multi instance possible
    Public Function LocationInit(ByVal ID As Integer) As System.Windows.Forms.TabPage Implements LabTech.Interfaces.ITabs.LocationInit
        Return Nothing
    End Function

    Public Sub LocationClose(ByVal ID As Integer) Implements LabTech.Interfaces.ITabs.LocationClose

    End Sub

    'The computer Screen, The ID is the ComputerID, Multi instance possible
    Public Function ComputerInit(ByVal ID As Integer) As System.Windows.Forms.TabPage Implements LabTech.Interfaces.ITabs.ComputerInit
        Return Nothing
    End Function

    Public Sub ComputerClose(ByVal ID As Integer) Implements LabTech.Interfaces.ITabs.ComputerClose

    End Sub

    'The NEtwork Device Screen, Multi instance possible
    Public Function DeviceInit(ByVal ID As Integer) As System.Windows.Forms.TabPage Implements LabTech.Interfaces.ITabs.DeviceInit
        'If you do not want a TAB then return NOTHING
        Return Nothing
    End Function

    Public Sub DeviceClose(ByVal ID As Integer) Implements LabTech.Interfaces.ITabs.DeviceClose

    End Sub

    'The Group Screen, Multi instance possible
    Public Function GroupInit(ByVal ID As Integer) As System.Windows.Forms.TabPage Implements LabTech.Interfaces.ITabs.GroupInit
        'If you do not want a TAB then return NOTHING
        Return Nothing
    End Function

    Public Sub GroupClose(ByVal ID As Integer) Implements LabTech.Interfaces.ITabs.GroupClose

    End Sub

    'The Ticket Screen, Multi instance possible
    Public Function TicketInit(ByVal ID As Integer) As System.Windows.Forms.TabPage Implements LabTech.Interfaces.ITabs.TicketInit

        'If you do not want a TAB then return NOTHING
        Return Nothing
    End Function

    Public Sub TicketClose(ByVal ID As Integer) Implements LabTech.Interfaces.ITabs.TicketClose

    End Sub

    'The Search Screen, Only One Instance
    Public Function SearchInit() As System.Windows.Forms.TabPage Implements LabTech.Interfaces.ITabs.SearchInit

        'If you do not want a TAB then return NOTHING
        Return Nothing
    End Function

    Public Sub SearchClose() Implements LabTech.Interfaces.ITabs.SearchClose

    End Sub

    'The Alerts Screen, Only One Instance
    Public Function AlertsInit() As System.Windows.Forms.TabPage Implements LabTech.Interfaces.ITabs.AlertsInit

        'If you do not want a TAB then return NOTHING
        Return Nothing
    End Function

    Public Sub AlertsClose() Implements LabTech.Interfaces.ITabs.AlertsClose

    End Sub

    'The Monitors Screen, Only one Instance
    Public Function MonitorsInit() As System.Windows.Forms.TabPage Implements LabTech.Interfaces.ITabs.MonitorsInit
        'If you do not want a TAB then return NOTHING
        Return Nothing
    End Function

    Public Sub MonitorsClose() Implements LabTech.Interfaces.ITabs.MonitorsClose

    End Sub


End Class
