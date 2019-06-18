Public Class PluginMain
    Implements LabTech.Interfaces.IPlugin
    Dim mMD5 As String
    Dim mFilename As String



    Public ReadOnly Property Name() As String Implements LabTech.Interfaces.IPlugin.Name
        Get
            Return PluginName
        End Get
    End Property

    Public ReadOnly Property About() As String Implements LabTech.Interfaces.IPlugin.About
        Get
            Return "LabTech " & PluginName & " Plugin" & vbCrLf & "       v" & Version.ToString & vbCrLf & "Copyright 2015 " & Author
        End Get
    End Property

    Public ReadOnly Property Author() As String Implements LabTech.Interfaces.IPlugin.Author
        Get
            Return mAuthor
        End Get
    End Property


    Public ReadOnly Property Version() As Integer Implements LabTech.Interfaces.IPlugin.Version
        Get
            Return mVersion
        End Get
    End Property

    Public Function IsCompatible(ByVal Objhost As LabTech.Interfaces.IControlCenter) As Boolean Implements LabTech.Interfaces.IPlugin.IsCompatible
        Return True
    End Function

    Public Function IsLicensed() As Boolean Implements LabTech.Interfaces.IPlugin.IsLicensed
        Return True
    End Function

    Public Function IsLicensed(ByVal Objhost As LabTech.Interfaces.IControlCenter) As Boolean Implements LabTech.Interfaces.IPlugin.IsLicensed
        Return True
    End Function

    Public Function Remove(ByVal Objhost As LabTech.Interfaces.IControlCenter) As Boolean Implements LabTech.Interfaces.IPlugin.Remove
        Return True
    End Function

    Public Function Install(ByVal Objhost As LabTech.Interfaces.IControlCenter) As Boolean Implements LabTech.Interfaces.IPlugin.Install
        Return True
    End Function

    Public Property Filename() As String Implements LabTech.Interfaces.IPlugin.Filename
        Get
            Return mFilename
        End Get
        Set(ByVal value As String)
            mFilename = value
        End Set
    End Property

    Public Property hMD5() As String Implements LabTech.Interfaces.IPlugin.hMD5
        Get
            Return mMD5
        End Get
        Set(ByVal value As String)
            mMD5 = value
        End Set
    End Property
End Class
