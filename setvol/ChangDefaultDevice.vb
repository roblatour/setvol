
Imports System.Runtime.InteropServices
Imports NAudio.CoreAudioApi

Module ChangDefaultDevice

    Friend Sub SetDefaultDevice(ByVal DeviceID As String, ByVal CommunicationsDevice As Boolean)

        AudioEndPoints.SetDefaultEndPoint(DeviceID, CommunicationsDevice)

    End Sub

End Module
Public Class AudioEndPoints

    <ComImport(), Guid("568b9108-44bf-40b4-9006-86afe5b5a620"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface IPolicyConfigVista
        Function GetMixFormat(<MarshalAs(UnmanagedType.LPWStr)> ByVal wszDeviceId As String, ByRef param1 As IntPtr) As Integer 'not available on Windows 7, use method from IPolicyConfig
        Function GetDeviceFormat(<MarshalAs(UnmanagedType.LPWStr)> ByVal wszDeviceId As String, ByVal param1 As Integer, ByRef wfx As tWAVEFORMATEX) As Integer
        Function SetDeviceFormat(<MarshalAs(UnmanagedType.LPWStr)> ByVal wszDeviceId As String, ByRef wfx1 As tWAVEFORMATEX, ByRef wfx2 As tWAVEFORMATEX) As Integer
        Function GetProcessingPeriod(<MarshalAs(UnmanagedType.LPWStr)> ByVal wszDeviceId As String, ByVal param1 As Integer, ByRef param2 As Long, ByRef param3 As Long) As Integer 'not available on Windows 7, use method from IPolicyConfig
        Function SetProcessingPeriod(<MarshalAs(UnmanagedType.LPWStr)> ByVal wszDeviceId As String, ByRef param1 As Long) As Integer 'not available on Windows 7, use method from IPolicyConfig
        Function GetShareMode(<MarshalAs(UnmanagedType.LPWStr)> ByVal wszDeviceId As String, ByRef DeviceShareMode As IntPtr) As Integer 'not available on Windows 7, use method from IPolicyConfig
        Function SetShareMode(<MarshalAs(UnmanagedType.LPWStr)> ByVal wszDeviceId As String, ByRef DeviceShareMode As IntPtr) As Integer 'not available on Windows 7, use method from IPolicyConfig
        Function GetPropertyValue(<MarshalAs(UnmanagedType.LPWStr)> ByVal wszDeviceId As String, ByRef propKey As PropertyKey, ByRef propVar As PropVariant) As Integer
        Function SetPropertyValue(<MarshalAs(UnmanagedType.LPWStr)> ByVal wszDeviceId As String, ByRef propKey As PropertyKey, ByRef propVar As PropVariant) As Integer
        Function SetDefaultEndpoint(<MarshalAs(UnmanagedType.LPWStr)> ByVal wszDeviceId As String, ByVal eRole As ERole) As Integer
        Function SetEndpointVisibility(<MarshalAs(UnmanagedType.LPWStr)> ByVal wszDeviceId As String, ByVal param1 As Integer) As Integer 'not available on Windows 7, use method from IPolicyConfig

    End Interface

    Friend Shared Sub SetDefaultEndPoint(devId As String, SetAsDefaultCommunicationDevice As Boolean)

        Dim IPCV As IPolicyConfigVista = Nothing
        Dim oIPCV As Object = Nothing

        Try

            Dim IID_IPolicyConfigVista As Guid = New Guid("568b9108-44bf-40b4-9006-86afe5b5a620")
            Dim CLSID_CPolicyConfigVistaClient As Guid = New Guid("294935CE-F637-4E7C-A41B-AB255460B862")

            If CoCreateInstance(CLSID_CPolicyConfigVistaClient, Nothing, CLSCTX.CLSCTX_INPROC_SERVER, IID_IPolicyConfigVista, oIPCV) <> 0 Then
                Throw New Exception("Failed:  CoCreateInstance")
            Else
                IPCV = CType(oIPCV, IPolicyConfigVista)
                If IPCV Is Nothing Then Throw New Exception("Failed: COM cast to IPolicyConfigVista")
            End If

            If SetAsDefaultCommunicationDevice Then
                IPCV.SetDefaultEndpoint(devId, ERole.eCommunications)
            Else
                IPCV.SetDefaultEndpoint(devId, ERole.eConsole)
            End If

        Finally

            If (IPCV IsNot Nothing) Then
                While Marshal.ReleaseComObject(IPCV) > 0
                End While
                IPCV = Nothing
            End If

            If (oIPCV IsNot Nothing) Then
                While Marshal.ReleaseComObject(oIPCV) > 0
                End While
                oIPCV = Nothing
            End If

        End Try

    End Sub

    Private Shared SPDRP_Guid As New Guid(&HA45C254EUI, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0)

    Private Const STGM_READ As Integer = &H0 'used for the (stgmAccess) parameter of the (OpenPropertyStore) function in the (IMMDevice) interface.

    <ComImport(), Guid("0BD7A1BE-7A1A-44DB-8397-CC5392387B5E"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface IMMDeviceCollection
        Function GetCount(ByRef pcDevices As UInteger) As Integer
        Function Item(ByVal nDevice As UInteger, ByRef ppDevice As IntPtr) As Integer
    End Interface

    <ComImport(), Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface IMMDeviceEnumerator
        Function EnumAudioEndpoints(<ComAliasName("MMDeviceAPILib.EDataFlow")> ByVal dataFlow As EDataFlow, ByVal dwStateMask As Integer, ByRef ppDevices As IntPtr) As Integer
        Function GetDefaultAudioEndpoint(<ComAliasName("MMDeviceAPILib.EDataFlow")> ByVal dataFlow As EDataFlow, <ComAliasName("MMDeviceAPILib.ERole")> ByVal role As ERole, ByRef ppEndpoint As IntPtr) As Boolean
        Function GetDevice(ByVal pwstrId As String, ByRef ppDevice As IntPtr) As Integer
        Function RegisterEndpointNotificationCallback(ByVal pClient As IMMNotificationClient) As Integer
        Function UnregisterEndpointNotificationCallback(ByVal pClient As IMMNotificationClient) As Integer
    End Interface

    <ComImport(), Guid("7991EEC9-7E89-4D85-8390-6C703CEC60C0"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface IMMNotificationClient
        Sub OnDeviceStateChanged(<MarshalAs(UnmanagedType.LPWStr)> ByVal pwstrDeviceId As String, ByVal dwNewState As UInteger)
        Sub OnDeviceAdded(<MarshalAs(UnmanagedType.LPWStr)> ByVal pwstrDeviceId As String)
        Sub OnDeviceRemoved(<MarshalAs(UnmanagedType.LPWStr)> ByVal pwstrDeviceId As String)
        Sub OnDefaultDeviceChanged(<ComAliasName("MMDeviceAPILib.EDataFlow")> ByVal flow As EDataFlow, <ComAliasName("MMDeviceAPILib.ERole")> ByVal role As ERole, <MarshalAs(UnmanagedType.LPWStr)> ByVal pwstrDefaultDeviceId As String)
        Sub OnPropertyValueChanged(<MarshalAs(UnmanagedType.LPWStr)> ByVal pwstrDeviceId As String, ByVal key As PropertyKey)
    End Interface

    <ComImport(), Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface IMMDevice
        Function Activate(ByRef iid As Guid, ByVal dwClsCtx As Integer, ByVal pActivationParams As IntPtr, ByRef ppInterface As IntPtr) As Boolean
        Function OpenPropertyStore(ByVal stgmAccess As Integer, ByRef ppProperties As IntPtr) As Integer
        Function GetId(<MarshalAs(UnmanagedType.LPWStr)> ByRef ppstrId As String) As Integer
        Function GetState(ByRef pdwState As IMMDeviceStatesFlags) As Integer
    End Interface

    <ComImport(), Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface IPropertyStore
        Function GetCount(ByRef cProps As UInteger) As UInteger
        Function GetAt(ByVal iProp As UInteger, ByRef pkey As PropertyKey) As UInteger
        Function GetValue(ByRef key As PropertyKey, ByVal pv As PropVariant) As UInteger
        Function SetValue(ByRef key As PropertyKey, ByVal pv As PropVariant) As UInteger
        Function Commit() As UInteger
    End Interface

    <Flags()>
    Private Enum IMMDeviceStatesFlags As Integer
        DEVICE_STATE_ACTIVE = &H1
        DEVICE_STATE_DISABLED = &H2
        DEVICE_STATE_NOTPRESENT = &H4
        DEVICE_STATE_UNPLUGGED = &H8
        DEVICE_STATE_ALL = &HF
    End Enum

    Public Enum EDataFlow As Integer
        eRender = &H0
        eCapture = &H1
        eAll = &H2
        EDataFlow_enum_count = &H3
    End Enum

    Private Enum ERole As Integer
        eConsole = &H0
        eMultimedia = &H1
        eCommunications = &H2
        ERole_enum_count = &H3
    End Enum

    <Flags()>
    Private Enum CLSCTX
        CLSCTX_INPROC_SERVER = &H1
        CLSCTX_INPROC_HANDLER = &H2
        CLSCTX_LOCAL_SERVER = &H4
        CLSCTX_INPROC_SERVER16 = &H8
        CLSCTX_REMOTE_SERVER = &H10
        CLSCTX_INPROC_HANDLER16 = &H20
        CLSCTX_RESERVED1 = &H40
        CLSCTX_RESERVED2 = &H80
        CLSCTX_RESERVED3 = &H100
        CLSCTX_RESERVED4 = &H200
        CLSCTX_NO_CODE_DOWNLOAD = &H400
        CLSCTX_RESERVED5 = &H800
        CLSCTX_NO_CUSTOM_MARSHAL = &H1000
        CLSCTX_ENABLE_CODE_DOWNLOAD = &H2000
        CLSCTX_NO_FAILURE_LOG = &H4000
        CLSCTX_DISABLE_AAA = &H8000
        CLSCTX_ENABLE_AAA = &H10000
        CLSCTX_FROM_DEFAULT_CONTEXT = &H20000
        CLSCTX_INPROC = CLSCTX_INPROC_SERVER Or CLSCTX_INPROC_HANDLER
        CLSCTX_SERVER = CLSCTX_INPROC_SERVER Or CLSCTX_LOCAL_SERVER Or CLSCTX_REMOTE_SERVER
        CLSCTX_ALL = CLSCTX_SERVER Or CLSCTX_INPROC_HANDLER
    End Enum

    <StructLayout(LayoutKind.Sequential, Pack:=4)>
    Private Structure PropertyKey
        Private m_formatId As Guid
        Private m_propertyId As Integer
        Public ReadOnly Property FormatId() As Guid
            Get
                Return Me.m_formatId
            End Get
        End Property
        Public ReadOnly Property PropertyId() As Integer
            Get
                Return Me.m_propertyId
            End Get
        End Property
        Public Sub New(ByVal formatId As Guid, ByVal propertyId As Integer)
            Me.m_formatId = formatId
            Me.m_propertyId = propertyId
        End Sub
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure tWAVEFORMATEX
        Public wFormatTag As UShort
        Public nChannels As UShort
        Public nSamplesPerSec As UInteger
        Public nAvgBytesPerSec As UInteger
        Public nBlockAlign As UShort
        Public wBitsPerSample As UShort
        Public cbSize As UShort
    End Structure

    <StructLayout(LayoutKind.Explicit)>
    Private NotInheritable Class PropVariant
        Implements IDisposable

        <FieldOffset(0)> Private valueType As UShort
        <FieldOffset(8)> Private ptr As IntPtr

        Public Property VarType() As VarEnum
            Get
                Return DirectCast(CInt(Me.valueType), VarEnum)
            End Get
            Set(ByVal value As VarEnum)
                Me.valueType = CUShort(value)
            End Set
        End Property

        Public ReadOnly Property IsNullOrEmpty() As Boolean
            Get
                Return (Me.valueType = CUShort(VarEnum.VT_EMPTY) OrElse Me.valueType = CUShort(VarEnum.VT_NULL))
            End Get
        End Property

        Public ReadOnly Property Value() As String
            Get
                Return Marshal.PtrToStringUni(Me.ptr)
            End Get
        End Property

        Public Sub New()
        End Sub

        Public Sub New(ByVal value As String)
            If value Is Nothing Then
                Throw New ArgumentNullException("value")
            End If

            Me.valueType = CUShort(VarEnum.VT_LPWSTR)
            Me.ptr = Marshal.StringToCoTaskMemUni(value)
        End Sub

        Protected Overrides Sub Finalize()
            Try
                Dispose()
            Finally
                MyBase.Finalize()
            End Try
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            PropVariantClear(Me)
            GC.SuppressFinalize(Me)
        End Sub
    End Class

    <DllImport("Ole32.dll", PreserveSig:=False)> Private Shared Sub PropVariantClear(<[In](), Out()> ByVal pvar As PropVariant)
    End Sub

    <DllImport("ole32.Dll")>
    Private Shared Function CoCreateInstance(ByRef clsid As Guid, <MarshalAs(UnmanagedType.IUnknown)> ByVal inner As Object, ByVal context As Integer, ByRef uuid As Guid, <MarshalAs(UnmanagedType.IUnknown)> ByRef rReturnedComObject As Object) As Integer
    End Function

End Class

Public Class EndPointDeviceInfo
    Public ReadOnly Property FriendlyName As String
    Public ReadOnly Property Description As String
    Public ReadOnly Property Id As String

    Public Sub New(devName As String, devId As String, devDescription As String)
        _FriendlyName = devName
        _Id = devId
        _Description = devDescription
    End Sub

    Public Overrides Function ToString() As String
        Return FriendlyName
    End Function

End Class


