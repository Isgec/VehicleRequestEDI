Imports System.Reflection
Imports System.Xml
Imports System.Xml.Serialization
Imports System.Web.Script.Serialization
Imports System.Net
Imports System.Drawing
Imports System.IO
Public Class JobProcessor
  Inherits TimerSupport
  Implements IDisposable

  Private jpConfig As ConfigFile = Nothing
  Private MayContinue As Boolean = True
  Private LibraryPath As String = ""
  Private LibraryID As String = ""
  Private RemoteLibraryConnected As Boolean = False

  Public Event JobStarted()
  Public Event JobStopped()
  Public Event Log(ByVal Message As String)
  Public Event Err(ByVal Message As String)

  Public Sub Msg(ByVal str As String)
    RaiseEvent Log(str)
  End Sub
  Public Sub MsgErr(ByVal str As String)
    RaiseEvent Err(str)
  End Sub
  Public Overrides Sub Process()
    Try
      For Each cmp As ConfigFile.Company In jpConfig.Companies
        If Not cmp.IsActive Then Continue For
        Msg("ERP Company: " & cmp.ERPCompany)
        Dim VRs As List(Of SIS.VR.vrVehicleRequest) = SIS.VR.vrVehicleRequest.SelectList(cmp.ERPCompany)
        Msg("To be processed: " & VRs.Count)
        For Each vr As SIS.VR.vrVehicleRequest In VRs
          Dim spr As SPRequest = GetSPR(vr)
          SubmitRequest(spr, cmp.ERPCompany)
          If IsStopping Then
            Exit For
          End If
        Next
        If IsStopping Then
          Msg("Cancelled")
          Exit For
        End If
      Next
    Catch ex As Exception
      MsgErr(ex.Message)
    End Try
  End Sub
  Public Shared Function SubmitRequest(spr As SPRequest, Comp As String) As SPResponse
    Dim url As String = "https://demo.superprocure.com/sp/requisition/createRequisitionForISGEC"
    Dim xResponse As SPResponse = Nothing
    Dim xWebRequest As HttpWebRequest = GetWebRequest(url)
    Dim jsonStr As String = New JavaScriptSerializer().Serialize(spr)
    Dim jsonByte() As Byte = System.Text.Encoding.ASCII.GetBytes(jsonStr)
    xWebRequest.ContentLength = jsonByte.Count
    xWebRequest.GetRequestStream().Write(jsonByte, 0, jsonByte.Count)
    Try
      Dim rs As WebResponse = xWebRequest.GetResponse()
      Dim st As IO.Stream = rs.GetResponseStream
      Dim sr As IO.StreamReader = New IO.StreamReader(st)
      Dim strResponse As String = sr.ReadToEnd
      sr.Close()
      xResponse = New JavaScriptSerializer().Deserialize(strResponse, GetType(SPApi.SPResponse))
    Catch ex As Exception
      Throw New Exception(ex.Message)
    End Try
    Return xResponse
  End Function
  Private Shared Function GetWebRequest(url As String) As HttpWebRequest
    Dim Uri As Uri = New Uri(url)
    Dim xWebRequest As HttpWebRequest = WebRequest.Create(Uri)
    With xWebRequest
      .Method = "POST"
      .ContentType = "application/json"
      .Accept = "*/*"
      .CachePolicy = New Cache.RequestCachePolicy(Net.Cache.RequestCacheLevel.NoCacheNoStore)
      '.Headers.Add("mode", "no-cors")
      .Headers.Add("ContentType", "application/json")
      .Headers.Add("Authorization", "Basic MTk3OmE3Y2MxMDA4NWRkYTg0Y2VhODU1ZTlkOTRiZDdiNzUy")
      .KeepAlive = True
    End With
    Return xWebRequest
  End Function
  Private Shared Function GetSPR(r As SIS.VR.vrVehicleRequest) As SPApi.SPRequest
    Dim t As New SPApi.SPRequest
    With t
      .RFQ = r.RequestNo
      .loadingDate = Convert.ToDateTime(r.VehicleRequiredOn).ToString("dd-MM-yyyy")
      .projectCode = r.ProjectID
      .projectName = r.IDM_Projects4_Description
      .vehicleType = r.VR_VehicleTypes9_cmba
      .supplierCode = r.SupplierID
      .supplierName = "jjjjj"
      .supplierLoadingPoint = r.FromLocation & " " & r.FromPinCode
      .deliveryUnloadingPoint = r.ToLocation & " " & r.ToPinCode
      .loadingPoint = r.FromLocation
      .unloadingPoint = r.ToLocation
      .noOfVehicles = 1
      .product = r.ItemDescription
      .length = r.Length
      .breadth = r.Width
      .height = r.Height
      .notes = r.Remarks
      .branchCode = "EU200"
      .deliveryCode = "NONE"
      .materialWt = r.MaterialWeight
      .projectLocation = r.ToLocation
    End With
    With t
      .RFQ = "LG" & GetNextID()
      .loadingDate = "20-03-2020"
      .projectCode = "ABD18"
      .projectName = r.IDM_Projects4_Description
      .vehicleType = "Daala Body"
      .supplierCode = "ERTY"
      .supplierName = "JJJJJJ"
      .supplierLoadingPoint = r.FromLocation & " " & r.FromPinCode
      .deliveryUnloadingPoint = r.ToLocation & " " & r.ToPinCode
      .loadingPoint = "Kolkata"
      .unloadingPoint = "Dhanbad"
      .noOfVehicles = 1
      .product = "Fuel"
      .length = 20
      .breadth = 10
      .height = 10
      .notes = ""
      .branchCode = "drew"
      .deliveryCode = "SSFF"
      .materialWt = 2000
      .projectLocation = "howrah"
    End With
    Return t
  End Function

  Public Class SPRequest
    Public Property RFQ As String = "" '200
    Public Property loadingDate As String = "" '10 dd-MM-yyyy
    Public Property projectCode As String = "" '255
    Public Property projectName As String = "" '255
    Public Property vehicleType As String = "" '255
    Public Property supplierCode As String = "" '255
    Public Property supplierName As String = "" '255
    Public Property loadingPoint As String = "" '1000
    Public Property unloadingPoint As String = "" '1000
    Public Property supplierLoadingPoint As String = "" '255
    Public Property noOfVehicles As Integer = 0
    Public Property product As String = "" '255
    Public Property length As Integer = 0
    Public Property breadth As Integer = 0
    Public Property height As Integer = 0
    Public Property notes As String = "" '1000
    Public Property branchCode As String = "" '255
    Public Property deliveryCode As String = "" '255
    Public Property materialWt As Integer = 0
    Public Property deliveryUnloadingPoint As String = "" '100
    Public Property projectLocation As String = "" '255

  End Class
  Public Class SPResponse
    Public Property Message As String = ""
    Public Property Code As String = ""
    Public Property Status As String = ""
    Public Property ReqId As String = ""
    Public ReadOnly Property IsError As Boolean
      Get
        Return (Code <> "201")
      End Get
    End Property
  End Class
  Public Overrides Sub Started()
    Try
      RaiseEvent JobStarted()
      Msg("Reading Settings")
      Dim ConfigPath As String = IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) & "\Settings.xml"
      jpConfig = ConfigFile.DeSerialize(ConfigPath)
      SIS.SYS.SQLDatabase.DBCommon.BaaNLive = jpConfig.BaaNLive
    Catch ex As Exception
      StopJob()
      MsgErr(ex.Message)
    End Try
  End Sub

  Public Overrides Sub Stopped()
    If RemoteLibraryConnected Then
      ConnectToNetworkFunctions.disconnectFromNetwork("X:")
      RemoteLibraryConnected = False
    End If
    jpConfig = Nothing
    RaiseEvent JobStopped()
    Msg("Stopped")
  End Sub


  Sub New()
    'dummy
  End Sub

#Region "IDisposable Support"
  Private disposedValue As Boolean ' To detect redundant calls

  ' IDisposable
  Protected Overridable Sub Dispose(disposing As Boolean)
    If Not disposedValue Then
      If disposing Then
        ' TODO: dispose managed state (managed objects).
      End If

      ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
      ' TODO: set large fields to null.
    End If
    disposedValue = True
  End Sub

  ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
  'Protected Overrides Sub Finalize()
  '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
  '    Dispose(False)
  '    MyBase.Finalize()
  'End Sub

  ' This code added by Visual Basic to correctly implement the disposable pattern.
  Public Sub Dispose() Implements IDisposable.Dispose
    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    Dispose(True)
    ' TODO: uncomment the following line if Finalize() is overridden above.
    ' GC.SuppressFinalize(Me)
  End Sub
#End Region
End Class
