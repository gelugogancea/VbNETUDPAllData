Option Explicit On
Option Compare Text
Imports System.IO
Imports System.Net.Sockets.Socket
Imports System
Imports System.Text
Imports System.Net
Imports System.Net.Sockets
Imports Microsoft.VisualBasic
Imports System.Runtime.InteropServices
Public Class Form1
    Public Shared UdpInit As UdpClient = Nothing
    Public Shared UdpInProccess As UdpClient = Nothing

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim sParam(3) As String, i As Integer = 0
        sParam(0) = "7236742100"
        sParam(1) = "131005230601"
        sParam(2) = "7236742200"
        sParam(3) = "131005430601"
        For i = 1 To 65535
            Send2Server(i, "ARCTIC_FSLine_MPL402", "ARCTIC_FSLine_BeforeFoamingClient", "usUnitStationCV023", sParam)
            Threading.Thread.Sleep(1000)
        Next i
    End Sub
    Private Sub Send2Server(ByVal iMsgID As Integer, ByVal sFrom As String, ByVal sTo As String, ByVal sObjectName As String, ByVal sParams() As String)
        Dim MyCallServer As New CallServer, MyNewUDPPort As Integer = 0, iRetThrdProc As Integer = 65535
        Dim sMyBuff(1) As Byte, bMyHeader(1) As Byte

        MyNewUDPPort = MyCallServer.Request4Process("10.40.3.102", 10101, 10000)
        Select Case MyNewUDPPort
            Case -1
                Debug.Print("Error On Send ")
                ListBox1.Items.Add("Error On Send ")
                Exit Sub
            Case -2
                Debug.Print("Error On Receive - Server is not ready ")
                ListBox1.Items.Add("Error On Receive - Server is not ready ")
                Exit Sub
            Case Else
                ListBox1.Items.Add("New UDP Port Is : " & MyNewUDPPort)
                Debug.Print("New UDP Port Is : " & MyNewUDPPort)
        End Select

        Dim MyThrdServerProc As New ServerProc, udClientProc As UdpClient = Nothing, RemoteIpEndPoint As IPEndPoint = Nothing
        Dim iAnswerServer As Integer = 65535
        RemoteIpEndPoint = New System.Net.IPEndPoint(System.Net.IPAddress.Any, 0)
        udClientProc = New UdpClient()
        udClientProc.Connect(IPAddress.Parse("10.40.3.102"), MyNewUDPPort)
        iRetThrdProc = MyThrdServerProc.WaitThreadFromServer(10000, udClientProc, RemoteIpEndPoint)
        Select Case iRetThrdProc
            Case -1
                Debug.Print("Error in communication")
                ListBox1.Items.Add("Wait Thread Error in communication")
            Case Is > 0
                ListBox1.Items.Add("Wait Thread Time Out : " & iRetThrdProc)
                Debug.Print("Time Out : " & iRetThrdProc)
            Case Else
                ListBox1.Items.Add("Wait Thread ok")
                Debug.Print("Ok")
        End Select

        sMyBuff(0) = 0    'to have a reference and not return NULL 

        MyThrdServerProc.EnCodeMsgShort(sFrom, sTo, sObjectName, sParams, sMyBuff)
        bMyHeader(0) = 0
        MyThrdServerProc.EnCodeHeader(iMsgID, ClientStruct.MSG_TYPE.OnDataReceive, 65535, sParams.Length, sMyBuff.Length, bMyHeader)
        Application.DoEvents()     'Let OS to have a refreshing time

        Dim iRetHeaderSent As Integer = &HFFFF, iRetDataSent As Integer = &HFFFF
        iRetHeaderSent = MyThrdServerProc.SendData2Server(udClientProc, RemoteIpEndPoint, 10000, bMyHeader)
        ListBox1.Items.Add("Send Header : " & iRetHeaderSent & " ID MSG : " & iMsgID)

        iRetDataSent = MyThrdServerProc.SendData2Server(udClientProc, RemoteIpEndPoint, 10000, sMyBuff)
        ListBox1.Items.Add("Send Data : " & iRetDataSent)
        If ListBox1.Items.Count > 50 Then ListBox1.Items.Clear()

        udClientProc.Close()
        udClientProc = Nothing
        RemoteIpEndPoint = Nothing
        GC.Collect()

    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        System.Threading.Thread.CurrentThread.Name = "MainProc"
        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")
        Dim b(254) As Byte, bb(254) As Byte, i As Integer, s As String
        For i = 1 To 255
            b(i - 1) = i
            s = s & Chr(b(i - 1))
        Next

        bb = System.Text.Encoding.ASCII.GetEncoding(1252).GetBytes(s)
    End Sub
End Class
