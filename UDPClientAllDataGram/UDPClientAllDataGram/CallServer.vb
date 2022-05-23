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
Public Class CallServer
    Public Function Request4Process(ByVal sIPAddress As String, ByVal iPort As Integer, ByVal MyTry As Integer) As Integer
        Dim udClient As UdpClient = Nothing, RemoteIpEndPoint As IPEndPoint = Nothing, bReq4Com(1) As Byte, bRecvBytes(1) As Byte, iRetBytes As Integer = 0, iNewUdpPort As Integer = 0
        Dim mStrucClass As New ClientStruct, iMyCountTry As Integer, iIfDataAvailable As Integer = 0
        udClient = New UdpClient()
        udClient.Connect(IPAddress.Parse(sIPAddress), iPort)
        RemoteIpEndPoint = New System.Net.IPEndPoint(System.Net.IPAddress.Any, 0)
        bReq4Com(0) = &HF9   '249 -0xF9
        bReq4Com(1) = &HF8   '248 -0xF8
        Try
            iRetBytes = udClient.Send(bReq4Com, bReq4Com.Length)
        Catch Ex As Exception
            iNewUdpPort = -1
            GoTo MY_END
        End Try

        Do While (iNewUdpPort <= 0)
            iIfDataAvailable = udClient.Available
            If iIfDataAvailable > 1 Then
                Try
                    bRecvBytes = udClient.Receive(RemoteIpEndPoint)
                Catch Ex As Exception
                    iMyCountTry = iMyCountTry + 1
                    System.Threading.Thread.Sleep(10)
                End Try

                iNewUdpPort = mStrucClass.BytesToInt16(bRecvBytes(0), bRecvBytes(1))
            Else
                If iMyCountTry > MyTry Then
                    iNewUdpPort = -2    '
                    Exit Do
                End If
                iMyCountTry = iMyCountTry + 1
                Threading.Thread.Sleep(1)
            End If
        Loop
MY_END:

        Request4Process = iNewUdpPort
        udClient.Close()
        udClient = Nothing
        RemoteIpEndPoint = Nothing
        mStrucClass = Nothing
        GC.Collect()
    End Function


End Class
