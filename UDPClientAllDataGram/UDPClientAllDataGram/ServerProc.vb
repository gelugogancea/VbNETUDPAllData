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
Public Class ServerProc
    Public Function SendData2Server(ByVal udClient As UdpClient, ByVal RemoteIpEndPoint As IPEndPoint, _
                                    ByVal MyTry As Integer, ByVal bData() As Byte) As Integer
        Dim iIfAvailable As Integer = 0, bRecvBytes(1) As Byte, iMyCountTry As Integer = 0
        Dim iRetFunc As Integer = 0
        udClient.Send(bData, bData.Length)
        Application.DoEvents()
        Do While (True)
            iIfAvailable = udClient.Available
            If iIfAvailable > 1 Then
                Try
                    bRecvBytes = udClient.Receive(RemoteIpEndPoint)
                Catch Ex As Exception
                    iMyCountTry = iMyCountTry + 1
                    Threading.Thread.Sleep(10)
                End Try

                If bRecvBytes(0) = 79 And bRecvBytes(1) = 75 Then     'If Ok => end of happy story
                    iRetFunc = 0
                Else
                    iRetFunc = -1
                End If
                GoTo MY_END
            Else
                iMyCountTry = iMyCountTry + 1
                Threading.Thread.Sleep(iMyCountTry)
            End If
        Loop
        iRetFunc = iMyCountTry
MY_END:
        SendData2Server = iRetFunc
    End Function
    Public Function ReceiveAnswerFromServer(ByVal MyTry As Integer, ByVal udClient As UdpClient, ByVal RemoteIpEndPoint As IPEndPoint) As Integer
        Dim iMyCountTry As Integer = 0, iRetFunction As Integer = 0, iIfDataAvailable As Integer = 0, bRecvBytes(1) As Byte

        Do While (True)
            iIfDataAvailable = udClient.Available
            If iIfDataAvailable > 1 Then
                Try
                    bRecvBytes = udClient.Receive(RemoteIpEndPoint)
                Catch Ex As Exception
                    iMyCountTry = iMyCountTry + 1
                    System.Threading.Thread.Sleep(10)
                End Try

                If bRecvBytes(0) = 79 And bRecvBytes(1) = 75 Then   'OK
                    iRetFunction = 0
                ElseIf bRecvBytes(0) = 78 And bRecvBytes(1) = 75 Then   'NK
                    iRetFunction = 1
                Else
                    iRetFunction = -2
                End If
                GoTo MY_END
            Else
                iMyCountTry = iMyCountTry + 1
                Threading.Thread.Sleep(1)
            End If

            If iMyCountTry >= MyTry Then
                iRetFunction = -1    '
                GoTo MY_END
            End If
        Loop
MY_END:
        ReceiveAnswerFromServer = iRetFunction
    End Function
    Public Function WaitThreadFromServer(ByVal MyTry As Integer, ByVal udClient As UdpClient, ByVal RemoteIpEndPoint As IPEndPoint) As Integer
        Dim bRecvBytes(1) As Byte, bConfirmBytes(1) As Byte
        Dim mStrucClass As New ClientStruct, iServerThreadConfirm As Integer = 0, iMyCountTry As Integer = 0, iIfDataAvailable As Integer = 0

        bConfirmBytes(0) = 240      '&HF0     send to server one confirmation that is known Server Thread is Ready
        bConfirmBytes(1) = 241      '&HF1

        Do While (True)
            If iMyCountTry >= MyTry Then
                iServerThreadConfirm = -1    '
                GoTo MY_END
            End If
            Try
                udClient.Send(bConfirmBytes, bConfirmBytes.Length)
            Catch
                iMyCountTry = iMyCountTry + 1
                System.Threading.Thread.Sleep(10)
            End Try
            Application.DoEvents()
            Threading.Thread.Sleep(10)
            iIfDataAvailable = udClient.Available
            If iIfDataAvailable > 1 Then
                Try
                    bRecvBytes = udClient.Receive(RemoteIpEndPoint)
                Catch Ex As Exception
                    iMyCountTry = iMyCountTry + 1
                    System.Threading.Thread.Sleep(10)
                End Try

                If bRecvBytes(0) = 240 And bRecvBytes(1) = 242 Then

                    iServerThreadConfirm = 0
                    GoTo MY_END
                End If
            Else
                iMyCountTry = iMyCountTry + 1
                Threading.Thread.Sleep(iMyCountTry)
            End If
        Loop
        iServerThreadConfirm = iMyCountTry

MY_END:
        WaitThreadFromServer = iServerThreadConfirm
        'udClient.Close()
        mStrucClass = Nothing
        'udClient = Nothing
        'RemoteIpEndPoint = Nothing
        GC.Collect()
    End Function
    Public Function EnCodeHeader(ByVal iMsgID As Integer, ByVal eMsgType As ClientStruct.MSG_TYPE, _
                                 ByVal iErrNo As Long, ByVal MaxNumbersOfParameters As Byte, _
                                 ByVal iDataGramLength As Integer, _
                                 ByRef bDataHeader() As Byte) As Integer

        Dim sTmp As String, bByteOrder(1) As Byte, bErrInByte(1) As Byte, bDataGramLength(1) As Byte
        Dim i As Integer, iCRC16 As Integer, bCRC(1) As Byte, bLengthDatagram(1) As Byte
        Dim m As New ClientStruct
        bByteOrder = m.HiLowInt16(iMsgID)
        bErrInByte = m.HiLowInt16(iErrNo)
        bDataGramLength = m.HiLowInt16(iDataGramLength)
        'Offset=0
        sTmp = Chr(1)                                           'Start Header
        'Offset +1
        sTmp = sTmp & Chr(bByteOrder(0)) & Chr(bByteOrder(1))   'MsgID
        'Offset+2
        sTmp = sTmp & Chr(MaxNumbersOfParameters)               'NumbersOfParameters
        'Offset+1
        sTmp = sTmp & Chr(CByte(eMsgType))                      'MessageType
        'Offset+1
        sTmp = sTmp & Chr(bDataGramLength(0)) & Chr(bDataGramLength(1))
        'Offset+2
        sTmp = sTmp & Chr(bErrInByte(0)) & Chr(bErrInByte(1))   'ErrorNumber
        'Offset+2
        iCRC16 = CRC16CCIT.ClassCRC16.CRC16(sTmp)
        bCRC = m.HiLowInt16(iCRC16)
        sTmp = sTmp & Chr(bCRC(0)) & Chr(bCRC(1))
        'bDataHeader = System.Text.Encoding.ASCII.GetEncoding(28591).GetBytes(sTmp)
        bDataHeader = System.Text.Encoding.ASCII.GetEncoding(1252).GetBytes(sTmp)

    End Function
    Public Function EnCodeMsgShort(ByVal sFrom As String, ByVal sTo As String, ByVal sObjectName As String, ByVal sParams() As String, _
                              ByRef bDataGramBuff() As Byte) As Integer
        Dim sTmp As String = "", bByteOrder(1) As Byte, bErrInByte(1) As Byte, bDataGramLength(1) As Byte
        Dim i As Integer, iCRC16 As Integer, bCRC(1) As Byte
        Dim m As New ClientStruct

        sTmp = sTmp & Chr(2)                                    'Start Text
        'Offset+1
        sTmp = sTmp & sFrom & vbTab & sTo & vbTab & sObjectName
        For i = 1 To sParams.Length
            sTmp = sTmp & vbTab & sParams(i - 1)
        Next i
        sTmp = sTmp & vbCrLf
        sTmp = sTmp & Chr(3)    'End Of Text
        iCRC16 = CRC16CCIT.ClassCRC16.CRC16(sTmp)
        bCRC = m.HiLowInt16(iCRC16)
        sTmp = sTmp & Chr(bCRC(0)) & Chr(bCRC(1))
        sTmp = sTmp & Chr(4)    'End of Transmision
        bDataGramBuff = System.Text.Encoding.ASCII.GetEncoding(1252).GetBytes(sTmp)
        EnCodeMsgShort = bDataGramBuff.Length

    End Function
End Class

