Option Explicit On
Option Compare Text
Public Class ClientStruct
    Enum MSG_TYPE
        OnDataReceive = 1
        OnObjectIdentified = 2
        OnWhileObjectPresent = 3
    End Enum
    Enum MSG_SEPARATOR
        MY_TAB = &H8
        MY_LINEFEED = &HA
        MY_ENTER = &HD
    End Enum
    Structure HeaderInfo
        'Offset=0
        Public SOH As Byte       'Start Of Header 
        'Offset+6
        Public HeaderSignature1 As Byte ' &FE
        Public HeaderSignature2 As Byte ' &FD
        Public HeaderSignature3 As Byte ' &FA - concatenate HeaderSignature = FEFDFA
        'Offset+2
        Public DataGramIDLowByte As Byte 'DataGram ID which should must to arrive , maximum 65535
        Public DataGramIDHighByte As Byte
        'Offset+1
        Public NumberOfParametersExpected As Byte ' no more than 255 . In this moment DataGramInfo.ParametersNo() shoud be ReDim(ension)
        'Offset+1
        Public MessageType As MSG_TYPE
        'Offset+2
        Public DataGramLengthLowByte As Byte 'needs a int16 number type  , maximum 65535
        Public DataGramLengthHighByte As Byte '
        'Offset+2
        Public LastErrorOccurLowByte As Byte
        Public LastErrorOccurHighByte As Byte ' error control handling int16 byte numeric type 
        'Offset+1
        Public HeaderLength As Byte ' in fact is most a confirmation that data was arrived in the right order.
        'Offset+1
        Public EOH As Byte  'End Of Header Info

        'HeaderLength is constant value of 15
        'Maxim Header Length = 14 Byte    
    End Structure
    Structure DataGramInfo
        'Offset+2
        Public CurentDataGramIDLowByte As Byte  'id message int16 numeric type
        Public CurentDataGramIDHighByte As Byte
        'Offset+2
        Public ErrorNoLowByte As Byte   'error handling int16 numeric type
        Public ErrorNoHighByte As Byte

        'After 5 byte start ASCII datagram and Tab
        Public MessageFrom As String
        Public MessageTo As String
        Public ObjectName As String
        Public ParametersNo() As String
    End Structure
    Structure MyInfo
        Public HeaderConfirmation As Byte 'confirmation that header is received and ACK
        Public ParametersConfirmation As Long 'confirmation that all parameters was received
        Public HowManyTimeResend As Long
        Public LastError As Long
        Public MyHeader As HeaderInfo
        Public MyMessage As DataGramInfo
    End Structure
    Public Function EnCodeMsgShort(ByVal sFrom As String, ByVal sTo As String, ByVal iMsgID As Integer, ByVal eMsgType As MSG_TYPE, _
                          ByVal iErrNo As Long, ByVal sObjectName As String, ByVal sParams() As String, _
                          ByRef bDataGramBuff() As Byte) As Integer
        Dim sTmp As String, bByteOrder(1) As Byte, bErrInByte(1) As Byte, bDataGramLength(1) As Byte
        Dim i As Integer, MaxNumbersOfParameters As Byte
        Dim enc As New System.Text.UTF8Encoding()
        bByteOrder = HiLowInt16(iMsgID)
        bErrInByte = HiLowInt16(iErrNo)
        MaxNumbersOfParameters = sParams.GetUpperBound(0) + 1
        'Offset=0
        sTmp = Chr(1)                                           'Start Header
        'Offset +1
        sTmp = sTmp & Chr(&HFE) & Chr(&HFD) & Chr(&HFA)         'Header Signature  
        'Offset+3
        sTmp = sTmp & Chr(bByteOrder(0)) & Chr(bByteOrder(1))   'MsgID
        'Offset+2
        sTmp = sTmp & Chr(MaxNumbersOfParameters)               'NumbersOfParameters
        'Offset+1
        sTmp = sTmp & Chr(CByte(eMsgType))                      'MessageType
        'Offset+1
        sTmp = sTmp & Chr(255) & Chr(255)                       'Reserved two byte for the entire diagram length
        'Offset+2
        sTmp = sTmp & Chr(bErrInByte(0)) & Chr(bErrInByte(1))   'ErrorNumber
        'Offset+2
        sTmp = sTmp & Chr(2)                                    'Start Text
        'Offset+1
        sTmp = sTmp & sFrom & vbTab & sTo & vbTab & sObjectName
        For i = 1 To MaxNumbersOfParameters
            sTmp = sTmp & vbTab & sParams(i - 1)
        Next i
        sTmp = sTmp & Chr(3)    'End Of Text
        sTmp = sTmp & Chr(4)    'End Of Message
        bDataGramLength = HiLowInt16(sTmp.Length)
        bDataGramBuff = System.Text.Encoding.ASCII.GetEncoding(28591).GetBytes(sTmp)
        bDataGramBuff(8) = bDataGramLength(0)
        bDataGramBuff(9) = bDataGramLength(1)
        EnCodeMsgShort = bDataGramBuff.GetUpperBound(0) + 1
    End Function
    Public Function HiLowInt16(ByVal wparam As Integer) As Byte()
        Dim y(1) As Byte
        y(0) = wparam And &HFF&     'Low
        y(1) = wparam \ &H100 And &HFF& 'High
        HiLowInt16 = y
    End Function
    Public Function HiLowLong32(ByVal lparam As Long) As Integer()
        Dim y(1) As Integer
        y(0) = lparam And &HFFFF&   'Low
        y(1) = lparam \ &H10000 And &HFFFF& 'High
        HiLowLong32 = y
    End Function
    Public Function BytesToInt16(ByVal LoByte As Byte, ByVal HiByte As Byte) As UShort
        BytesToInt16 = (HiByte * &H100) + (LoByte And &HFF&)
    End Function
    Public Function Int16ToLong(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
        Int16ToLong = (HiWord * &H10000) + (LoWord And &HFFFF&)
    End Function
End Class

