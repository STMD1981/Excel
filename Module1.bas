Attribute VB_Name = "Module1"
Option Explicit

'==========================================================
' Windows CNG API declarations
' Supports both 32-bit and 64-bit Excel
'==========================================================

#If VBA7 Then

    Private Declare PtrSafe Function BCryptOpenAlgorithmProvider _
        Lib "bcrypt.dll" ( _
        ByRef phAlgorithm As LongPtr, _
        ByVal pszAlgId As LongPtr, _
        ByVal pszImplementation As LongPtr, _
        ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function BCryptCreateHash _
        Lib "bcrypt.dll" ( _
        ByVal hAlgorithm As LongPtr, _
        ByRef phHash As LongPtr, _
        ByVal pbHashObject As LongPtr, _
        ByVal cbHashObject As Long, _
        ByVal pbSecret As LongPtr, _
        ByVal cbSecret As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function BCryptHashData _
        Lib "bcrypt.dll" ( _
        ByVal hHash As LongPtr, _
        ByRef pbInput As Any, _
        ByVal cbInput As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function BCryptFinishHash _
        Lib "bcrypt.dll" ( _
        ByVal hHash As LongPtr, _
        ByRef pbOutput As Any, _
        ByVal cbOutput As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function BCryptDestroyHash _
        Lib "bcrypt.dll" ( _
        ByVal hHash As LongPtr) As Long

    Private Declare PtrSafe Function BCryptCloseAlgorithmProvider _
        Lib "bcrypt.dll" ( _
        ByVal hAlgorithm As LongPtr, _
        ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function WideCharToMultiByte _
        Lib "kernel32.dll" ( _
        ByVal CodePage As Long, _
        ByVal dwFlags As Long, _
        ByVal lpWideCharStr As LongPtr, _
        ByVal cchWideChar As Long, _
        ByVal lpMultiByteStr As LongPtr, _
        ByVal cbMultiByte As Long, _
        ByVal lpDefaultChar As LongPtr, _
        ByVal lpUsedDefaultChar As LongPtr) As Long

#Else

    Private Declare Function BCryptOpenAlgorithmProvider _
        Lib "bcrypt.dll" ( _
        ByRef phAlgorithm As Long, _
        ByVal pszAlgId As Long, _
        ByVal pszImplementation As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare Function BCryptCreateHash _
        Lib "bcrypt.dll" ( _
        ByVal hAlgorithm As Long, _
        ByRef phHash As Long, _
        ByVal pbHashObject As Long, _
        ByVal cbHashObject As Long, _
        ByVal pbSecret As Long, _
        ByVal cbSecret As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare Function BCryptHashData _
        Lib "bcrypt.dll" ( _
        ByVal hHash As Long, _
        ByRef pbInput As Any, _
        ByVal cbInput As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare Function BCryptFinishHash _
        Lib "bcrypt.dll" ( _
        ByVal hHash As Long, _
        ByRef pbOutput As Any, _
        ByVal cbOutput As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare Function BCryptDestroyHash _
        Lib "bcrypt.dll" ( _
        ByVal hHash As Long) As Long

    Private Declare Function BCryptCloseAlgorithmProvider _
        Lib "bcrypt.dll" ( _
        ByVal hAlgorithm As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare Function WideCharToMultiByte _
        Lib "kernel32.dll" ( _
        ByVal CodePage As Long, _
        ByVal dwFlags As Long, _
        ByVal lpWideCharStr As Long, _
        ByVal cchWideChar As Long, _
        ByVal lpMultiByteStr As Long, _
        ByVal cbMultiByte As Long, _
        ByVal lpDefaultChar As Long, _
        ByVal lpUsedDefaultChar As Long) As Long

#End If

Private Const CP_UTF8 As Long = 65001
Private Const MD5_HASH_LENGTH As Long = 16


'==========================================================
' PUBLIC WORKSHEET FUNCTION
'
' If InputValue identifies an existing file:
'     Hashes the file's binary contents.
'
' Otherwise:
'     Hashes InputValue as UTF-8 text.
'==========================================================

Public Function MD5Checksum(ByVal InputValue As String) As Variant

    Dim Data() As Byte
    Dim DataLength As Long

    On Error GoTo ErrorHandler

    If ExistingFile(InputValue) Then
        Data = ReadBinaryFile(InputValue)
    Else
        Data = UTF8Bytes(InputValue)
    End If

    DataLength = SafeByteCount(Data)

    MD5Checksum = CalculateMD5(Data, DataLength)
    Exit Function

ErrorHandler:
    MD5Checksum = "ERROR " & CStr(Err.Number) & _
                  ": " & Err.Description

End Function


'==========================================================
' CHECK WHETHER INPUT IS AN EXISTING FILE
'==========================================================

Private Function ExistingFile(ByVal FilePath As String) As Boolean

    Dim FileAttributes As Long

    On Error GoTo NotFound

    If Len(FilePath) = 0 Then
        ExistingFile = False
        Exit Function
    End If

    FileAttributes = GetAttr(FilePath)

    ExistingFile = _
        ((FileAttributes And vbDirectory) <> vbDirectory)

    Exit Function

NotFound:
    ExistingFile = False

End Function


'==========================================================
' READ AN EXISTING FILE AS BINARY BYTES
'==========================================================

Private Function ReadBinaryFile( _
    ByVal FilePath As String) As Byte()

    Dim FileNumber As Integer
    Dim FileLength As Long
    Dim Data() As Byte

    On Error GoTo ReadError

    FileNumber = FreeFile

    Open FilePath For Binary Access Read As #FileNumber

    FileLength = LOF(FileNumber)

    If FileLength > 0 Then

        ReDim Data(0 To FileLength - 1)

        Get #FileNumber, 1, Data

    End If

    Close #FileNumber

    ReadBinaryFile = Data
    Exit Function

ReadError:

    If FileNumber <> 0 Then
        On Error Resume Next
        Close #FileNumber
        On Error GoTo 0
    End If

    Err.Raise Err.Number, _
              "ReadBinaryFile", _
              Err.Description

End Function


'==========================================================
' CONVERT VBA UNICODE TEXT TO UTF-8 BYTES
' WITHOUT A BYTE-ORDER MARK
'==========================================================

Private Function UTF8Bytes( _
    ByVal Text As String) As Byte()

    Dim Data() As Byte
    Dim RequiredBytes As Long
    Dim Result As Long

    If Len(Text) = 0 Then
        UTF8Bytes = Data
        Exit Function
    End If

    RequiredBytes = WideCharToMultiByte( _
        CP_UTF8, _
        0, _
        StrPtr(Text), _
        Len(Text), _
        0, _
        0, _
        0, _
        0)

    If RequiredBytes <= 0 Then
        Err.Raise vbObjectError + 100, _
                  "UTF8Bytes", _
                  "Unable to determine the UTF-8 byte length."
    End If

    ReDim Data(0 To RequiredBytes - 1)

    Result = WideCharToMultiByte( _
        CP_UTF8, _
        0, _
        StrPtr(Text), _
        Len(Text), _
        VarPtr(Data(0)), _
        RequiredBytes, _
        0, _
        0)

    If Result <> RequiredBytes Then
        Err.Raise vbObjectError + 101, _
                  "UTF8Bytes", _
                  "Unable to convert the text to UTF-8."
    End If

    UTF8Bytes = Data

End Function


'==========================================================
' RETURN THE NUMBER OF BYTES IN A BYTE ARRAY
'==========================================================

Private Function SafeByteCount( _
    ByRef Data() As Byte) As Long

    On Error GoTo EmptyArray

    SafeByteCount = _
        UBound(Data) - LBound(Data) + 1

    Exit Function

EmptyArray:
    SafeByteCount = 0

End Function


'==========================================================
' CALCULATE THE MD5 HASH USING WINDOWS CNG
'==========================================================

Private Function CalculateMD5( _
    ByRef Data() As Byte, _
    ByVal DataLength As Long) As String

#If VBA7 Then

    Dim AlgorithmHandle As LongPtr
    Dim HashHandle As LongPtr

#Else

    Dim AlgorithmHandle As Long
    Dim HashHandle As Long

#End If

    Dim Status As Long
    Dim HashBytes(0 To MD5_HASH_LENGTH - 1) As Byte
    Dim HashText As String
    Dim Position As Long

    On Error GoTo APIError

    Status = BCryptOpenAlgorithmProvider( _
        AlgorithmHandle, _
        StrPtr("MD5"), _
        0, _
        0)

    If Status <> 0 Then
        Err.Raise vbObjectError + 200, _
                  "CalculateMD5", _
                  "BCryptOpenAlgorithmProvider failed. " & _
                  "Status: 0x" & Right$("00000000" & _
                  Hex$(Status), 8)
    End If

    Status = BCryptCreateHash( _
        AlgorithmHandle, _
        HashHandle, _
        0, _
        0, _
        0, _
        0, _
        0)

    If Status <> 0 Then
        Err.Raise vbObjectError + 201, _
                  "CalculateMD5", _
                  "BCryptCreateHash failed. " & _
                  "Status: 0x" & Right$("00000000" & _
                  Hex$(Status), 8)
    End If

    If DataLength > 0 Then

        Status = BCryptHashData( _
            HashHandle, _
            Data(0), _
            DataLength, _
            0)

        If Status <> 0 Then
            Err.Raise vbObjectError + 202, _
                      "CalculateMD5", _
                      "BCryptHashData failed. " & _
                      "Status: 0x" & Right$("00000000" & _
                      Hex$(Status), 8)
        End If

    End If

    Status = BCryptFinishHash( _
        HashHandle, _
        HashBytes(0), _
        MD5_HASH_LENGTH, _
        0)

    If Status <> 0 Then
        Err.Raise vbObjectError + 203, _
                  "CalculateMD5", _
                  "BCryptFinishHash failed. " & _
                  "Status: 0x" & Right$("00000000" & _
                  Hex$(Status), 8)
    End If

    For Position = 0 To MD5_HASH_LENGTH - 1

        HashText = HashText & _
                   LCase$(Right$("0" & _
                   Hex$(HashBytes(Position)), 2))

    Next Position

    CalculateMD5 = HashText

CleanUp:

    If HashHandle <> 0 Then
        BCryptDestroyHash HashHandle
    End If

    If AlgorithmHandle <> 0 Then
        BCryptCloseAlgorithmProvider _
            AlgorithmHandle, 0
    End If

    Exit Function

APIError:

    If HashHandle <> 0 Then
        BCryptDestroyHash HashHandle
    End If

    If AlgorithmHandle <> 0 Then
        BCryptCloseAlgorithmProvider _
            AlgorithmHandle, 0
    End If

    Err.Raise Err.Number, _
              Err.Source, _
              Err.Description

End Function

