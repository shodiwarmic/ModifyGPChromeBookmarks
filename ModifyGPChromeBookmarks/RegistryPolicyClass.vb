Imports System.Text.RegularExpressions

Public Class RegistryPolicyClass
    Friend Const REG_NONE = 0
    Friend Const REG_SZ = 1
    Friend Const REG_EXPAND_SZ = 2
    Friend Const REG_BINARY = 3
    Friend Const REG_DWORD = 4
    Friend Const REG_DWORD_BIG_ENDIAN = 5
    Friend Const REG_MULTI_SZ = 7
    Friend Const REG_QWORD = 11

    Friend Const PF_SUCCESS = 0
    Friend Const PF_FILENOEXIST = 1
    Friend Const PF_READERROR = 2
    Friend Const PF_WRITEERROR = 3
    Friend Const PF_FORMATERROR = 4
    Friend Const PF_INVALID_HEADER = 5
    Friend Const PF_NOTLOADED = 6
    Friend Const PF_VALUENOEXIST = 7
    Friend Const PF_FILEALREADYEXISTS = 9

    Const cPolHeader = "67655250"
    Const cPolVersion = "00000001"

    Function RegTypeToString(intType)
        RegTypeToString = "**UNKNOWN**"

        Select Case intType
            Case REG_NONE
                RegTypeToString = "REG_NONE"
            Case REG_SZ
                RegTypeToString = "REG_SZ"
            Case REG_EXPAND_SZ
                RegTypeToString = "REG_EXPAND_SZ"
            Case REG_BINARY
                RegTypeToString = "REG_BINARY"
            Case REG_DWORD
                RegTypeToString = "REG_DWORD"
            Case REG_DWORD_BIG_ENDIAN
                RegTypeToString = "REG_DWORD_BIG_ENDIAN"
            Case REG_MULTI_SZ
                RegTypeToString = "REG_MULTI_SZ"
            Case REG_QWORD
                RegTypeToString = "REG_QWORD"
        End Select
    End Function

    Class PolSetting
        Property varData
        Property intType
    End Class

    Class PolFile
        Private m_objSettings
        Private m_blnLoaded

        Private Sub Class_Initialize()
            m_objSettings = CreateObject("Scripting.Dictionary")
            m_objSettings.CompareMode = 1 'vbText
            m_blnLoaded = False
        End Sub

        Public Function LoadFile(strFilePath)
            LoadFile = PF_SUCCESS

            ' Reinitialize member settings dictionary (in case this function aborts)
            m_objSettings = CreateObject("Scripting.Dictionary")
            m_objSettings.CompareMode = 1 'vbText
            m_blnLoaded = False

            ' Create temporary settings object (to be assigned to member settings if the function is successful)
            Dim objSettings

            objSettings = CreateObject("Scripting.Dictionary")
            objSettings.CompareMode = 1 'vbText

            ' Check for file
            Dim objFSO
            objFSO = CreateObject("Scripting.FileSystemObject")

            If (objFSO.FileExists(strFilePath) = False) Then
                LoadFile = PF_FILENOEXIST
                Exit Function
            End If

            ' Load
            Dim objBuffer
            objBuffer = My.Computer.FileSystem.ReadAllBytes(strFilePath)

            If (UBound(objBuffer) < 0) Then
                LoadFile = PF_READERROR
                Exit Function
            End If

            ' Check for presence of header
            If (UBound(objBuffer) < 7) Then
                LoadFile = PF_INVALID_HEADER
                Exit Function
            End If

            ' Check header contents
            Dim strHeader
            strHeader = ""

            Dim strHex
            Dim i, x

            For i = 3 To 0 Step -1
                strHex = Hex(objBuffer(i))
                If (Len(strHex) = 1) Then strHex = "0" & strHex

                strHeader = strHeader & strHex
            Next

            If (strHeader <> cPolHeader) Then
                LoadFile = PF_INVALID_HEADER
                Exit Function
            End If

            strHeader = ""
            For i = 7 To 4 Step -1
                strHex = Hex(objBuffer(i))
                If (Len(strHex) = 1) Then strHex = "0" & strHex

                strHeader = strHeader & strHex
            Next

            If (strHeader <> cPolVersion) Then
                LoadFile = PF_INVALID_HEADER
                Exit Function
            End If

            ' We have a proper header, now process the entries.

            Dim intVal
            Dim strChar

            Dim strKey, strValue, intType, intSize, varData, arrData()

            i = 8
            Do While (i < (UBound(objBuffer) - 1))

                ' Check for presence of opening square bracket
                If ((objBuffer(i) <> &H5B) Or (objBuffer(i + 1) <> 0)) Then Exit Do
                i = i + 2

                strKey = ""
                strValue = ""
                intType = -1
                intSize = -1
                varData = ""
                ReDim arrData(0)

                ' Read Key
                Do While (True)
                    If (i >= UBound(objBuffer)) Then
                        LoadFile = PF_FORMATERROR
                        Exit Function
                    End If

                    intVal = (objBuffer(i + 1) * 256) + objBuffer(i)
                    strChar = ChrW(intVal)

                    i = i + 2

                    If (intVal = 0) Then
                        ' We've reached the null terminator.  Test for presence of semicolon separator
                        If (i >= UBound(objBuffer)) Then
                            LoadFile = PF_FORMATERROR
                            Exit Function
                        End If

                        If ((objBuffer(i) <> &H3B) Or (objBuffer(i + 1) <> 0)) Then
                            LoadFile = PF_FORMATERROR
                            Exit Function
                        End If

                        i = i + 2

                        Exit Do
                    End If

                    strKey = strKey & strChar
                Loop

                ' Read Value Name
                Do While (True)
                    If (i >= UBound(objBuffer)) Then
                        LoadFile = PF_FORMATERROR
                        Exit Function
                    End If

                    intVal = (objBuffer(i + 1) * 256) + objBuffer(i)
                    strChar = ChrW(intVal)

                    i = i + 2

                    If (intVal = 0) Then
                        ' We've reached the null terminator.  Test for presence of semicolon separator
                        If (i >= UBound(objBuffer)) Then
                            LoadFile = PF_FORMATERROR
                            Exit Function
                        End If

                        If ((objBuffer(i) <> &H3B) Or (objBuffer(i + 1) <> 0)) Then
                            LoadFile = PF_FORMATERROR
                            Exit Function
                        End If

                        i = i + 2

                        Exit Do
                    End If

                    strValue = strValue & strChar
                Loop

                ' Read Type
                If (True) Then
                    ' Conditional is just an excuse to indent this section, making it easier to read (and fold with editor)
                    If (i >= UBound(objBuffer) - 2) Then
                        LoadFile = PF_FORMATERROR
                        Exit Function
                    End If

                    intVal = (objBuffer(i + 1) * 256) + objBuffer(i)

                    i = i + 4

                    intType = intVal

                    ' Test for presence of semicolon separator
                    If (i >= UBound(objBuffer)) Then
                        LoadFile = PF_FORMATERROR
                        Exit Function
                    End If

                    If ((objBuffer(i) <> &H3B) Or (objBuffer(i + 1) <> 0)) Then
                        LoadFile = PF_FORMATERROR
                        Exit Function
                    End If

                    i = i + 2
                End If

                ' Read Size
                If (True) Then
                    ' Conditional is just an excuse to indent this section, making it easier to read (and fold with editor)
                    If (i >= UBound(objBuffer) - 2) Then
                        LoadFile = PF_FORMATERROR
                        Exit Function
                    End If

                    intVal = (objBuffer(i + 1) * 256) + objBuffer(i)

                    i = i + 4

                    intSize = intVal

                    ' Test for presence of semicolon separator
                    If (i >= UBound(objBuffer)) Then
                        LoadFile = PF_FORMATERROR
                        Exit Function
                    End If

                    If ((objBuffer(i) <> &H3B) Or (objBuffer(i + 1) <> 0)) Then
                        LoadFile = PF_FORMATERROR
                        Exit Function
                    End If

                    i = i + 2
                End If

                ' Read Data
                If (True) Then
                    ' Conditional is just an excuse to indent this section, making it easier to read (and fold with editor)
                    If ((i + intSize) > UBound(objBuffer)) Then
                        LoadFile = PF_FORMATERROR
                        Exit Function
                    End If

                    Select Case intType
                        Case REG_NONE
                    ' Do nothing
                        Case REG_SZ
                            ' Unicode string.  intSize includes the null terminator, so ignore the last two bytes
                            For x = 0 To intSize - 3 Step 2
                                intVal = (objBuffer(i + x + 1) * 256) + objBuffer(i + x)
                                strChar = ChrW(intVal)

                                varData = varData & strChar
                            Next
                        Case REG_EXPAND_SZ
                            ' Same as REG_SZ for our purposes
                            For x = 0 To intSize - 3 Step 2
                                intVal = (objBuffer(i + x + 1) * 256) + objBuffer(i + x)
                                strChar = ChrW(intVal)

                                varData = varData & strChar
                            Next
                        Case REG_BINARY
                            ' not sure of the byte order here, and I've never encountered a REG_BINARY in a pol file yet for testing.
                            For x = 0 To intSize - 1
                                strChar = Hex(objBuffer(i + x))
                                If (Len(strChar) = 1) Then strChar = "0" & strChar

                                varData = varData & strChar
                            Next
                        Case REG_DWORD
                            ' Little-endian
                            varData = CLng(0)

                            For x = 0 To intSize - 1
                                varData = CLng(varData) + (CLng(objBuffer(i + x)) * (256 ^ x))
                            Next
                        Case REG_DWORD_BIG_ENDIAN
                            ' Big-endian
                            varData = CLng(0)

                            For x = intSize - 1 To 0 Step -1
                                varData = CLng(varData) + (CLng(objBuffer(i + x)) * (256 ^ ((intSize - 1) - x)))
                            Next
                        Case REG_QWORD
                            ' Little-endian.  Since VBScript can't really handle unsigned 64-bit numbers, convert this
                            ' to a string containing hex characters, similar to binary.
                            varData = ""

                            For x = intSize - 1 To 0 Step -1
                                strChar = Hex(objBuffer(i + x))
                                If (Len(strChar) = 1) Then strChar = "0" & strChar

                                varData = varData & strChar
                            Next
                        Case REG_MULTI_SZ
                            ' Terminated by two null characters.  Before then, zero or more null-terminated Unicode strings.

                            ReDim arrData(0)

                            For x = 0 To intSize - 5 Step 2
                                intVal = (objBuffer(i + x + 1) * 256) + objBuffer(i + x)
                                strChar = ChrW(intVal)

                                If (intVal = 0) Then
                                    ' Null terminator.  Begin a new string
                                    ReDim Preserve arrData(UBound(arrData) + 1)
                                    arrData(UBound(arrData)) = ""
                                Else
                                    arrData(UBound(arrData)) = arrData(UBound(arrData)) & strChar
                                End If
                            Next
                        Case Else
                            LoadFile = PF_FORMATERROR
                            Exit Function
                    End Select

                    i = i + intSize

                    ' Test for presence of closing square bracket
                    If (i >= UBound(objBuffer)) Then
                        LoadFile = PF_FORMATERROR
                        Exit Function
                    End If

                    If ((objBuffer(i) <> &H5D) Or (objBuffer(i + 1) <> 0)) Then
                        LoadFile = PF_FORMATERROR
                        Exit Function
                    End If

                    i = i + 2
                End If

                objSettings(strKey & "\" & strValue) = New PolSetting
                objSettings(strKey & "\" & strValue).intType = intType

                If (intType = REG_NONE) Then
                    objSettings(strKey & "\" & strValue).varData = ""
                ElseIf (intType = REG_MULTI_SZ) Then
                    objSettings(strKey & "\" & strValue).varData = arrData

                Else
                    objSettings(strKey & "\" & strValue).varData = varData
                End If
            Loop

            m_objSettings = objSettings

            m_blnLoaded = True
        End Function

        Public Function Save(strFilePath, blnOverwriteExisting)
            If (m_blnLoaded = False) Then
                Save = PF_NOTLOADED
                Exit Function
            End If

            Dim objFSO
            objFSO = CreateObject("Scripting.FileSystemObject")

            If (objFSO.FileExists(strFilePath) = True) Then
                If (blnOverwriteExisting = False) Then
                    Save = PF_FILEALREADYEXISTS
                    Exit Function
                End If
            End If

            ' Friend Construct buffer of bytes
            Dim objBuffer()
            ReDim objBuffer(-1)

            ' Add header
            AddByte(objBuffer, &H50)
            AddByte(objBuffer, &H52)
            AddByte(objBuffer, &H65)
            AddByte(objBuffer, &H67)
            AddByte(objBuffer, &H1)
            AddByte(objBuffer, &H0)
            AddByte(objBuffer, &H0)
            AddByte(objBuffer, &H0)

            ' Process contents of m_objSettings
            Dim strKey

            Dim strRegKey, strRegValue
            Dim intUnicode, intFirstByte, intSecondByte
            Dim i

            Dim objRegExp, colMatches, objMatch, blnMatchSuccess

            objRegExp = New Regex("^(.+)\\([^\\]*)$")

            For Each strKey In m_objSettings.Keys
                ' Separate Key and Value name
                strRegKey = ""
                strRegValue = ""

                blnMatchSuccess = False
                colMatches = objRegExp.Matches(strKey)

                For Each objMatch In colMatches
                    blnMatchSuccess = True
                    strRegKey = objMatch.Groups(1).Value
                    strRegValue = objMatch.Groups(2).Value
                Next

                If (blnMatchSuccess = False) Then
                    Save = PF_FORMATERROR
                    Exit Function
                End If

                ' Add open square bracket
                AddByte(objBuffer, &H5B)
                AddByte(objBuffer, &H0)

                ' Convert Key name to bytes
                For i = 1 To Len(strRegKey)
                    intUnicode = AscW(Mid(strRegKey, i, 1))

                    intFirstByte = intUnicode Mod 256
                    intSecondByte = (intUnicode - intFirstByte) / 256

                    AddByte(objBuffer, intFirstByte)
                    AddByte(objBuffer, intSecondByte)
                Next

                ' Add null terminator and semicolon separator
                AddByte(objBuffer, &H0)
                AddByte(objBuffer, &H0)
                AddByte(objBuffer, &H3B)
                AddByte(objBuffer, &H0)

                ' Convert Value name to bytes
                For i = 1 To Len(strRegValue)
                    intUnicode = AscW(Mid(strRegValue, i, 1))

                    intFirstByte = intUnicode Mod 256
                    intSecondByte = (intUnicode - intFirstByte) / 256

                    AddByte(objBuffer, intFirstByte)
                    AddByte(objBuffer, intSecondByte)
                Next

                ' Add null terminator and semicolon separator
                AddByte(objBuffer, &H0)
                AddByte(objBuffer, &H0)
                AddByte(objBuffer, &H3B)
                AddByte(objBuffer, &H0)

                ' Add Type and semicolon separator

                AddByte(objBuffer, m_objSettings(strKey).intType)
                AddByte(objBuffer, &H0)
                AddByte(objBuffer, &H0)
                AddByte(objBuffer, &H0)
                AddByte(objBuffer, &H3B)
                AddByte(objBuffer, &H0)

                Select Case m_objSettings(strKey).intType
                    Case REG_NONE
                        If (AddBinaryBytes(objBuffer, m_objSettings(strKey).varData) = False) Then
                            Save = PF_FORMATERROR
                            Exit Function
                        End If
                    Case REG_BINARY
                        If (AddBinaryBytes(objBuffer, m_objSettings(strKey).varData) = False) Then
                            Save = PF_FORMATERROR
                            Exit Function
                        End If
                    Case REG_SZ
                        If (AddStringBytes(objBuffer, m_objSettings(strKey).varData) = False) Then
                            Save = PF_FORMATERROR
                            Exit Function
                        End If
                    Case REG_EXPAND_SZ
                        If (AddStringBytes(objBuffer, m_objSettings(strKey).varData) = False) Then
                            Save = PF_FORMATERROR
                            Exit Function
                        End If
                    Case REG_MULTI_SZ
                        If (AddMultiStringBytes(objBuffer, m_objSettings(strKey).varData) = False) Then
                            Save = PF_FORMATERROR
                            Exit Function
                        End If
                    Case REG_DWORD
                        If (AddDWORDBytes(objBuffer, m_objSettings(strKey).varData) = False) Then
                            Save = PF_FORMATERROR
                            Exit Function
                        End If
                    Case REG_DWORD_BIG_ENDIAN
                        If (AddDWORD_BigEndianBytes(objBuffer, m_objSettings(strKey).varData) = False) Then
                            Save = PF_FORMATERROR
                            Exit Function
                        End If
                    Case REG_QWORD
                        If (AddQWORDBytes(objBuffer, m_objSettings(strKey).varData) = False) Then
                            Save = PF_FORMATERROR
                            Exit Function
                        End If
                    Case Else
                        Save = PF_FORMATERROR
                        Exit Function
                End Select
            Next

            Save = WriteBinary(strFilePath, objBuffer)
        End Function

        Public ReadOnly Property Keys()
            Get
                Keys = m_objSettings.Keys
            End Get
        End Property

        Public Function GetValue(strKey, ByRef varData, ByRef intType)
            If (Not m_blnLoaded) Then
                GetValue = PF_NOTLOADED
                Exit Function
            End If

            If (m_objSettings.Exists(strKey) = True) Then
                GetValue = PF_SUCCESS

                intType = m_objSettings(strKey).intType
                varData = m_objSettings(strKey).varData

                Exit Function
            Else
                GetValue = PF_VALUENOEXIST
                Exit Function
            End If
        End Function

        Public Function SetValue(strKey, strValue, intType, varData)
            SetValue = PF_FORMATERROR

            Dim objRegExp, colMatches, objMatch
            objRegExp = New Regex("[^0-9A-F]")

            Dim objSetting
            objSetting = New PolSetting

            Select Case intType
                Case REG_NONE
                    objSetting.intType = REG_NONE
                    objSetting.varData = ""
                Case REG_SZ
                    On Error Resume Next
                    varData = CStr(varData)

                    If (Err.Number <> 0) Then Exit Function
                    On Error GoTo 0

                    objSetting.intType = REG_SZ
                    objSetting.varData = varData
                Case REG_EXPAND_SZ
                    On Error Resume Next
                    varData = CStr(varData)

                    If (Err.Number <> 0) Then Exit Function
                    On Error GoTo 0

                    objSetting.intType = REG_EXPAND_SZ
                    objSetting.varData = varData
                Case REG_BINARY
                    On Error Resume Next
                    varData = CStr(varData)

                    If (Err.Number <> 0) Then Exit Function
                    On Error GoTo 0

                    objRegExp.Pattern = "[^0-9A-F]"
                    objRegExp.IgnoreCase = True

                    ' String should contain only valid Hex characters
                    If (objRegExp.Test(varData) = True) Then Exit Function

                    ' String should contain bytes (even number of hex characters)
                    If ((Len(varData) Mod 2) = 1) Then Exit Function

                    objSetting.intType = REG_BINARY
                    objSetting.varData = varData
                Case REG_DWORD
                    On Error Resume Next
                    varData = CLng(varData)

                    If (Err.Number <> 0) Then Exit Function
                    On Error GoTo 0

                    objSetting.intType = REG_DWORD
                    objSetting.varData = varData
                Case REG_DWORD_BIG_ENDIAN
                    On Error Resume Next
                    varData = CLng(varData)

                    If (Err.Number <> 0) Then Exit Function
                    On Error GoTo 0

                    objSetting.intType = REG_DWORD_BIG_ENDIAN
                    objSetting.varData = varData
                Case REG_MULTI_SZ
                    ' varData must be an array containing zero or more strings
                    Dim i

                    If (IsArray(varData) = False) Then Exit Function

                    On Error Resume Next

                    For i = LBound(varData) To UBound(varData)
                        varData(i) = CStr(varData(i))

                        If (Err.Number <> 0) Then Exit Function
                    Next

                    objSetting.intType = REG_MULTI_SZ
                    objSetting.varData = varData
                Case REG_QWORD
                    On Error Resume Next
                    varData = CStr(varData)

                    If (Err.Number <> 0) Then Exit Function
                    On Error GoTo 0

                    objRegExp.Pattern = "[^0-9A-F]"
                    objRegExp.IgnoreCase = True

                    ' String should contain only valid Hex characters
                    If (objRegExp.Test(varData) = True) Then Exit Function

                    ' String should contain no more than 8 bytes (16 characters)
                    If (Len(varData) > 16) Then Exit Function

                    objSetting.intType = REG_QWORD
                    objSetting.varData = varData
                Case Else
                    Exit Function
            End Select

            m_objSettings(strKey & "\" & strValue) = objSetting
            SetValue = PF_SUCCESS
        End Function

        Public Function DeleteValue(strKey, strValue)
            If (Not m_blnLoaded) Then
                DeleteValue = PF_NOTLOADED
                Exit Function
            End If

            If (m_objSettings.Exists(strKey & "\" & strValue) = True) Then
                DeleteValue = PF_SUCCESS

                m_objSettings.Remove(strKey & "\" & strValue)

                Exit Function
            Else
                DeleteValue = PF_VALUENOEXIST
                Exit Function
            End If
        End Function

        Public Function DeleteValue(strKey)
            If (Not m_blnLoaded) Then
                DeleteValue = PF_NOTLOADED
                Exit Function
            End If

            If (m_objSettings.Exists(strKey) = True) Then
                DeleteValue = PF_SUCCESS

                m_objSettings.Remove(strKey)

                Exit Function
            Else
                DeleteValue = PF_VALUENOEXIST
                Exit Function
            End If
        End Function

        Private Sub AddByte(ByRef objBuffer, intByte)
            ReDim Preserve objBuffer(UBound(objBuffer) + 1)
            objBuffer(UBound(objBuffer)) = intByte
        End Sub

        Private Function AddBinaryBytes(ByRef objBuffer, ByRef varData)
            AddBinaryBytes = False

            ' varData is a string containing text representation of hex bytes

            Dim intSize
            intSize = Len(varData)

            If (intSize Mod 2 = 1) Then
                ' This shouldn't happen unless someone has forgotten to pad zeroes, which results in corrupt data
                ' The text should always be pairs of hex digits (even number of characters)
                Exit Function
            End If

            intSize = intSize / 2
            If (intSize > 65535) Then
                ' Maximum size limit exceeded
                Exit Function
            End If

            Dim intFirstByte, intSecondByte

            intFirstByte = intSize Mod 256
            intSecondByte = (intSize - intFirstByte) / 256

            ' Write Size and separator to stream
            AddByte(objBuffer, intFirstByte)
            AddByte(objBuffer, intSecondByte)
            AddByte(objBuffer, &H0)
            AddByte(objBuffer, &H0)
            AddByte(objBuffer, &H3B)
            AddByte(objBuffer, &H0)

            Dim i, strByte
            For i = 1 To Len(varData) Step 2
                strByte = Mid(varData, i, 2)

                On Error Resume Next
                intFirstByte = CLng("&H" & strByte)

                If (Err.Number <> 0) Then
                    Exit Function
                End If
                On Error GoTo 0

                AddByte(objBuffer, intFirstByte)
            Next

            ' Add closing square bracket
            AddByte(objBuffer, &H5D)
            AddByte(objBuffer, &H0)

            AddBinaryBytes = True
        End Function

        Private Function AddStringBytes(ByRef objBuffer, ByRef varData)
            AddStringBytes = False

            ' varData is a unicode string that needs to be converted to byte pairs in little-endian order

            Dim intSize

            ' Size (in bytes) is two for each character in the string, plus two bytes for a null terminator.
            intSize = (Len(varData) * 2) + 2

            If (intSize > 65535) Then
                ' Maximum size limit exceeded
                Exit Function
            End If

            Dim intFirstByte, intSecondByte, intUnicode

            intFirstByte = intSize Mod 256
            intSecondByte = (intSize - intFirstByte) / 256

            ' Write Size and separator to stream
            AddByte(objBuffer, intFirstByte)
            AddByte(objBuffer, intSecondByte)
            AddByte(objBuffer, &H0)
            AddByte(objBuffer, &H0)
            AddByte(objBuffer, &H3B)
            AddByte(objBuffer, &H0)

            Dim i

            For i = 1 To Len(varData)
                intUnicode = AscW(Mid(varData, i, 1))

                intFirstByte = intUnicode Mod 256
                intSecondByte = (intUnicode - intFirstByte) / 256

                AddByte(objBuffer, intFirstByte)
                AddByte(objBuffer, intSecondByte)
            Next

            ' Add null terminator
            AddByte(objBuffer, &H0)
            AddByte(objBuffer, &H0)

            ' Add closing square bracket
            AddByte(objBuffer, &H5D)
            AddByte(objBuffer, &H0)

            AddStringBytes = True
        End Function

        Private Function AddMultiStringBytes(ByRef objBuffer, ByRef varData)
            AddMultiStringBytes = False

            ' varData is an array of zero or more unicode strings that needs to be converted to byte pairs in little-endian order
            ' and separated by null terminators.  There will be a double null terminator at the end of all the strings.

            If (IsArray(varData) = False) Then Exit Function

            Dim i, x

            Dim intSize
            intSize = 4 ' 4 extra bytes for the double null terminators at the end of the string

            For x = LBound(varData) To UBound(varData)
                intSize = intSize + (Len(varData(x)) * 2) + 2
            Next

            ' If strings are present in the array, ignore the extra terminator our loop just added
            If (intSize > 4) Then intSize = intSize - 2

            If (intSize > 65535) Then
                '            Maximum size limit exceeded
                Exit Function
            End If

            Dim intFirstByte, intSecondByte, intUnicode

            intFirstByte = intSize Mod 256
            intSecondByte = (intSize - intFirstByte) / 256

            ' Write Size and separator to stream
            AddByte(objBuffer, intFirstByte)
            AddByte(objBuffer, intSecondByte)
            AddByte(objBuffer, &H0)
            AddByte(objBuffer, &H0)
            AddByte(objBuffer, &H3B)
            AddByte(objBuffer, &H0)

            For x = LBound(varData) To UBound(varData)
                For i = 1 To Len(varData(x))
                    intUnicode = AscW(Mid(varData(x), i, 1))

                    intFirstByte = intUnicode Mod 256
                    intSecondByte = (intUnicode - intFirstByte) / 256

                    AddByte(objBuffer, intFirstByte)
                    AddByte(objBuffer, intSecondByte)
                Next

                ' Add null terminator, except for last element in the array (which will be covered
                ' when the double null terminator is added after this loop).
                If (x <> UBound(varData)) Then
                    AddByte(objBuffer, &H0)
                    AddByte(objBuffer, &H0)
                End If
            Next

            ' Add double null terminator
            AddByte(objBuffer, &H0)
            AddByte(objBuffer, &H0)
            AddByte(objBuffer, &H0)
            AddByte(objBuffer, &H0)

            ' Add closing square bracket
            AddByte(objBuffer, &H5D)
            AddByte(objBuffer, &H0)

            AddMultiStringBytes = True
        End Function

        Private Function AddDWORDBytes(ByRef objBuffer, ByRef varData)
            AddDWORDBytes = False

            ' varData is a number.  Convert it to a hex string with the UnsignedHex
            ' function, make sure the number of bytes is legal (no more than 4 for a DWORD), then
            ' write it to the stream in little-endian order.

            On Error Resume Next
            varData = CLng(varData)

            If (Err.Number <> 0) Then Exit Function
            On Error GoTo 0


            Dim strHex
            strHex = UnsignedHex(varData)

            If (Len(strHex) > 8) Then
                ' number is too large for a DWORD
                Exit Function
            End If

            Do While (Len(strHex) < 8)
                strHex = "0" & strHex
            Loop

            ' Write Size (always 4) and separator to stream
            AddByte(objBuffer, &H4)
            AddByte(objBuffer, &H0)
            AddByte(objBuffer, &H0)
            AddByte(objBuffer, &H0)
            AddByte(objBuffer, &H3B)
            AddByte(objBuffer, &H0)

            Dim i, strByte, intByte
            For i = 7 To 1 Step -2
                strByte = Mid(strHex, i, 2)

                On Error Resume Next
                intByte = CLng("&H" & strByte)

                If (Err.Number <> 0) Then
                    Exit Function
                End If
                On Error GoTo 0

                AddByte(objBuffer, intByte)
            Next

            ' Add closing square bracket
            AddByte(objBuffer, &H5D)
            AddByte(objBuffer, &H0)

            AddDWORDBytes = True
        End Function

        Private Function AddDWORD_BigEndianBytes(ByRef objBuffer, ByRef varData)
            AddDWORD_BigEndianBytes = False

            ' varData is a number.  Convert it to a hex string with the UnsignedHex
            ' function, make sure the number of bytes is legal (no more than 4 for a DWORD), then
            ' write it to the stream in big-endian order.

            On Error Resume Next
            varData = CLng(varData)

            If (Err.Number <> 0) Then Exit Function
            On Error GoTo 0

            Dim strHex
            strHex = UnsignedHex(varData)

            If (Len(strHex) > 8) Then
                ' number is too large for a DWORD
                Exit Function
            End If

            Do While (Len(strHex) < 8)
                strHex = "0" & strHex
            Loop

            ' Write Size (always 4) and separator to stream
            AddByte(objBuffer, &H4)
            AddByte(objBuffer, &H0)
            AddByte(objBuffer, &H0)
            AddByte(objBuffer, &H0)
            AddByte(objBuffer, &H3B)
            AddByte(objBuffer, &H0)

            Dim i, strByte, intByte
            For i = 1 To 7 Step 2
                strByte = Mid(strHex, i, 2)

                On Error Resume Next
                intByte = CLng("&H" & strByte)

                If (Err.Number <> 0) Then
                    Exit Function
                End If
                On Error GoTo 0

                AddByte(objBuffer, intByte)
            Next

            ' Add closing square bracket
            AddByte(objBuffer, &H5D)
            AddByte(objBuffer, &H0)

            AddDWORD_BigEndianBytes = True
        End Function

        Private Function AddQWORDBytes(ByRef objBuffer, ByRef varData)
            AddQWORDBytes = False

            ' varData is a string containing hex characters (Up to 8 bytes, 16 characters)

            Dim strHex
            strHex = varData

            If (Len(strHex) > 16) Then
                ' number is too large for a QWORD
                Exit Function
            End If

            Do While (Len(strHex) < 16)
                strHex = "0" & strHex
            Loop

            ' Write Size (always 8) and separator to stream
            AddByte(objBuffer, &H8)
            AddByte(objBuffer, &H0)
            AddByte(objBuffer, &H0)
            AddByte(objBuffer, &H0)
            AddByte(objBuffer, &H3B)
            AddByte(objBuffer, &H0)

            Dim i, strByte, intByte
            For i = 15 To 1 Step -2
                strByte = Mid(strHex, i, 2)

                On Error Resume Next
                intByte = CLng("&H" & strByte)

                If (Err.Number <> 0) Then
                    Exit Function
                End If
                On Error GoTo 0

                AddByte(objBuffer, intByte)
            Next

            ' Add closing square bracket
            AddByte(objBuffer, &H5D)
            AddByte(objBuffer, &H0)

            AddQWORDBytes = True
        End Function

        Private Function WriteBinary(FileName, Buf)
            On Error Resume Next

            Dim aBuf, aStream
            aBuf = BuildString(Buf)
            aStream = CreateObject("ADODB.Stream")
            aStream.Type = 1 : aStream.Open
            With CreateObject("ADODB.Stream")
                .Type = 2 : .Open : .WriteText(aBuf)
                ' Copy data from  a text stream to  a binary stream. 
                ' (skip Unicode mark? :FFFE) 
                .Position = 2 : .CopyTo(aStream, UBound(Buf) + 1) : .Close
            End With
            ' At this point aStream.Read give an array of bytes. 

            Err.Clear()

            aStream.SaveToFile(FileName, 2)
            If (Err.Number <> 0) Then
                WriteBinary = PF_WRITEERROR
            End If

            aStream.Close
            aStream = Nothing

            WriteBinary = PF_SUCCESS

            On Error GoTo 0
        End Function

        Private Function BuildString(Buf)
            On Error Resume Next

            Dim I, aBuf(), Size
            Size = UBound(Buf) : ReDim aBuf(Size \ 2)
            For I = 0 To Size - 1 Step 2
                aBuf(I \ 2) = ChrW(Buf(I + 1) * 256 + Buf(I))
            Next
            If I = Size Then aBuf(I \ 2) = ChrW(Buf(I))
            BuildString = Join(aBuf, "")

            On Error GoTo 0
        End Function

        Private Function ReinterpretSignedAsUnsigned(ByVal x)
            If x < 0 Then x = x + 2 ^ 32
            ReinterpretSignedAsUnsigned = x
        End Function

        Private Function UnsignedDecimalStringToHex(ByVal x)
            x = CDbl(x)
            If x > 2 ^ 31 - 1 Then x = x - 2 ^ 32
            UnsignedDecimalStringToHex = Hex(x)
        End Function

        Private Function UnsignedHex(ByVal Number)
            Number = ReinterpretSignedAsUnsigned(Number)
            UnsignedHex = UnsignedDecimalStringToHex(Number)
        End Function
    End Class


End Class
