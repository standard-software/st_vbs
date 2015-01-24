'--------------------------------------------------
'Standard Software Library For VBScript
'
'ModuleName:    StandardSoftwareLibrary.vbs
'--------------------------------------------------
'バージョン     2015/01/21
'--------------------------------------------------

'--------------------------------------------------
'■マーク
'--------------------------------------------------

    '--------------------------------------------------
    '■
    '--------------------------------------------------

    '----------------------------------------
    '◆
    '----------------------------------------

    '------------------------------
    '◇
    '------------------------------

    '--------------------
    '・
    '--------------------

Option Explicit

'--------------------------------------------------
'■定数/型宣言
'--------------------------------------------------

'----------------------------------------
'◆FileSystemObject
'----------------------------------------
Public fso: Set fso = CreateObject("Scripting.FileSystemObject")

'--------------------------------------------------
'■実装
'--------------------------------------------------

'----------------------------------------
'◆テスト
'----------------------------------------
Call test01

Public Sub test01
'    Call WScript.Echo("test")
'    Call WScript.Echo(vbObjectError)
'    Call testAssert
'    Call testCheck
'    Call testOrValue
'    Call testLoadTextFile
'    Call testSaveTextFile
'    Call testFirstStrFirstDelim
'    Call testFirstStrLastDelim
'    Call testLastStrFirstDelim
'    Call testLastStrLastDelim
End Sub

'----------------------------------------
'◆条件判断
'----------------------------------------

'--------------------
'・Assert
'--------------------
'Assert = 主張する
'Err番号 vbObjectError + 1 は
'ユーザー定義エラー番号の1
'動作させると次のようにエラーダイアログが表示される
'   エラー: Messageの内容
'   コード: 80040001
'   ソース: Sub Assert
'--------------------
Public Sub Assert(ByVal Value, ByVal Message)
    If Value = False Then
        Call Err.Raise(vbObjectError + 1, "Sub Assert", Message)
    End If
End Sub

Private Sub testAssert()
    Call Assert(False, "テスト")
End Sub

'--------------------
'・Check
'--------------------
'2つの値を比較して一致しなければ
'メッセージを出す関数
'--------------------
Public Function Check(ByVal A, ByVal B)
    Check = (A = B)
    If Check = False Then
        Call WScript.Echo("A <> B" + vbCrLf + _
            "A = " + CStr(A) + vbCrLf + _
            "B = " + CStr(B))
    End If
End Function

Private Sub testCheck()
    Call Check(10, 20)
End Sub

'--------------------
'・OrValue
'--------------------
'例：If OrValue(ValueA, Array(1, 2, 3)) Then
'--------------------
Public Function OrValue(ByVal Value, ByVal Values)
    OrValue = False
    Dim I
    For I = LBound(Values) To UBound(Values)
        If Value = Values(I) Then
            OrValue = True
            Exit For
        End If
    Next
End Function

Private Sub testOrValue()
    Call Check(True, OrValue(10, Array(20, 30, 40, 10)))
    Call Check(False, OrValue(50, Array(20, 30, 40, 10)))
End Sub

'----------------------------------------
'◆文字列処理
'----------------------------------------

'------------------------------
'◇First/Last Delimiter
'------------------------------

'--------------------
'・FirstStrFirstDelim
'--------------------
Public Function FirstStrFirstDelim(ByVal Value, ByVal Delimiter)
    Dim Result: Result = ""
    Dim Index: Index = InStr(Value, Delimiter)
    If 1 <= Index Then
        Result = Left(Value, Index - 1)
    End If
    FirstStrFirstDelim = Result
End Function

Private Sub testFirstStrFirstDelim
    Call Check("123", FirstStrFirstDelim("123,456", ","))
    Call Check("123", FirstStrFirstDelim("123,456,789", ","))
    Call Check("123", FirstStrFirstDelim("123ttt456", "ttt"))
    Call Check("123", FirstStrFirstDelim("123ttt456", "tt"))
    Call Check("123", FirstStrFirstDelim("123ttt456", "t"))
    Call Check("", FirstStrFirstDelim("123ttt456", ","))
End Sub

'--------------------
'・FirstStrLastDelim
'--------------------
Public Function FirstStrLastDelim(ByVal Value, ByVal Delimiter)
    Dim Result: Result = ""
    Dim Index: Index = InStrRev(Value, Delimiter)
    If 1 <= Index Then
        Result = Left(Value, Index - 1)
    End If
    FirstStrLastDelim = Result
End Function

Private Sub testFirstStrLastDelim
    Call Check("123", FirstStrLastDelim("123,456", ","))
    Call Check("123,456", FirstStrLastDelim("123,456,789", ","))
    Call Check("123", FirstStrLastDelim("123ttt456", "ttt"))
    Call Check("123t", FirstStrLastDelim("123ttt456", "tt"))
    Call Check("123tt", FirstStrLastDelim("123ttt456", "t"))
    Call Check("", FirstStrLastDelim("123ttt456", ","))
End Sub

'--------------------
'・LastStrFirstDelim
'--------------------
Public Function LastStrFirstDelim(ByVal Value, ByVal Delimiter)
    Dim Result: Result = ""
    Dim Index: Index = InStr(Value, Delimiter)
    If 1 <= Index Then
        Result = Mid(Value, Index + Len(Delimiter))
    End If
    LastStrFirstDelim = Result
End Function

Private Sub testLastStrFirstDelim
    Call Check("456", LastStrFirstDelim("123,456", ","))
    Call Check("456,789", LastStrFirstDelim("123,456,789", ","))
    Call Check("456", LastStrFirstDelim("123ttt456", "ttt"))
    Call Check("t456", LastStrFirstDelim("123ttt456", "tt"))
    Call Check("tt456", LastStrFirstDelim("123ttt456", "t"))
    Call Check("", LastStrFirstDelim("123ttt456", ","))
End Sub

'--------------------
'・LastStrLastDelim
'--------------------
Public Function LastStrLastDelim(ByVal Value, ByVal Delimiter)
    Dim Result: Result = ""
    Dim Index: Index = InStrRev(Value, Delimiter)
    If 1 <= Index Then
        Result = Mid(Value, Index + Len(Delimiter))
    End If
    LastStrLastDelim = Result
End Function

Private Sub testLastStrLastDelim
    Call Check("456", LastStrLastDelim("123,456", ","))
    Call Check("789", LastStrLastDelim("123,456,789", ","))
    Call Check("456", LastStrLastDelim("123ttt456", "ttt"))
    Call Check("456", LastStrLastDelim("123ttt456", "tt"))
    Call Check("456", LastStrLastDelim("123ttt456", "t"))
    Call Check("", LastStrLastDelim("123ttt456", ","))
End Sub

'----------------------------------------
'◆テキストファイル読み書き
'----------------------------------------

Public Function CheckEncodeName(ByVal EncodeName)
    CheckEncodeName = OrValue(UCase(EncodeName), Array(_
        "SHIFT_JIS", _
        "UNICODE", "UNICODEFFFE", "UTF-16LE", "UTF-16", _
        "UNICODEFEFF", _
        "UTF-16BE", _
        "UTF-8", _
        "UTF-8N", _
        "ISO-2022-JP", _
        "EUC-JP", _
        "UTF-7") )
End Function

'------------------------------
'◇テキストファイル読込
'エンコード指定は下記の通り
'   エンコード          指定文字
'   ShiftJIS            SHIFT_JIS
'   UTF-16LE BOM有/無   UNICODEFFFE/UNICODE/UTF-16/UTF-16LE
'                       BOMの有無に関わらず読込可能
'   UTF-16BE BOM有り    UNICODEFEFF
'   UTF-16BE BOM無し    UTF-16BE
'   UTF-8 BOM有/無      UTF-8/UTF-8N
'                       BOMの有無に関わらず読込可能
'   JIS                 ISO-2022-JP
'   EUC-JP              EUC-JP
'   UTF-7               UTF-7
'UTF-16LEとUTF-8は、BOMの有無にかかわらず読み込める
'------------------------------
Const StreamTypeEnum_adTypeBinary = 1
Const StreamTypeEnum_adTypeText = 2
Const StreamReadEnum_adReadAll = -1
Const StreamReadEnum_adReadLine = -2
Const SaveOptionsEnum_adSaveCreateOverWrite = 2

Public Function LoadTextFile( _
ByVal TextFilePath, ByVal EncodeName)

    If CheckEncodeName(EncodeName) = False Then
        Call Assert(False, "Error:LoadTextFile")
    End If

    Dim Stream: Set Stream = CreateObject("ADODB.Stream")
    Stream.Type = StreamTypeEnum_adTypeText

    Select Case UCase(EncodeName)
    Case "UTF-8N"
        Stream.Charset = "UTF-8"
    Case Else
        Stream.Charset = EncodeName
    End Select
    Call Stream.Open
    Call Stream.LoadFromFile(TextFilePath)
    LoadTextFile = Stream.ReadText
    Call Stream.Close
End Function

Private Sub testLoadTextFile()
    Call Check("あいうえお", LoadTextFile("TestLoadTextFile\SJIS_File.txt", "Shift_JIS"))
    Call Check("あいうえお", LoadTextFile("TestLoadTextFile\UTF-7_File.txt", "UTF-7"))

    Call Check("あいうえお", LoadTextFile("TestLoadTextFile\UTF-8_File.txt", "UTF-8"))
    Call Check("あいうえお", LoadTextFile("TestLoadTextFile\UTF-8_File.txt", "UTF-8N"))

    Call Check("あいうえお", LoadTextFile("TestLoadTextFile\UTF-8N_File.txt", "UTF-8N"))
    Call Check("あいうえお", LoadTextFile("TestLoadTextFile\UTF-8N_File.txt", "UTF-8"))

    Call Check("あいうえお", LoadTextFile("TestLoadTextFile\UTF-16BEBOM無し_File.txt", "UTF-16BE"))

    Call Check("あいうえお", LoadTextFile("TestLoadTextFile\UTF-16BEBOM有り_File.txt", "UNICODEFEFF"))

    Call Check("あいうえお", LoadTextFile("TestLoadTextFile\UTF-16LEBOM無し_File.txt", "UTF-16"))
    Call Check("あいうえお", LoadTextFile("TestLoadTextFile\UTF-16LEBOM無し_File.txt", "UTF-16LE"))
    Call Check("あいうえお", LoadTextFile("TestLoadTextFile\UTF-16LEBOM無し_File.txt", "UNICODE"))
    Call Check("あいうえお", LoadTextFile("TestLoadTextFile\UTF-16LEBOM無し_File.txt", "UNICODEFFFE"))

    Call Check("あいうえお", LoadTextFile("TestLoadTextFile\UTF-16LEBOM有り_File.txt", "UTF-16"))
    Call Check("あいうえお", LoadTextFile("TestLoadTextFile\UTF-16LEBOM有り_File.txt", "UTF-16LE"))
    Call Check("あいうえお", LoadTextFile("TestLoadTextFile\UTF-16LEBOM有り_File.txt", "UNICODE"))
    Call Check("あいうえお", LoadTextFile("TestLoadTextFile\UTF-16LEBOM有り_File.txt", "UNICODEFFFE"))
End Sub

'------------------------------
'◇テキストファイル保存
'エンコード指定は下記の通り
'   エンコード          指定文字
'   ShiftJIS            SHIFT_JIS
'   UTF-16LE BOM有り    UNICODEFFFE/UNICODE/UTF-16
'   UTF-16LE BOM無し    UTF-16LE
'   UTF-16BE BOM有り    UNICODEFEFF
'   UTF-16BE BOM無し    UTF-16BE
'   UTF-8 BOM有り       UTF-8
'   UTF-8 BOM無し       UTF-8N
'   JIS                 ISO-2022-JP
'   EUC-JP              EUC-JP
'   UTF-7               UTF-7
'UTF-16LEとUTF-8はそのままだとBOM有りになるので
'BON無し指定の場合は特殊処理をしている
'------------------------------
Public Sub SaveTextFile(ByVal Text, _
ByVal TextFilePath, ByVal EncodeName)
    If CheckEncodeName(EncodeName) = False Then
        Call Assert(False, "Error:SaveTextFile")
    End If

    Dim Stream: Set Stream = CreateObject("ADODB.Stream")
    Stream.Type = StreamTypeEnum_adTypeText

    Select Case UCase(EncodeName)
    Case "UTF-8N"
        Stream.Charset = "UTF-8"
    Case Else
        Stream.Charset = EncodeName
    End Select
    Call Stream.Open
    Call Stream.WriteText(Text)
    Dim ByteData
    Select Case UCase(EncodeName)
    Case "UTF-16LE"
        Stream.Position = 0
        Stream.Type = StreamTypeEnum_adTypeBinary
        Stream.Position = 2
        ByteData = Stream.Read
        Stream.Position = 0
        Call Stream.Write(ByteData)
        Call Stream.SetEOS
    Case "UTF-8N"
        Stream.Position = 0
        Stream.Type = StreamTypeEnum_adTypeBinary
        Stream.Position = 3
        ByteData = Stream.Read
        Stream.Position = 0
        Call Stream.Write(ByteData)
        Call Stream.SetEOS
    End Select
    Call Stream.SaveToFile(TextFilePath, _
        SaveOptionsEnum_adSaveCreateOverWrite)
    Call Stream.Close
End Sub

Private Sub testSaveTextFile()
    Call SaveTextFile("123ABCあいうえお", "TestSaveTextFile\SJIS_File.txt", "Shift_JIS")
    Call SaveTextFile("123ABCあいうえお", "TestSaveTextFile\UTF-7_File.txt", "UTF-7")
    Call SaveTextFile("123ABCあいうえお", "TestSaveTextFile\UTF-8_File.txt", "UTF-8")
    Call SaveTextFile("123ABCあいうえお", "TestSaveTextFile\UTF-8N_File.txt", "UTF-8N")

    Call SaveTextFile("123ABCあいうえお", "TestSaveTextFile\UTF-16BEBOM無し_File.txt", "UTF-16BE")
    Call SaveTextFile("123ABCあいうえお", "TestSaveTextFile\UTF-16BEBOM有り_File.txt", "UNICODEFEFF")
    Call SaveTextFile("123ABCあいうえお", "TestSaveTextFile\UTF-16LEBOM無し_File.txt", "UTF-16LE")
    Call SaveTextFile("123ABCあいうえお", "TestSaveTextFile\UTF-16LEBOM有り1_File.txt", "UNICODEFFFE")
    Call SaveTextFile("123ABCあいうえお", "TestSaveTextFile\UTF-16LEBOM有り2_File.txt", "UNICODE")
    Call SaveTextFile("123ABCあいうえお", "TestSaveTextFile\UTF-16LEBOM有り3_File.txt", "UTF-16")
End Sub

'--------------------------------------------------
'■履歴
'◇ ver 2015/01/20
'・ 作成
'・ Assert/Check/OrValue
'・ fso
'・ TestStandardSoftwareLibrary.vbs/Include
'・ LoadTextFile/SaveTextFile
'◇ ver 2015/01/21
'・ FirstStrFirstDelim/FirstStrLastDelim
'   /LastStrFirstDelim/LastStrLastDelim
'--------------------------------------------------
