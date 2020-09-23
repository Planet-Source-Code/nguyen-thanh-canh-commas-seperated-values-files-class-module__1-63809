VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NewCSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private RecordStr() As String

'local variable(s) to hold property value(s)
Private mvarFileName As String 'local copy
'local variable(s) to hold property value(s)
Private mvarDelimiter As String 'local copy
'local variable(s) to hold property value(s)
Private mvarCurrentRecord As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarFields As String 'local copy
'local variable(s) to hold property value(s)
Private mvarFieldNo As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarRecordCount As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarRecords As String 'local copy
Public Header As Boolean
'local variable(s) to hold property value(s)
Private mvarHeaderStr As String 'local copy
'local variable(s) to hold property value(s)
Private mvarFieldsInRecords As String 'local copy
'local variable(s) to hold property value(s)
Private mvarBOF As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarEOF As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarBOfF As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarEOfF As Boolean 'local copy

Public Sub WriteFileAs(NewFileName As String)
Dim FileNum As Integer
Dim TempRecord As Long

    'Ghi du lieu tu mang ra tap tin
    On Error GoTo EndOfWriteFileAs
    FileNum = FreeFile
    Open NewFileName For Output Lock Read Write As FileNum
    TempRecord = IIf(Header, 0, 1)
    Do While TempRecord <= RecordCount
        Print #FileNum, RecordStr(TempRecord)
        TempRecord = TempRecord + 1
    Loop
    Close FileNum
EndOfWriteFileAs:
End Sub

Public Property Let EOfF(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EOfF = 5
    mvarEOfF = vData
End Property


Public Property Get EOfF() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EOfF
    mvarEOfF = (CurrentRecord > RecordCount Or RecordCount = 0)

    EOfF = mvarEOfF
End Property



Public Property Let BOfF(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BOfF = 5
    mvarBOfF = vData
End Property


Public Property Get BOfF() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BOfF
    mvarBOfF = (CurrentRecord <= 0)

    BOfF = mvarBOfF
End Property


Public Sub Move(Index As Long)
  CurrentRecord = Index
End Sub


Public Property Let FieldsInRecords(RecordIndex As Long, FieldIndex As Integer, ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FieldsInRecords = 5
Dim TempRecord As Long
    
    mvarFieldsInRecords = vData
    TempRecord = CurrentRecord
    CurrentRecord = RecordIndex
    Fields(FieldIndex) = mvarFieldsInRecords
    CurrentRecord = TempRecord
End Property

Public Property Get FieldsInRecords(RecordIndex As Long, FieldIndex As Integer) As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FieldsInRecords
Dim TempRecord As Long
    
    TempRecord = CurrentRecord
    CurrentRecord = RecordIndex
    mvarFieldsInRecords = Fields(FieldIndex)
    CurrentRecord = TempRecord
    FieldsInRecords = mvarFieldsInRecords
End Property

Public Property Let Fields(Index As Integer, ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Fields = 5
Dim TempRecord As Long
Dim FieldArr() As String
Dim FieldNo As Integer
    
    mvarFields = vData
    On Error GoTo EndOfSetFields
    If (RecordCount > 0) And (CurrentRecord <= RecordCount) Then
        FieldArr = Split(RecordStr(CurrentRecord), Delimiter)
        If Index > UBound(FieldArr) Then ReDim Preserve FieldArr(Index)
        FieldArr(Index) = mvarFields
    End If
    For FieldNo = 0 To UBound(FieldArr)
        If FieldNo = 0 Then
            RecordStr(CurrentRecord) = FieldArr(FieldNo)
        Else: RecordStr(CurrentRecord) = RecordStr(CurrentRecord) & Delimiter & FieldArr(FieldNo)
        End If
    Next FieldNo
EndOfSetFields:

End Property

Public Property Get Fields(Index As Integer) As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Fields
Dim FieldArr() As String

    Fields = mvarFields
    On Error GoTo EndOfGetFields
    If (RecordCount > 0) And (CurrentRecord <= RecordCount) Then
        FieldArr = Split(RecordStr(CurrentRecord), Delimiter)
    End If
    If (Index <= UBound(FieldArr)) Then Fields = FieldArr(Index)
EndOfGetFields:

End Property

Public Property Let Records(Index As Long, ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Records = 5
    mvarRecords = vData
    If RecordCount < Index Then ReDim Preserve RecordStr(Index)
    RecordStr(Index) = mvarRecords
End Property

Public Property Get Records(Index As Long) As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Records
  If Index <= RecordCount Then
    mvarRecords = RecordStr(Index)
  Else: mvarRecords = ""
End If
    Records = mvarRecords
End Property

Public Property Let HeaderStr(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HeaderStr = 5
    mvarHeaderStr = vData
    Records(0) = mvarHeaderStr
End Property

Public Property Get HeaderStr() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HeaderStr
    mvarHeaderStr = Records(0)
    HeaderStr = mvarHeaderStr
End Property

Public Property Get RecordCount() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RecordCount
    RecordCount = UBound(RecordStr)
End Property

Public Property Get FieldsCount() As Integer
Dim FieldArr() As String

    On Error GoTo EndOfFieldsCount
    If (RecordCount > 0) And (CurrentRecord <= RecordCount) Then
        FieldArr = Split(RecordStr(CurrentRecord), Delimiter)
    End If
    FieldsCount = UBound(FieldArr) + 1
EndOfFieldsCount:
End Property

Public Property Let CurrentRecord(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CurrentRecord = 5
    mvarCurrentRecord = vData
End Property

Public Property Get CurrentRecord() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CurrentRecord
    CurrentRecord = mvarCurrentRecord
End Property

Public Property Let Delimiter(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Delimiter = 5
    mvarDelimiter = vData
End Property

Public Property Get Delimiter() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Delimiter
    Delimiter = mvarDelimiter
End Property

Public Property Let FileName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileName = 5
    mvarFileName = vData
End Property

Public Property Get FileName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileName
    FileName = mvarFileName
End Property


Public Sub MovePrevious()
    If Not BOfF Then CurrentRecord = CurrentRecord - 1
End Sub

Public Sub MoveNext()
    If Not EOfF Then CurrentRecord = CurrentRecord + 1
End Sub

Public Sub MoveLast()
    CurrentRecord = RecordCount
End Sub

Public Sub MoveFirst()
    If RecordCount > 0 Then
      CurrentRecord = 1
    Else: CurrentRecord = 0
  End If
End Sub

Public Sub OpenFile()
Dim FileNum As Integer
Dim TempRecord As Long

    'Doc du lieu tu tap tin va luu vao mang
    On Error GoTo EndOfOpenFile:
    ReDim RecordStr(0)
    TempRecord = IIf(Header, -1, 0)
    FileNum = FreeFile
    Open FileName For Input As FileNum
    Do While Not EOF(FileNum)
        TempRecord = TempRecord + 1
        ReDim Preserve RecordStr(TempRecord)
        Line Input #FileNum, RecordStr(TempRecord)
    Loop
    Close FileNum
EndOfOpenFile:
End Sub

Public Sub WriteFile()
Dim FileNum As Integer
Dim TempRecord As Long

    'Ghi du lieu tu mang ra tap tin
    On Error GoTo EndOfWriteFile
    FileNum = FreeFile
    Open FileName For Output Lock Read Write As FileNum
    TempRecord = IIf(Header, 0, 1)
    Do While TempRecord <= RecordCount
        Print #FileNum, RecordStr(TempRecord)
        TempRecord = TempRecord + 1
    Loop
    Close FileNum
EndOfWriteFile:
End Sub

Public Sub Add(AddStr As String)
    
    ReDim Preserve RecordStr(RecordCount + 1)
    RecordStr(RecordCount) = AddStr
    CurrentRecord = RecordCount
End Sub

Public Sub Insert(InsertStr As String)
Dim TempRecord As Long

    'Doc du lieu tu tap tin va luu vao mang
    ReDim Preserve RecordStr(RecordCount + 1)
    TempRecord = RecordCount
    Do While TempRecord >= CurrentRecord
        If TempRecord > CurrentRecord Then
            RecordStr(TempRecord) = RecordStr(TempRecord - 1)
        Else
            RecordStr(TempRecord) = InsertStr
        End If
        TempRecord = TempRecord - 1
    Loop
End Sub

Public Sub Delete()
Dim TempRecord As Long
    
    'Xoa record hien tai ra khoi mang
    TempRecord = CurrentRecord
    Do While TempRecord < RecordCount
        RecordStr(TempRecord) = RecordStr(TempRecord + 1)
        TempRecord = TempRecord + 1
    Loop
    If RecordCount > 0 Then ReDim Preserve RecordStr(RecordCount - 1)
End Sub

Public Sub Sort(Keys As Integer, Formats As String, Ascending As Boolean)
Dim Temp As String, i As Long, j As Long
Dim Reversed As Boolean

    'BoubleSort
    For i = 1 To RecordCount - 1
        For j = 1 To RecordCount - i
            If Formats = "" Then
                Select Case Ascending
                    Case True: If FieldsInRecords(j, Keys) > FieldsInRecords(j + 1, Keys) Then Reversed = True
                    Case False: If FieldsInRecords(j, Keys) < FieldsInRecords(j + 1, Keys) Then Reversed = True
                End Select
            Else
                 Select Case Ascending
                    Case True: If Format(FieldsInRecords(j, Keys), Formats) > Format(FieldsInRecords(j + 1, Keys), Formats) Then Reversed = True
                    Case False: If Format(FieldsInRecords(j, Keys), Formats) < Format(FieldsInRecords(j + 1, Keys), Formats) Then Reversed = True
                End Select
           End If
           If Reversed Then
                Temp = RecordStr(j)
                RecordStr(j) = RecordStr(j + 1)
                RecordStr(j + 1) = Temp
                Reversed = False
            End If
        Next j
    Next i
End Sub

Public Function FindFirst(StartRecord As Integer, Keys As Integer, FindString As String, Operator As String) As String
Dim TempRecord As Long

    MoveFirst
    Do While Not EOfF
        If Fields(Keys) = FindString Then
            FindFirst = Records(CurrentRecord)
            Exit Do
        Else: MoveNext
        End If
    Loop
End Function

Public Function FindNext(Keys As Integer, FindString As String) As String
    MoveNext
    Do While Not BOfF
        If Fields(Keys) = FindString Then
            FindNext = Records(CurrentRecord)
        Else: MoveNext
        End If
    Loop
End Function

Public Sub Class_Initialize()
    ReDim RecordStr(0)
End Sub
