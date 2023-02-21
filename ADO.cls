VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "adoUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

#If False Then
    Private ConnectClientSide
    Private ConnectDataShape
    Private ConnectServerSide
#End If

Public Enum ADOConnectType
    ConnectClientSide
    ConnectDataShape
    ConnectServerSide
End Enum

Private Const mModName  As String = "ADO Utilities"
Private Const sqlDSN    As String = "DRIVER={SQL SERVER};Provider=SQLOLEDB.1;Server=[+SERVER+];Database=[+DATABASE+];UID=[+UID+];pwd=[+PWD+];"
Private Const mdbDSN    As String = "Provider=Microsoft.Jet.OLEDB.4.0;Password=[+PWD+];User ID=[+UID+];Data Source=[+SERVER+][+DATABASE+];Mode=Share Deny None;Extended Properties=;Jet OLEDB:System database=;Jet OLEDB:Registry Path=;Jet OLEDB:Database Password=;Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password=;Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False"

Private mConnectString        As String
Private mLastError      As String
Private mLastSQL        As String
Private mWinHandle      As Long

Public WithEvents cnADO As ADODB.Connection     'Standard Connection
Attribute cnADO.VB_VarHelpID = -1
Public WithEvents cnADX As ADODB.Connection     'DataShape Connection
Attribute cnADX.VB_VarHelpID = -1
Public WithEvents cnSVR As ADODB.Connection     'Server Side Connection
Attribute cnSVR.VB_VarHelpID = -1

Public SmartSQL As SmartSQL

'####################
'# Class Properties #
'####################

Public Property Let ConnectString(vData As String)
    mConnectString = vData
End Property
Public Property Get ConnectString() As String
    ConnectString = mConnectString
End Property

Public Property Get DBLocation() As String

    Static mDBLocation  As String
    Dim lTemp           As Variant

    If Len(mDBLocation) = 0 Then
        Dim x As Integer
        lTemp = Split(ConnectString, ";")
        If IsArray(lTemp) Then
            For x = LBound(lTemp) To UBound(lTemp)
                If Left(UCase(lTemp(x)), 12) = "DATA SOURCE=" Then
                    mDBLocation = Replace(lTemp(x), "DATA SOURCE=", "", , , vbTextCompare)
                    mDBLocation = Replace(mDBLocation, ".mdb", "")
                    Exit For
                End If
            Next x
        End If
    End If

    DBLocation = mDBLocation

End Property

Public Property Get DBPath() As String
    'TODO: - DBPath = oUtils.FilePart(DBLocation, FilePathOnly)
End Property

Public Property Get DBName() As String
    'TODO - DBName = oUtils.FilePart(DBLocation, FileNameOnly)
End Property

Private Property Get IsSQL() As Boolean
    IsSQL = SmartSQL.SQLType = SQL_TYPE_ANSI
End Property

Public Property Let LastError(vData As String)
    If Len(mLastError) = 0 Then mLastError = vData
End Property
Public Property Get LastError() As String
    LastError = mLastError
End Property

Public Property Let LastSQL(vData As String)
    mLastSQL = vData
End Property
Public Property Get LastSQL() As String
    LastSQL = mLastSQL
End Property

Public Property Get MySQL() As String
    MySQL = SmartSQL.SQL
    SmartSQL.Reset
End Property

Public Property Let WinHandle(vData As Long)
    mWinHandle = vData
End Property
Public Property Get WinHandle() As Long
    WinHandle = mWinHandle
End Property

'####################
'# Public Functions #
'####################

Public Function Connect(ConnectType As ADOConnectType, Optional wHandle As Long) As Boolean

    On Error GoTo LocalError

    If Len(ConnectString) = 0 Then Exit Function

    Select Case ConnectType
        Case ConnectClientSide                      ' Client Side Connection
            If cnADO Is Nothing Then Set cnADO = New ADODB.Connection
            With cnADO
                If .State = adStateOpen Then        ' Already Connected
                    Connect = True
                Else                                ' Try and establish a connection
                    .ConnectionString = ConnectString
                    .CursorLocation = adUseClient
                    .Open
                    Connect = True
                End If
                If .State Then .Errors.Clear
            End With

        Case ConnectDataShape                       ' Data Shape Connection
            If cnADX Is Nothing Then Set cnADX = New ADODB.Connection
            With cnADX
                If .State = adStateOpen Then        ' Already Connected
                    Connect = True
                Else                                ' Try and establish a connection
                    If IsSQL Then                   ' Use MS SQL Connection
                        .ConnectionString = Replace(ConnectString, "Provider=SQLOLEDB.1;", "Data Provider=MSDASQL;", 1, -1, 1)
                    Else                            ' Use MS Access Connection
                        .ConnectionString = Replace(ConnectString, "Provider=", "Data Provider=", 1, -1, 1)
                    End If
                    .CursorLocation = adUseServer
                    .Provider = "MSDataShape"
                    .Open
                    Connect = True
                End If
                If .State Then .Errors.Clear
            End With

        Case ConnectServerSide
            If cnSVR Is Nothing Then Set cnSVR = New ADODB.Connection
            With cnSVR
                If .State = adStateOpen Then        ' Already Connected
                    Connect = True
                Else                                ' Try and establish a connection
                    .ConnectionString = ConnectString
                    .CursorLocation = adUseServer
                    .Open
                    Connect = True
                End If
                If .State Then .Errors.Clear
            End With

        Case Else:  'mmm....

    End Select

Exit Function

LocalError:
    LastError = Err.Source & " (" & Err.Number & ") " & Err.Description
End Function

Private Function ConObject(ConnectType As ADOConnectType) As ADODB.Connection
    Select Case ConnectType
        Case ConnectClientSide:   If Connect(ConnectType) = True Then Set ConObject = cnADO
        Case ConnectDataShape:    If Connect(ConnectType) = True Then Set ConObject = cnADX
        Case ConnectServerSide:   If Connect(ConnectType) = True Then Set ConObject = cnSVR
        Case Else:  'mmm....
    End Select
End Function

Public Function CloseALL() As Boolean

    On Error GoTo ErrHandler

    If Not cnADO Is Nothing Then If cnADO.State Then cnADO.Close
    If Not cnADX Is Nothing Then If cnADX.State Then cnADX.Close
    If Not cnSVR Is Nothing Then If cnSVR.State Then cnSVR.Close

    Set cnADO = Nothing
    Set cnADX = Nothing
    Set cnSVR = Nothing

    CloseALL = True

ErrHandler:
End Function

Public Function Clone(ByVal oRS As ADODB.Recordset, Optional LockType As LockTypeEnum = adLockUnspecified) As ADODB.Recordset

    On Error GoTo LocalError

    Set Clone = New ADODB.Recordset

    'This is very efficient
    Dim stm As ADODB.Stream
    Set stm = New ADODB.Stream

    oRS.Fields.Append "Child", adVariant
    oRS.Save stm
    Clone.Open stm, , , LockType

Exit Function

LocalError:
    LastError = Err.Source & " (" & Err.Number & ") " & Err.Description
End Function

Public Function CloneFieldStructure(ByVal oRS) As ADODB.Recordset

    On Error GoTo LocalError

    Dim fld As ADODB.Field

    Set CloneFieldStructure = New ADODB.Recordset

    ' create a set of fields with same attributes
    For Each fld In oRS.Fields
        CloneFieldStructure.Fields.Append fld.Name, fld.Type, fld.DefinedSize, fld.Attributes
        'special handling for data types with numeric scale & precision
        Select Case fld.Type
            Case adNumeric, adDecimal
                With CloneFieldStructure
                    .Fields(.Fields.Count - 1).Precision = fld.Precision
                    .Fields(.Fields.Count - 1).NumericScale = fld.NumericScale
                End With
        End Select
    Next

Exit Function

LocalError:
    LastError = Err.Source & " (" & Err.Number & ") " & Err.Description
End Function

Public Function DataShape(ByVal tblParent, _
                          ByVal tblChild, _
                          ByVal fldParent, _
                          ByVal fldChild, _
                          Optional ordParent = "", _
                          Optional ordChild = "", _
                          Optional WhereParent = "", _
                          Optional WhereChild = "", _
                          Optional ParentFields = "", _
                          Optional ChildFields = "", _
                          Optional MaxRecords = 0, _
                          Optional LockType As ADODB.LockTypeEnum = adLockReadOnly, _
                          Optional CursorType As ADODB.CursorTypeEnum = adOpenStatic, _
                          Optional Disconnected As Boolean = True) As ADODB.Recordset
    '=========================================================
    'This function will return a SHAPEd RecordSet
    'Assumptions:
    '
    'tblParent      = Valid Table in the Database   - String \ Parent Table
    'tblChild       = Valid Table in the Database   - String / Child  Table
    '
    'fldParent      = Valid Field in Parent Table   - String \ relate this field
    'fldChild       = Valid Field in Child Table    - String / ..to this field
    '
    'ordParent      = Valid Field in Parent Table   - String \ ordering
    'ordChild       = Valid Field in Child Table    - String /
    '
    'WhereParent    = Valid SQL Where Clause        - Variant (Optional)
    'WhereChild     = Valid SQL Where Clause        - Variant (Optional)
    '
    'ParentFields   = Specific Fields to return     - String (pipe delimitered)
    'ChildFields    = Specific Fields to return     - String (pipe delimitered)
    'MaxRecords     = Specify Maximum Child Records - Long (0 = ALL)
    '=========================================================
    If Len(WhereChild) = 0 Then WhereChild = ""
    If Len(ParentFields) = 0 Then ParentFields = "*"
    If Len(ChildFields) = 0 Then ChildFields = "*"
    If Len(MaxRecords) = 0 Then MaxRecords = 0
    If Len(LockType) = 0 Then LockType = adLockReadOnly
    If Len(CursorType) = 0 Then CursorType = adOpenStatic
    If Len(Disconnected) = 0 Then Disconnected = True

    Dim lSQL As String
    Dim pSQL As String
    Dim cSQL As String
    Dim pWhere As String
    Dim cWhere As String
    Dim pOrder As String
    Dim cOrder As String

    On Error GoTo LocalError
    'Define the SQL Statement
    lSQL = ""
    ParentFields = CStr(Replace(ParentFields, "|", ", "))
    ChildFields = CStr(Replace(ChildFields, "|", ", "))
    pWhere = CStr(WhereParent)
    cWhere = CStr(WhereChild)
    pOrder = CStr(ordParent)
    cOrder = CStr(ordChild)
    If pWhere <> "" Then pWhere = " WHERE " & pWhere
    If cWhere <> "" Then cWhere = " WHERE " & cWhere
    If pOrder <> "" Then pOrder = " ORDER By " & pOrder
    If cOrder <> "" Then cOrder = " ORDER By " & cOrder
    'Define Parent SQL Statement
    pSQL = ""
    If MaxRecords > 0 Then
        If IsSQL Then
            pSQL = pSQL & "{SET ROWCOUNT " & MaxRecords & " SELECT [@PARENTFIELDS]"
        Else
            pSQL = pSQL & "{SELECT TOP " & MaxRecords & " [@PARENTFIELDS]"
        End If
    Else
        pSQL = pSQL & "{SELECT " & "[@PARENTFIELDS]"
    End If
    pSQL = pSQL & " FROM [@PARENT]"
    pSQL = pSQL & " [@WHEREPARENT]"
    pSQL = pSQL & " [@ORDPARENT]} "
    'Substitute for actual values
    pSQL = Replace(pSQL, "[@PARENTFIELDS]", ParentFields)
    pSQL = Replace(pSQL, "[@PARENT]", tblParent)
    pSQL = Replace(pSQL, "[@WHEREPARENT]", pWhere)
    pSQL = Replace(pSQL, "[@ORDPARENT]", pOrder)
    'Define Child SQL Statement
    cSQL = ""
    cSQL = cSQL & "{SELECT " & "[@CHILDFIELDS]"
    cSQL = cSQL & " FROM [@CHILD]"
    cSQL = cSQL & " [@WHERECHILD]"
    cSQL = cSQL & " [@ORDCHILD]} "
    'Substitute for actual values
    cSQL = Replace(cSQL, "[@CHILDFIELDS]", ChildFields)
    cSQL = Replace(cSQL, "[@CHILD]", tblChild)
    cSQL = Replace(cSQL, "[@WHERECHILD]", cWhere)
    cSQL = Replace(cSQL, "[@ORDCHILD]", cOrder)

    'Define Parent Properties
    lSQL = "SHAPE " & pSQL & vbCrLf
    'Define Child Properties
    lSQL = lSQL & "APPEND (" & cSQL & " RELATE " & fldParent & " TO " & fldChild & ") AS ChildItems"
    'TODO: - lSQL = TrimALL(lSQL)

    'Get the data
    LastSQL = lSQL

    Set DataShape = New ADODB.Recordset

    With DataShape
        .CursorType = CursorType
        .LockType = LockType
        .Source = lSQL
        .ActiveConnection = ConObject(ConnectDataShape)
        .Open
        If Disconnected Then Set .ActiveConnection = Nothing
    End With

Exit Function

LocalError:
    LastError = Err.Source & " (" & Err.Number & ") " & Err.Description
End Function

Public Function EmptyRS(ByVal oRS) As Boolean

    On Error Resume Next
    'Checks for an EMPTY RecordSet
    EmptyRS = True
    If Not oRS Is Nothing Then
        EmptyRS = ((oRS.BOF = True) And (oRS.EOF = True))
    End If

End Function

Public Sub ErrorClear()
    LastError = ""
End Sub

Public Function Execute(SQL) As Boolean

    On Error GoTo LocalError

    If Connect(ConnectServerSide) Then
        LastSQL = CStr(SQL)
        With cnSVR
            .BeginTrans
            .Execute CStr(SQL)
            .CommitTrans
        End With
    End If
    Execute = True

Exit Function

LocalError:
    LastError = Err.Source & " (" & Err.Number & ") " & Err.Description

    If cnSVR.State = adStateOpen Then
        cnSVR.RollbackTrans
    End If

End Function

'Tested with ACCESS 2000 and SQL7.0 using ADO2.5
Public Function ExecuteID(SQL) As Long

    On Error GoTo LocalError

    Dim oRS     As New ADODB.Recordset

    With oRS
        'Prepare the RecordSet
        .CursorLocation = adUseServer
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "SELECT @@IDENTITY"
    End With

    If Connect(ConnectServerSide) Then
        With cnSVR  'NB: Server Side Connection
            .ConnectionString = ConnectString
            .CursorLocation = adUseServer
            .Open
            LastSQL = CStr(SQL)
            .BeginTrans
            .Execute CStr(SQL), , adCmdText + adExecuteNoRecords
            .CommitTrans
            oRS.ActiveConnection = cnSVR
            oRS.Open , , , , adCmdText
            ExecuteID = oRS(0).Value
            oRS.Close
            .Close
        End With
    End If

ExitHere:
    If oRS.State = adStateOpen Then oRS.Close
    Set oRS = Nothing
Exit Function

LocalError:
    LastError = Err.Source & " (" & Err.Number & ") " & Err.Description
    On Error Resume Next
    If cnSVR.State = adStateOpen Then cnSVR.RollbackTrans
    Resume ExitHere
End Function

Public Function GetCount(TableName As Variant, Optional WhereClause As Variant = "") As Long

    Dim oRS As ADODB.Recordset
    Dim lSQL As String

    On Error GoTo LocalError

    TableName = CStr(TableName)
    WhereClause = CStr(WhereClause)
    GetCount = -1
    GetCount = 0

    If WhereClause <> "" Then
        lSQL = "Select COUNT (*) FROM " & TableName & " WHERE " & WhereClause
    Else
        lSQL = "Select COUNT (*) FROM " & TableName
    End If

    If Connect(ConnectServerSide) Then
        LastSQL = lSQL
        Set oRS = New ADODB.Recordset
        With cnSVR
            Set oRS = .Execute(lSQL)
            GetCount = oRS.Fields(0).Value
            oRS.Close
        End With
    End If

    If Not oRS Is Nothing Then Set oRS = Nothing

Exit Function

LocalError:
    LastError = Err.Source & " (" & Err.Number & ") " & Err.Description
    If Not oRS Is Nothing Then Set oRS = Nothing
End Function

Public Function GetIncrement(TableName As String, FieldName As String, KeyField As String, KeyID As Variant) As Long

    'This function gets a Field Value
    '   Increments its Value by ONE and saves the result

    Dim oRS     As ADODB.Recordset
    Dim NextNum As Long
    Dim numSQL  As String
    Dim updSQL  As String
    Dim Started As Date
    
    If cnSVR Is Nothing Then Set cnSVR = New ADODB.Connection
    Set oRS = New ADODB.Recordset

    With SmartSQL
        .StatementType = TYPE_SELECT
        .AddTable TableName
        .AddField FieldName
        .AddSimpleWhereClause KeyField, KeyID
        numSQL = MySQL
    End With

    With oRS    'Prepare the RecordSet
        .CursorLocation = adUseServer
        .CursorType = adOpenDynamic
        .LockType = adLockPessimistic
        .Source = numSQL
    End With

    If Connect(ConnectServerSide) Then
        With cnSVR  'NB: Server Side Connection
            .ConnectionString = ConnectString
            .CursorLocation = adUseServer
            .Open
            .BeginTrans
            Set oRS.ActiveConnection = cnSVR
            oRS.Open
            NextNum = oRS(0) + 1
            oRS(0) = NextNum
            oRS.Update
            .CommitTrans
            oRS.Close
            .Close
        End With
        GetIncrement = NextNum
    End If

ExitHere:
    Set oRS = Nothing
Exit Function

LocalError:
    LastError = Err.Source & " (" & Err.Number & ") " & Err.Description
    On Error Resume Next
    If cnSVR.State = adStateOpen Then cnSVR.RollbackTrans
    If oRS.State = adStateOpen Then oRS.Close
    Resume ExitHere
End Function

Public Function GetRS(SQL As Variant, Optional LockType As ADODB.LockTypeEnum = adLockReadOnly, Optional CursorType As ADODB.CursorTypeEnum = adOpenStatic, Optional Disconnected As Boolean = True, Optional ConnectType As ADOConnectType = ConnectClientSide) As ADODB.Recordset

    On Error GoTo LocalError

    LastSQL = CStr(SQL)

    Set GetRS = New ADODB.Recordset

    With GetRS
        .LockType = LockType
        .CursorType = CursorType
        .Source = CStr(SQL)
        .ActiveConnection = ConObject(ConnectType)
        .Open
        If Disconnected Then Set .ActiveConnection = Nothing
    End With

Exit Function

LocalError:
    LastError = Err.Source & " (" & Err.Number & ") " & Err.Description
End Function

Public Function Find(oRS As ADODB.Recordset, ThisField As ADODB.Field, ThisItem As Variant, Optional Operator As CLAUSE_OPERATOR = CLAUSE_EQUALS)

    On Error GoTo LocalError

    Dim lFindString As String

    If EmptyRS(oRS) Then Exit Function

    SmartSQL.Reset
    lFindString = SmartSQL.AddSimpleWhereClause(ThisField.Name, ThisItem, , Operator)
    SmartSQL.Reset

    With oRS
        'Try forward First
        .Find lFindString, , adSearchForward
        If .EOF Then    'Try backward next
            .Find lFindString, , adSearchBackward
        End If
        Find = Not .BOF 'Success or Failure
    End With

LocalError:
End Function

Public Function ImageLoad(oPictureControl As Object, adoRS As ADODB.Recordset, sFieldName As String) As Boolean
'    Dim oPict       As StdPicture
'    Dim sDir        As String
'    Dim sTempFile   As String
'    Dim iFileNum    As Integer
'    Dim lFileLength As Long
'    Dim abBytes()   As Byte
'    Dim iCtr        As Integer
'
'    On Error GoTo ErrHandler
'
'    sTempFile = oUtils.MyComputer.Directory(dirTEMP) & "tmpImage"
'    If oUtils.FileExists(sTempFile) Then Kill sTempFile
'
'    iFileNum = FreeFile
'    Open sTempFile For Binary As #iFileNum
'        lFileLength = LenB(adoRS(sFieldName))
'        abBytes = adoRS(sFieldName).GetChunk(lFileLength)
'        Put #iFileNum, , abBytes()
'    Close #iFileNum
'
'    oPictureControl.Picture = LoadPicture(sTempFile)
'
'    Kill sTempFile
'    ImageLoad = True
'
'Exit Function
'ErrHandler:
'    ImageLoad = False
'    Debug.Print Err.Description
End Function

Public Function ImageSave(oPictureControl As Object, adoRS As ADODB.Recordset, sFieldName As String) As Boolean

'    Dim oPict       As StdPicture
'    Dim sDir        As String
'    Dim sTempFile   As String
'    Dim iFileNum    As Integer
'    Dim lFileLength As Long
'    Dim abBytes()   As Byte
'    Dim iCtr        As Integer
'
'    On Error GoTo ErrHandler
'
'    Set oPict = oPictureControl.Picture
'
'    If oPict Is Nothing Then
'        ImageSave = False
'        Exit Function
'    End If
'
'    'Save picture to temp file
'    sTempFile = oUtils.MyComputer.Directory(dirTEMP) & "tmpImage"
'    If oUtils.FileExists(sTempFile) Then Kill sTempFile
'    SavePicture oPict, sTempFile
'
'    'read file contents to byte array
'    iFileNum = FreeFile
'    Open sTempFile For Binary Access Read As #iFileNum
'        lFileLength = LOF(iFileNum)
'        ReDim abBytes(lFileLength)
'        Get #iFileNum, , abBytes()
'        'put byte array contents into db field
'        adoRS.Fields(sFieldName).AppendChunk abBytes()
'    Close #iFileNum
'
'    'Don't return false if file can't be deleted
'    On Error Resume Next
'    Kill sTempFile
'    ImageSave = True
'
'Exit Function
'ErrHandler:
'    ImageSave = False
'    Debug.Print Err.Description
End Function

Public Function Mask(pDataType As ADODB.DataTypeEnum, pDataValue As Variant) As String

    On Error Resume Next

    Select Case pDataType
        Case adChapter, adArray 'Nothing we can do with this
        Case adBSTR, adChar, adLongVarChar, adLongVarWChar, adVarChar, adVarWChar, adWChar
            'Add single quotes around string values and remove single quotes from within the string.
            Mask = "'" & sqlEncode(CStr(pDataValue)) & "'"
        Case adDBDate, adDBTime, adDBTime, adDBTimeStamp, adFileTime
            'Add Hash Marks around dates/times.
            Mask = SmartSQL.prepDateForSQL(CDate(pDataValue))
        Case adBoolean
            Mask = SQLBoolean(CBool(pDataValue))
        Case Else   'It is Numeric
            Mask = pDataValue
    End Select

End Function

Public Function Optimize(ByVal oRS As Recordset, ByVal sField As String) As Boolean

    On Error GoTo LocalError

    'Create an Index for the specified field
    ' automaticaly uses the index for any Find, Sort, and Filter
    ' operations on the Recordset:
    If EmptyRS(oRS) Then
        'Do Nothing
    ElseIf oRS.CursorLocation = adUseClient Then
        'Works ONLY on Client Side record sets
        oRS.Fields(sField).Properties("OPTIMIZE").Value = True
    End If
    Optimize = True
Exit Function

LocalError:
    LastError = Err.Source & " (" & Err.Number & ") " & Err.Description
End Function

Public Function PutRS(ByRef oRS As ADODB.Recordset) As Boolean

    On Error GoTo LocalError

    Dim Disconnected As Boolean

    Disconnected = oRS.ActiveConnection Is Nothing

    If EmptyRS(oRS) Then
        Exit Function
    ElseIf oRS.LockType = adLockReadOnly Then
        Exit Function
    ElseIf Disconnected Then
        Dim lField As ADODB.Field
        Dim lDirty As Boolean

        cnADX.BeginTrans
        With oRS
            .MoveFirst
            SmartSQL.Reset
            SmartSQL.AddTable oRS.DataMember
            While Not .EOF
                lDirty = False
                For Each lField In .Fields
                    If lField.Value <> lField.OriginalValue Then
                        lDirty = True
                        SmartSQL.AddField lField.Name
                        SmartSQL.AddValue lField.Value
                    End If
                Next lField
                If lDirty Then
                    SmartSQL.AddSimpleWhereClause oRS.Fields(0).Name, oRS.Fields(0).Value, , CLAUSE_EQUALS
                    If Not Execute(MySQL) Then
                        PutRS = False
                        Exit Function
                    End If
                End If
                .MoveNext
            Wend
            PutRS = True
        End With
    Else
        oRS.UpdateBatch adAffectAllChapters
        PutRS = True
    End If

    PutRS = True

Exit Function

LocalError:
    LastError = Err.Source & " (" & Err.Number & ") " & Err.Description
    If cnADO.State = adStateOpen Then
        cnADO.RollbackTrans
    End If
    PutRS = False
End Function

Function RecordsetToHTML(ByVal oRS As ADODB.Recordset, TableAttribs As Variant, Optional NullValues As Variant = "", Optional ShowFieldNames As Boolean = True, Optional IncludeWhiteSpace As Boolean = True) As String

    TableAttribs = CStr(TableAttribs)
    NullValues = CStr(NullValues)
    If NullValues = "" Then NullValues = "&nbsp;"

    Dim res As String
    Dim fld As ADODB.Field
    Dim tmp As String

    ' fill these variables only if spaces are to be kept
    ' prepare the <TABLE> tag
    res = "<TABLE " & TableAttribs & ">" & vbCrLf
    ' show field names, if required
    If ShowFieldNames Then
        res = res & vbTab & "<HEAD>" & vbCrLf
        For Each fld In oRS.Fields
            res = res & vbTab & vbTab & "<TD><B>" & fld.Name & "</B></TD>" & vbCrLf
        Next
        res = res & vbTab & "</HEAD>" & vbCrLf
    End If
    ' get all the records in a semi-formatted string
    tmp = oRS.GetString(, , "</TD>" & vbCrLf & vbTab & vbTab & "<TD>", "</TD>" & vbCrLf & vbTab & "</TR>" & vbCrLf & vbTab & "<TR>" & vbCrLf & vbTab & vbTab & "<TD>", NullValues)
    ' strip what has been appended to the last cell of the last row
    tmp = Left(tmp, Len(tmp) - Len(vbCrLf & vbTab & "<TR>" & vbCrLf & vbTab & vbTab & "<TD>"))
    ' add opening tags to the first cell of the first row of the table and complete the table
    RecordsetToHTML = res & vbTab & "<TR>" & vbCrLf & vbTab & vbTab & "<TD>" & tmp & vbCrLf & "</TABLE>"

End Function

Public Function RefreshRS(ByRef oRS As ADODB.Recordset) As Boolean

    If oRS Is Nothing Then Exit Function

    On Error GoTo LocalError

    Dim Disconnected As Boolean

    With oRS
        If oRS.ActiveConnection Is Nothing Then
            Disconnected = True
            If Connect(False) Then
                Set .ActiveConnection = cnADO
            Else
                Exit Function
            End If
        End If
        'Requery the Recordset
        .Requery
        .MoveFirst
        If Disconnected Then Set .ActiveConnection = Nothing
    End With

    RefreshRS = Not EmptyRS(oRS)
Exit Function

LocalError:
    LastError = Err.Source & " (" & Err.Number & ") " & Err.Description
End Function

Public Function RS2XML(ByVal oRS As ADODB.Recordset, FullPath As String) As Boolean

    On Error GoTo LocalError

    FullPath = CStr(FullPath)
    On Error GoTo 0
    oRS.Save FullPath, adPersistXML
    RS2XML = True
Exit Function

LocalError:
    LastError = Err.Source & " (" & Err.Number & ") " & Err.Description
End Function

Public Function XML2RS(FullPath As Variant) As ADODB.Recordset

    On Error GoTo LocalError

    Set XML2RS = New ADODB.Recordset
    XML2RS.Open FullPath, "Provider=MSPersist;", adOpenForwardOnly, adLockReadOnly, adCmdFile
Exit Function

LocalError:
    LastError = Err.Source & " (" & Err.Number & ") " & Err.Description
End Function

Public Function Schema(SchemaType As ADODB.SchemaEnum, ParamArray Criteria() As Variant) As ADODB.Recordset
    If Connect(False) Then Set Schema = cnADO.OpenSchema(adSchemaColumns, Criteria)
End Function

Public Function SQLBoolean(TrueFalse As Boolean) As Integer

    'This is because SQL True = 1 & VB True = -1
    SQLBoolean = TrueFalse
    If IsSQL Then If TrueFalse = True Then SQLBoolean = TrueFalse * TrueFalse

End Function

Public Function sqlEncode(sqlValue, Optional Encapsulate As Boolean = False) As String

    On Error Resume Next

    sqlEncode = CStr(Replace(sqlValue, "'", "''"))
    If Encapsulate Then sqlEncode = "'" & sqlEncode & "'"

End Function

'##################
'# Class Specific #
'##################
Private Sub Class_Initialize()

    Set SmartSQL = New SmartSQL

End Sub

Private Sub Class_Terminate()

    On Error Resume Next

    CloseALL

    Set SmartSQL = Nothing
    Set cnADO = Nothing
    Set cnADX = Nothing
    Set cnSVR = Nothing

End Sub

Private Sub cnADO_BeginTransComplete(ByVal TransactionLevel As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    'Transaction has begun
End Sub

Private Sub cnADO_CommitTransComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    'Transaction commited
End Sub

Private Sub cnADO_ConnectComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    'Connect
End Sub

Private Sub cnADO_Disconnect(adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    'Disconnect
End Sub

Private Sub cnADO_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    'Transaction executed
End Sub

Private Sub cnADO_InfoMessage(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    'Status of Event
End Sub

Private Sub cnADO_RollbackTransComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    'Transaction rolled back
End Sub

Private Sub cnADO_WillConnect(ConnectionString As String, UserID As String, Password As String, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    'Connecting
End Sub

Private Sub cnADO_WillExecute(Source As String, CursorType As ADODB.CursorTypeEnum, LockType As ADODB.LockTypeEnum, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    'About to Execute
End Sub

Private Sub cnADX_BeginTransComplete(ByVal TransactionLevel As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    'Transaction has begun
End Sub

Private Sub cnADX_CommitTransComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    'Transaction commited
End Sub

Private Sub cnADX_ConnectComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    'Connect
End Sub

Private Sub cnADX_Disconnect(adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    'Disconnect
End Sub

Private Sub cnADX_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    'Transaction executed
End Sub

Private Sub cnADX_InfoMessage(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    'Status of Event
End Sub

Private Sub cnADX_RollbackTransComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    'Transaction rolled back
End Sub

Private Sub cnADX_WillConnect(ConnectionString As String, UserID As String, Password As String, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    'Connecting
End Sub

Private Sub cnADX_WillExecute(Source As String, CursorType As ADODB.CursorTypeEnum, LockType As ADODB.LockTypeEnum, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    'About to Execute
End Sub


