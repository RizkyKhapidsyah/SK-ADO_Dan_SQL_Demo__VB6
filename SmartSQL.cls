VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SmartSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

#If False Then
    Private Const sqlDateAndTime = 0
    Private Const sqlDateOnly = 1
    Private Const sqlTimeOnly = 2
#End If

Public Enum SQLDateTypes
    sqlDateAndTime
    sqlDateOnly
    sqlTimeOnly
End Enum

Public Enum JOIN_TYPE
    INNER_JOIN = 1
    LEFT_JOIN = 2
    RIGHT_JOIN = 3
    FULL_JOIN = 4 'SUPPORTED BY SQL SERVER ONLY
End Enum

Public Enum CLAUSE_OPERATOR
    CLAUSE_EQUALS
    CLAUSE_LIKE
    CLAUSE_GREATERTHAN
    CLAUSE_LESSTHAN
    CLAUSE_GREATERTHANOREQUAL
    CLAUSE_LESSTHANOREQUAL
    CLAUSE_DOESNOTEQUAL
    CLAUSE_STARTWITH
    CLAUSE_ENDWITH
End Enum

Public Enum WHERE_CLAUSE_LOGIC
    LOGIC_AND
    LOGIC_OR
End Enum

Public Enum STATEMENT_TYPE
    TYPE_SELECT
    TYPE_INSERT
    TYPE_UPDATE
    TYPE_DELETE
    TYPE_OTHER 'NOT CURRENTLY USED
End Enum

Public Enum SQL_TYPE
    SQL_TYPE_ACCESS
    SQL_TYPE_ANSI
End Enum

Private Enum ERR_NUMBERS
    ERR_TABLE_REQUIRED = 25000
    ERR_LIST_REQUIRED = 25010
    ERR_INVALID_VALUE = 25020
    ERR_INVALID_LISTITEM = 25030
End Enum

Const ERR_TABLE_REQUIRED_DESC = "Table Name property must be set"
Const ERR_LIST_REQUIRED_DESC = "Invalid argument; array or collection required."
Const ERR_INVALID_VALUE_DESC = "Invalid argument type"
Const ERR_INVALID_LISTITEM_DESC = "At least one element in the argument list is invalid"

Private pColOrderClause As Collection
Private pColFieldNames As Collection
Private pColValues As Collection 'for INSERT, UPDATE QUERIES

Private pColWhereClauses As Collection

Private psFromClause As String

Private psWhereClause As String

Private psTableNames() As String
Private psJoinTables As String
Private psJoinFields As String
Private piJoinOp As CLAUSE_OPERATOR
Private piJoinType As JOIN_TYPE

Private piWhereLogic() As WHERE_CLAUSE_LOGIC
Private pbOrderByDesc() As Boolean

Private piStatementType As STATEMENT_TYPE
Private piSQLType As SQL_TYPE
Private psSQL As String
Private psOrderClause As String

Private pbAutoQuote As Boolean
Private pbAutoLike As Boolean
Private pbAutoBracket As Boolean

Const Delimiter = "@*"
'
'#########################################################################################
'                                   Public Subroutines
'#########################################################################################
Public Sub AddComplexWhereClause(ByVal Clause As String, Optional Logic As WHERE_CLAUSE_LOGIC = LOGIC_AND)
    Dim i As Integer

    i = UBound(piWhereLogic) + 1
    ReDim Preserve piWhereLogic(i) As WHERE_CLAUSE_LOGIC
    piWhereLogic(i) = Logic
    pColWhereClauses.Add Clause
End Sub
 
 Public Sub AddField(ByVal FieldName As String, Optional ByVal TableName As String)
    Dim sTable As String
    Dim sField As String
    
    If Len(TableName) Then
        sTable = DoAutoBracket(TableName) & "."
    End If
    sField = DoAutoBracket(FieldName)
    sField = sTable & sField
    
    pColFieldNames.Add sField
End Sub

Public Sub AddFields(ParamArray args() As Variant)
    Dim sSplit() As String
    Dim i As Integer
    Dim sField As String
    
    For i = 0 To UBound(args)
        If ValidateValues(args(i)) = False Then
            Err.Raise ERR_INVALID_VALUE, , ERR_INVALID_VALUE_DESC
        End If
        sField = DoAutoBracket(args(i))
        pColFieldNames.Add sField
    Next
End Sub

Public Sub AddOrderClause(ByVal FieldName As String, Optional OrderDesc = False, Optional ByVal TableName As String)
    Dim iCount As Integer
    Dim sField As String

    iCount = UBound(pbOrderByDesc) + 1
    ReDim Preserve pbOrderByDesc(iCount)
    pbOrderByDesc(iCount) = OrderDesc
    If Len(TableName) Then sField = DoAutoBracket(TableName) & "."
    sField = sField & DoAutoBracket(FieldName)
    pColOrderClause.Add sField
End Sub

Public Function AddSimpleWhereClause(ByVal FieldName As String, ByVal Value As Variant, Optional ByVal TableName As String, Optional Op As CLAUSE_OPERATOR = CLAUSE_EQUALS, Optional Logic As WHERE_CLAUSE_LOGIC = LOGIC_AND) As String

    Dim i               As Integer
    Dim sField          As String
    Dim sWhereStatement As String
    Dim bString         As Boolean
    Dim sValueClause    As String

    If ValidateValues(Value) = False Then Err.Raise ERR_INVALID_VALUE, , ERR_INVALID_VALUE_DESC

    i = UBound(piWhereLogic) + 1
    ReDim Preserve piWhereLogic(i) As WHERE_CLAUSE_LOGIC
    piWhereLogic(i) = Logic

    bString = (VarType(Value) = vbString)

    If Len(TableName) > 0 Then sWhereStatement = DoAutoBracket(TableName) & "."

    sWhereStatement = sWhereStatement & DoAutoBracket(FieldName)
    If Not bString And Op = CLAUSE_LIKE Then
        Op = CLAUSE_EQUALS
    End If

    sWhereStatement = sWhereStatement & " " & TransformOp(Op)

    sValueClause = CStr(Value)

    If Op = CLAUSE_LIKE Then
        If pbAutoLike Then
            sValueClause = LikeCharacter & sValueClause & LikeCharacter
            If pbAutoQuote Then sValueClause = prepStringForSQL(sValueClause)
        Else
            If pbAutoQuote Then sValueClause = prepStringForSQL(sValueClause)
        End If
    ElseIf Op = CLAUSE_STARTWITH Then
        If pbAutoLike Then
            sValueClause = sValueClause & LikeCharacter
            If pbAutoQuote Then sValueClause = prepStringForSQL(sValueClause)
        Else
            If pbAutoQuote Then sValueClause = prepStringForSQL(sValueClause)
        End If
    ElseIf Op = CLAUSE_ENDWITH Then
        If pbAutoLike Then
            sValueClause = LikeCharacter & sValueClause
            If pbAutoQuote Then sValueClause = prepStringForSQL(sValueClause)
        Else
            If pbAutoQuote Then sValueClause = prepStringForSQL(sValueClause)
        End If
    Else
        If pbAutoQuote And bString Then sValueClause = prepStringForSQL(sValueClause)
    End If

    sValueClause = " " & sValueClause
    sWhereStatement = sWhereStatement & sValueClause
    pColWhereClauses.Add sWhereStatement
    AddSimpleWhereClause = sWhereStatement

End Function

Public Sub AddTable(ByVal TableName As String)
    Dim iCount      As Integer
    Dim sTableName  As String

    sTableName = DoAutoBracket(TableName)

    If Not TablePresent(sTableName) Then
        iCount = UBound(psTableNames) + 1
        ReDim Preserve psTableNames(iCount)
        psTableNames(iCount) = sTableName
    End If

    'clear jointables and complex from
    psFromClause = ""
    psJoinTables = ""
    psJoinFields = ""
    piJoinOp = CLAUSE_EQUALS 'default
    piJoinType = INNER_JOIN 'default
End Sub

Public Sub AddValue(ByVal Value As Variant)
    Dim sValue As String

    If UCase(TypeName(Value)) = "FIELD" Then
        Value = Value.Value
    End If

    If Not ValidateValues(Value) Then Err.Raise ERR_INVALID_VALUE, , ERR_INVALID_VALUE_DESC
    If ValidateValues(Value) Then
        sValue = Value
        If VarType(Value) = vbString And pbAutoQuote Then
            sValue = prepStringForSQL(sValue)
        ElseIf VarType(Value) = vbDate And pbAutoQuote Then
            sValue = prepDateForSQL(sValue)
        End If
        pColValues.Add sValue
    End If
End Sub

Public Sub AddValues(ParamArray args() As Variant)
    Dim sSplit() As String
    Dim i    As Integer
    Dim iCtr As Integer
    Dim sAns As String

    For i = 0 To UBound(args)
        If ValidateValues(args(i)) = False Then
            If UCase(TypeName(args(i))) = "FIELD" Then
                args(i) = args(i).Value
            End If
            Err.Raise ERR_INVALID_VALUE, , ERR_INVALID_VALUE_DESC
        End If
    Next

    For i = 0 To UBound(args)
        sAns = args(i)
        If VarType(args(i)) = vbString And pbAutoQuote Then sAns = prepStringForSQL(sAns)
        pColValues.Add sAns
    Next
End Sub

Public Sub ClearFromClause()
    psFromClause = ""
    psJoinTables = ""
    psJoinFields = ""
    piJoinType = INNER_JOIN
    piJoinOp = CLAUSE_EQUALS
    ReDim psTableNames(0) As String
End Sub

Public Sub ClearWhereClause()
    Set pColWhereClauses = New Collection
    ReDim piWhereLogic(0) As WHERE_CLAUSE_LOGIC
End Sub

Public Sub ClearOrderClause()
    Set pColOrderClause = New Collection
    ReDim pbOrderByDesc(0) As Boolean
End Sub

Public Sub ClearFields()
    Set pColFieldNames = New Collection
End Sub

Public Sub ClearValues()
    Set pColValues = New Collection
End Sub

Public Sub ListAddFields(ByVal FieldList As Variant, Optional ByVal TableName As String)
    Dim bValid      As Boolean
    Dim bCollection As Boolean
    Dim sAns        As String
    Dim l           As Long
    Dim v           As Variant
    Dim sItem       As String
    Dim lStartPoint As Long

    If IsObject(FieldList) Then
        bValid = (TypeOf FieldList Is Collection)
        bCollection = True
    Else
        bValid = IsArray(FieldList)
    End If

    If Not bValid Then
        Err.Raise ERR_LIST_REQUIRED, , ERR_LIST_REQUIRED_DESC
        Exit Sub
    End If

    'optional: add type check for each value in array or collection
    'can't have objects,
    
    If bCollection Then
        For Each v In FieldList
            If Not ValidateValues(v) Then Err.Raise ERR_INVALID_LISTITEM, , ERR_INVALID_LISTITEM_DESC
            If Len(v) > 0 Then
                If Len(TableName) Then
                    sAns = DoAutoBracket(TableName) & "."
                End If

                sItem = Trim(CStr(v))
                sAns = sAns & DoAutoBracket(sItem)

                pColFieldNames.Add sAns
                sAns = ""
             End If 'len(v)
        Next v
    Else
        On Error Resume Next
        v = FieldList(0)
        lStartPoint = IIf(Err.Number = 0, 0, 1)
        Err.Clear
        
        On Error GoTo 0
    
        For l = lStartPoint To UBound(FieldList)
            sAns = ""
            If Not ValidateValues(FieldList(l)) Then Err.Raise ERR_INVALID_LISTITEM, , ERR_INVALID_LISTITEM_DESC
                If Len(FieldList(l)) > 0 Then
                    If Len(TableName) Then
                        sAns = sAns & DoAutoBracket(TableName) & "."
                    End If
                    sAns = sAns & DoAutoBracket(FieldList(l))
                pColFieldNames.Add sAns
            End If 'len(FieldList(l) > 0
        Next
    End If 'bcollection
End Sub

Public Sub ListAddValues(ByVal ValueList As Variant)
    Dim bValid      As Boolean
    Dim bCollection As Boolean
    Dim sAns        As String
    Dim l           As Long
    Dim sSplit()    As String
    Dim iCtr        As Integer
    Dim vTest       As String
    Dim lStart      As Long
    Dim v           As Variant
    Dim lStartPoint As Long

    'PURPOSE: ADD A list of values to the values collection
    'Values are for Update or Insert queries
    'The List can be either an array or a collection

    If IsObject(ValueList) Then
        If TypeOf ValueList Is Collection Then
            bValid = True
        Else
            bValid = False
        End If
        bCollection = True
    Else
        bValid = IsArray(ValueList)
    End If

    If Not bValid Then
        Err.Raise ERR_LIST_REQUIRED, , ERR_LIST_REQUIRED_DESC
        Exit Sub
    End If
    
    If bCollection Then
        For Each v In ValueList
            If Not ValidateValues(v) Then Err.Raise ERR_INVALID_LISTITEM, , ERR_INVALID_LISTITEM_DESC
            If VarType(v) <> vbString Or Not pbAutoQuote Then
                   pColValues.Add v
            Else
                pColValues.Add prepStringForSQL(CStr(v))
            End If
        Next
    Else
        'Determine if we are dealing with 0 or 1 bound arrays
        Err.Clear
        On Error Resume Next
        vTest = ValueList(0)
        lStartPoint = IIf(Err.Number = 0, 0, 1)
        Err.Clear
        
        On Error GoTo 0
        
        For l = lStartPoint To UBound(ValueList)
            If Not ValidateValues(ValueList(l)) Then Err.Raise ERR_INVALID_LISTITEM, , ERR_INVALID_LISTITEM_DESC
            sAns = ""
            If VarType(ValueList(l)) <> vbString Or Not pbAutoQuote Then
                pColValues.Add ValueList(l)
            Else
                pColValues.Add prepStringForSQL(CStr(ValueList(l)))
            End If
        Next
    End If
End Sub

Public Sub Reset()
    ClearFromClause
    ClearWhereClause
    ClearOrderClause
    ClearFields
    ClearValues
    'key for field, value for value
    piStatementType = TYPE_SELECT 'default
End Sub

Public Sub SetupJoin(ByVal Table1 As String, ByVal Field1 As String, ByVal Table2 As String, ByVal Field2 As String, Optional Op As CLAUSE_OPERATOR = CLAUSE_EQUALS, Optional JoinType As JOIN_TYPE)
    'check for bracketing and add if not present
    Dim sTable1 As String, sTable2 As String
    Dim sField1 As String, sField2 As String

    sTable1 = DoAutoBracket(Table1)
    sTable2 = DoAutoBracket(Table2)
    sField1 = DoAutoBracket(Field1)
    sField2 = DoAutoBracket(Field2)

    psJoinTables = sTable1 & Delimiter & sTable2
    psJoinFields = sField1 & Delimiter & sField2
    piJoinOp = Op

    Select Case JoinType
        Case Is <= 0, Is > FULL_JOIN
            If piJoinType = 0 Then piJoinType = INNER_JOIN
        Case FULL_JOIN
            piJoinType = IIf(piSQLType = SQL_TYPE_ANSI, FULL_JOIN, IIf(piJoinType = 0, INNER_JOIN, piJoinType))
        Case Else
            piJoinType = JoinType
    End Select

    'reset other from related options
    ReDim psTableNames(0) As String
    psFromClause = ""
End Sub

'#########################################################################################
'                                   Public Properties
'#########################################################################################
Public Property Get AutoBracket() As Boolean
    AutoBracket = pbAutoBracket
End Property
Public Property Let AutoBracket(ByVal NewValue As Boolean)
    pbAutoBracket = NewValue
End Property

Public Property Get AutoLike() As Boolean
    AutoLike = pbAutoLike
End Property
Public Property Let AutoLike(ByVal NewValue As Boolean)
    pbAutoLike = NewValue
End Property

Public Property Get AutoQuote() As Boolean
    AutoQuote = pbAutoQuote
End Property
Public Property Let AutoQuote(ByVal NewValue As Boolean)
    pbAutoQuote = NewValue
End Property

Public Property Let ComplexFromClause(ByVal NewValue As String)
    'reset everything else
    Dim sWkg     As String
    Dim sSplit() As String

    sWkg = Trim$(NewValue)
    
    If Left$(NewValue, 4) = "FROM" Then
        sSplit = Split(sWkg, "FROM")
        psFromClause = Trim$(sSplit(1))
    Else
        psFromClause = sWkg
    End If

    'reset to defaults
    ReDim psTableNames(0) As String

    psJoinTables = ""
    psJoinFields = ""
    piJoinOp = CLAUSE_EQUALS
    piJoinType = INNER_JOIN
End Property
Public Property Get ComplexFromClause() As String
    ComplexFromClause = psFromClause
End Property

Public Property Get JoinType() As JOIN_TYPE
    JoinType = piJoinType
End Property
Public Property Let JoinType(ByVal NewValue As JOIN_TYPE)
If NewValue >= INNER_JOIN And NewValue <= FULL_JOIN Then
    If NewValue = FULL_JOIN And piSQLType = SQL_TYPE_ACCESS And piJoinType = 0 Then
        piJoinType = INNER_JOIN
    Else
        piJoinType = NewValue
    End If
End If
End Property

Public Property Get StatementType() As STATEMENT_TYPE
    StatementType = piStatementType
End Property
Public Property Let StatementType(ByVal NewValue As STATEMENT_TYPE)
    piStatementType = IIf(NewValue <= TYPE_DELETE, NewValue, TYPE_SELECT)
End Property

Public Property Get SQLType() As SQL_TYPE
    SQLType = piSQLType
End Property
Public Property Let SQLType(ByVal NewValue As SQL_TYPE)
    If NewValue = SQL_TYPE_ACCESS Or NewValue = SQL_TYPE_ANSI Then piSQLType = NewValue
End Property

Public Property Get SQL() As String
    MakeStatement
    SQL = psSQL
End Property

Public Property Get TableCount() As Long
    Dim lAns As Long
    If psTableNames(0) = "" And UBound(psTableNames) = 0 Then
        TableCount = ComplexTableCount
    Else
        TableCount = UBound(psTableNames)
    End If
End Property

'#########################################################################################
'                                   Private Properties
'#########################################################################################
Private Property Get LikeCharacter() As String
    LikeCharacter = IIf(piSQLType = SQL_TYPE_ACCESS, "*", "%")
End Property

'#########################################################################################
'                                   Private Subroutines
'#########################################################################################
Private Sub MakeStatement()
    Dim sAns            As String
    Dim sWhereClause    As String
    Dim sOrderClause    As String
    Dim sJoinClause     As String
    Dim sCommand        As String
    Dim i               As Integer
    Dim sOp             As String
    Dim sTemp           As String
    Dim sSplitTables()  As String
    Dim sSplitFields()  As String
    Dim lUpLimit        As Long
    
    Select Case piStatementType
        Case TYPE_SELECT
            If UBound(psTableNames) > 0 Or psJoinTables <> "" Or psFromClause <> "" Then
                'MAKE SELECT CLAUSE
                sCommand = "SELECT "
                'RULES FROM FROM CLAUSES ARE:
                'IF NOT JOINS OR COMPLEX FROM STATEMENTS,
                'YOU CAN HAVE AS MANY TABLES AS YOU WANT:
                'OTHERWISE, USE JUST ONE JOIN OR ONE COMPLEXFROM STATEMENT
                'COLLISIONS:
                'OR: GO WITH THE LATEST ADDED: WHEN ADDING OF ONE TYPE, CLEAR THE OTHER TWO
                If pColFieldNames.Count = 0 Then
                    sCommand = sCommand & "* "
                Else
                    For i = 1 To pColFieldNames.Count
                        sCommand = sCommand & pColFieldNames(i)
                        If i <> pColFieldNames.Count Then sCommand = sCommand & ","
                        sCommand = sCommand & " "
                    Next
                End If
                sCommand = sCommand & "FROM "
                On Error Resume Next
                If Len(psFromClause) > 0 Then
                    sCommand = sCommand & psFromClause
                Else
                    If UBound(psTableNames) >= 1 Then
                        For i = 1 To UBound(psTableNames)
                            sCommand = sCommand & psTableNames(i)
                            If i <> UBound(psTableNames) Then sCommand = sCommand & ", "
                        Next
                    Else
                        sSplitTables = Split(psJoinTables, Delimiter)
                        sSplitFields = Split(psJoinFields, Delimiter)
                        sCommand = sCommand & sSplitTables(0)
                        sCommand = sCommand & " "
                        'FIX TO DEAL WITH JOIN TYPES
                        If piJoinType < 1 Or piJoinType > 4 Then piJoinType = INNER_JOIN
                        Select Case piJoinType
                            Case INNER_JOIN
                                If piSQLType = SQL_TYPE_ACCESS Then sCommand = sCommand & "INNER "
                            Case LEFT_JOIN
                                sCommand = sCommand & " LEFT "
                            Case RIGHT_JOIN
                                sCommand = sCommand & " RIGHT "
                            Case FULL_JOIN
                                sCommand = sCommand & IIf(piSQLType = SQL_TYPE_ACCESS, " INNER ", " FULL ")
                        End Select
                        sCommand = sCommand & "JOIN " & sSplitTables(1) & " ON "
                        If InStr(sSplitFields(0), ".") > 0 Then
                            sCommand = sCommand & sSplitFields(0)
                        Else
                            sCommand = sCommand & sSplitTables(0) & "." & sSplitFields(0)
                        End If
                        sCommand = sCommand & " " & TransformOp(piJoinOp) & " "
                        If InStr(sSplitFields(1), ".") > 0 Then
                            sCommand = sCommand & sSplitFields(1)
                        Else
                            sCommand = sCommand & sSplitTables(1) & "." & sSplitFields(1)
                        End If
                    End If
                End If
            End If 'first condition, testing for at least one table
        Case TYPE_INSERT '?
            If Trim(psTableNames(1)) = "" Then
                Err.Raise ERR_TABLE_REQUIRED, , ERR_TABLE_REQUIRED_DESC
                Exit Sub
            End If
            sCommand = "INSERT INTO " & psTableNames(1)
            If pColFieldNames.Count > 0 Then
                sCommand = sCommand & " ("
                For i = 1 To pColFieldNames.Count
                    sCommand = sCommand & pColFieldNames(i)
                    If i <> pColFieldNames.Count Then sCommand = sCommand & ", "
                    
                Next
                sCommand = sCommand & ")"
            End If
            If pColValues.Count > 0 Then
                sCommand = sCommand & " VALUES ("
                For i = 1 To pColValues.Count
                    sCommand = sCommand & pColValues(i)
                    If i <> pColValues.Count Then sCommand = sCommand & ", "
                Next
                sCommand = sCommand & ")"
            End If
    Case TYPE_UPDATE
          If pColFieldNames.Count > 0 And pColValues.Count > 0 And psTableNames(1) <> "" Then
          lUpLimit = IIf(pColFieldNames.Count > pColValues.Count, pColValues.Count, pColFieldNames.Count)
          sCommand = "UPDATE " & psTableNames(1) & " SET "
          For i = 1 To lUpLimit
            sCommand = sCommand & pColFieldNames(i) & " = " & pColValues(i)
            If i <> lUpLimit Then sCommand = sCommand & ", "
          Next
          End If
    Case TYPE_DELETE
          If psTableNames(1) <> "" Then
            sCommand = "DELETE FROM " & psTableNames(1)
          End If
    End Select
    
    If piStatementType <> TYPE_INSERT And sCommand <> "" Then
        For i = 1 To pColWhereClauses.Count
            If i = 1 Then
                sWhereClause = "WHERE"
            Else
                sWhereClause = sWhereClause & IIf(piWhereLogic(i) = LOGIC_AND, " AND", " OR")
            End If
            sWhereClause = sWhereClause & " (" & pColWhereClauses.Item(i) & ")"
            'If Not pbWhereClauseNumeric(i) Then sWhereClause = sWhereClause & "'"
        Next
    End If ' pistatement type <> ..
    
    'ORDER CLAUSE
    If piStatementType = TYPE_SELECT Then
        For i = 1 To pColOrderClause.Count
            If i = 1 Then sOrderClause = "ORDER BY "
            sOrderClause = sOrderClause & pColOrderClause.Item(i)
            If pbOrderByDesc(i) = True Then sOrderClause = sOrderClause & " DESC"
           If i <> pColOrderClause.Count Then sOrderClause = sOrderClause & ", "
         Next
    End If
    sAns = sCommand
    If Len(sWhereClause) > 0 Then sAns = sAns & " " & sWhereClause
    If Len(sOrderClause) > 0 Then sAns = sAns & " " & sOrderClause
    psOrderClause = sOrderClause
    psSQL = sAns
End Sub

'#########################################################################################
'                                   Private Functions
'#########################################################################################
Private Function DistinctValues(InputArray As Variant) As String()
    Dim asAns()     As String
    Dim lStartPoint As Long
    Dim lEndPoint   As Long
    Dim lCount      As Long
    Dim col         As New Collection
    Dim l           As Long
    Dim vTest       As Variant

    ReDim asAns(0) As String

    lCount = UBound(InputArray)

    On Error Resume Next
    vTest = InputArray(0)
    lStartPoint = IIf(Err.Number = 0, 0, 1)
    Err.Clear

    For l = lStartPoint To lCount
        col.Add 0, InputArray(l)
        If Err.Number = 0 Then
            If asAns(0) = "" Then
                asAns(0) = InputArray(l)
            Else
                ReDim Preserve asAns(UBound(asAns) + 1) As String
                asAns(UBound(asAns)) = InputArray(l)
            End If
        End If
        Err.Clear
    Next
    DistinctValues = asAns
End Function

Private Function DoAutoBracket(ByVal DBObjectName As String) As String
    Dim sSplit() As String
    Dim sAns As String
    Dim iCtr As Integer

    If InStr(DBObjectName, ".") > 0 Then
        sSplit = Split(DBObjectName, ".")
        For iCtr = 0 To UBound(sSplit)
            If InStr(sSplit(iCtr), " ") > 0 And InStr(sSplit(iCtr), "(") = 0 And InStr(sSplit(iCtr), ")") = 0 And InStr(sSplit(iCtr), "[") = 0 And pbAutoBracket Then
                sAns = sAns & "[" & sSplit(iCtr) & "]"
            Else
                sAns = sAns & Trim(sSplit(iCtr))
            End If
        
            If iCtr < UBound(sSplit) Then sAns = sAns & "."
        Next
    Else
        sAns = Trim(DBObjectName)
        If InStr(sAns, " ") > 0 And InStr(sAns, "(") = 0 And Left$(sAns, 1) <> "[" And pbAutoBracket Then
            sAns = "[" & sAns & "]"
        End If
    End If

    DoAutoBracket = sAns
End Function

Public Function prepDateForSQL(ByVal vDate As Variant, Optional vType As SQLDateTypes = sqlDateAndTime) As String

    On Error GoTo LocalError

    'Remove all invalid characters
    vDate = Trim(CStr(vDate))
    vDate = Replace(vDate, "#", "")
    vDate = Replace(vDate, "'", "")
    vDate = Replace(vDate, Chr(34), "")

    '--------------------------------------
    'Convert the Date to a Double Precision
    '   for international compatability
    '--------------------------------------
    prepDateForSQL = ""

    'First see in what format the data came
    If Not IsDate(vDate) Or IsNull(vDate) Then
        'Maybe it is a number
        If IsNumeric(vDate) Then
            vDate = CDate(vDate)
        End If
        If Not IsDate(vDate) Then
            'Still not a date
            Exit Function
        End If
    End If

    'Valid if we get this far
    Dim lDelim As String

    lDelim = IIf(SQL_TYPE_ANSI, "'", "#")
    If vType = sqlDateOnly Then
        prepDateForSQL = IIf(SQL_TYPE_ANSI, Format(vDate, "mm\/dd\/yyyy"), FormatDateTime(vDate, vbShortDate))
    ElseIf vType = sqlTimeOnly Then
        prepDateForSQL = IIf(SQL_TYPE_ANSI, Format(vDate, "hh\:mm\:ss"), Format(vDate, "hh\:mm\:ss"))
    Else    'Return Date and Time
        prepDateForSQL = IIf(SQL_TYPE_ANSI, Format(vDate, "mm\/dd\/yyyy hh\:mm\:ss"), FormatDateTime(vDate, vbShortDate) & " " & Format(vDate, "hh\:mm\:ss"))
    End If

    prepDateForSQL = lDelim & prepDateForSQL & lDelim

Exit Function

LocalError:
    prepDateForSQL = ""
End Function

Private Function prepStringForSQL(ByVal sValue As String) As String
    Dim sAns As String
    sAns = Replace(sValue, Chr(39), "''")
    sAns = "'" & sAns & "'"
    prepStringForSQL = sAns
End Function

Private Function ComplexTableCount() As Long
    Dim sSplit()    As String
    Dim sSplit2()   As String
    
    Dim sInput      As String
    Dim asTables()  As String

    Dim sFinal      As String
    Dim iCtr        As Integer

    Dim vUnique     As Variant
    Dim iPos        As Integer
    Dim lAns        As Long

    If psFromClause <> "" Then
        sSplit = Split(psFromClause, " ")
        ReDim asTables(0) As String
        asTables(0) = sSplit(0)
        
        sSplit = Split(psFromClause, "JOIN")
    
        For iCtr = 1 To UBound(sSplit)
            sSplit2 = Split(Trim$(sSplit(iCtr)), " ")
            ReDim Preserve asTables(UBound(asTables) + 1)
            asTables(UBound(asTables)) = sSplit2(0)
        Next
    
        vUnique = DistinctValues(asTables)
    
        lAns = UBound(vUnique) + 1
    ElseIf Trim$(psJoinTables) <> "" Then
        sSplit = Split(psJoinTables, Delimiter)
        vUnique = DistinctValues(sSplit)
        lAns = UBound(vUnique) + 1
    End If

    ComplexTableCount = lAns
End Function

Private Function TablePresent(TableName As String) As Boolean
    Dim iCtr As Integer
    Dim bAns As Boolean

    If UBound(psTableNames) = 0 Then Exit Function

    For iCtr = 1 To UBound(psTableNames)
        If TableName = psTableNames(iCtr) Then bAns = True
    Next
    TablePresent = bAns
End Function

Private Function TransformOp(Op As CLAUSE_OPERATOR) As String
    Dim sOp As String

    Select Case Op
        Case CLAUSE_EQUALS
            sOp = "="
        Case CLAUSE_LIKE
            sOp = "LIKE"
        Case CLAUSE_STARTWITH
            sOp = "LIKE"
        Case CLAUSE_ENDWITH
            sOp = "LIKE"
        Case CLAUSE_GREATERTHAN
            sOp = ">"
        Case CLAUSE_LESSTHAN
            sOp = "<"
        Case CLAUSE_GREATERTHANOREQUAL
            sOp = ">="
        Case CLAUSE_LESSTHANOREQUAL
            sOp = "<="
         Case CLAUSE_DOESNOTEQUAL
            sOp = "<>"
        Case Else
            sOp = "="
    End Select
    TransformOp = sOp
End Function

Private Function ValidateValues(Values As Variant) As Boolean
    'Purpose: Determines if a collection, variant array, or single value
    'has valid values for an SQL String
    Dim bCollection     As Boolean
    Dim iBadVarTypes(4) As Integer
    Dim v               As Variant
    Dim i               As Integer
    Dim lCtr            As Long
    Dim lListCount      As Long
    Dim lStartPoint     As Long
    Dim iCount          As Integer

    Dim bAns As Boolean

    iBadVarTypes(0) = vbObject
    iBadVarTypes(1) = vbError
    iBadVarTypes(2) = vbDataObject
    iBadVarTypes(3) = vbUserDefinedType
    iBadVarTypes(4) = vbArray

    bAns = True
    iCount = UBound(iBadVarTypes)

    If IsObject(Values) Then
        If Not TypeOf Values Is Collection Then
            ValidateValues = False
            Exit Function
        End If
    Else
        If Not VarType(Values) = vbArray Then
            For i = 0 To iCount
                If VarType(Values) = iBadVarTypes(i) Then
                    bAns = False
                    Exit For
                End If
            Next
            ValidateValues = bAns
            Exit Function
        End If
    End If

    bCollection = IsObject(Values) 'has to be collection

    If bCollection Then
        For Each v In Values
            For i = 1 To iCount
                If VarType(v) = iBadVarTypes(i) Or VarType(v) = iBadVarTypes(i) + vbVariant Then
                    bAns = False
                    Exit For
                End If
            Next
            If bAns = False Then Exit For
        Next
    Else
        lListCount = UBound(Values)
        On Error Resume Next
        v = Values(0)
        lStartPoint = IIf(Err.Number = 0, 0, 1)
        Err.Clear
        On Error GoTo 0
        For lCtr = lStartPoint To lListCount
            For i = 1 To iCount
                If VarType(Values(lCtr)) = iBadVarTypes(i) Or VarType(v) = iBadVarTypes(i) + vbVariant Then
                    bAns = False
                    Exit For
                End If
            Next
            If bAns = False Then Exit For
        Next
    End If

    ValidateValues = bAns
End Function


Private Sub Class_Initialize()
    Reset
    pbAutoLike = True
    pbAutoQuote = True
    piSQLType = SQL_TYPE_ACCESS
    pbAutoBracket = True
End Sub

