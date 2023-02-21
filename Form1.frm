VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   60
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   2445
      Left            =   60
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   540
      Width           =   4485
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This is much improved version of the ADO Utilities Class

Private Const cnString  As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Program Files\Microsoft Visual Studio\VB98\NWIND.mdb;Persist Security Info=False"
Private oADO            As adoUtils
Private rsCust          As ADODB.Recordset

Private Sub Form_Load()

    Set oADO = New adoUtils
    With oADO
        .ConnectString = cnString
        If .Connect(ConnectClientSide) Then
            'Modified version of FreeVBCode SmartSQL
            With .SmartSQL
                .SQLType = SQL_TYPE_ACCESS
                .StatementType = TYPE_SELECT
                .AddTable "Customers"
                .AddField "CustomerID"
                .AddField "CompanyName"
                .AddField "ContactName"
            End With
            Set rsCust = .GetRS(.MySQL, adLockOptimistic, adOpenKeyset, True, ConnectClientSide)
        End If
    End With

    If Not oADO.EmptyRS(rsCust) Then
        With Combo1
            .Clear
            While Not rsCust.EOF
                .AddItem rsCust![ContactName]
                rsCust.MoveNext
            Wend
        End With
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set rsCust = Nothing
    Set oADO = Nothing

End Sub
