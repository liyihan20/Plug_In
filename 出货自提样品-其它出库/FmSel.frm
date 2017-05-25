VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FmSel 
   Caption         =   "������Ʒ���"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   10260
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command2 
      Caption         =   "��������"
      Height          =   495
      Left            =   3840
      TabIndex        =   9
      Top             =   4800
      Width           =   2535
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3735
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   6588
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      Caption         =   "��������"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin VB.TextBox Model 
         Height          =   390
         Left            =   7080
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox BathNo 
         Height          =   390
         Left            =   4200
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox BillNO 
         Height          =   390
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "��ѯ"
         Height          =   375
         Left            =   9240
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "����ͺţ�"
         Height          =   255
         Left            =   6240
         TabIndex        =   6
         Top             =   308
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "���ţ�"
         Height          =   255
         Left            =   3600
         TabIndex        =   4
         Top             =   308
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "�������뵥�ţ�"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   308
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FmSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ColItemID As Integer
Private fieldsNum As Integer
Private headIndex() As Integer
Private entryIndex() As Integer

Private Sub Command1_Click()
    If BillNO.Text = "" Then
        MsgBox "�������뵥�Ų���Ϊ��"
        Exit Sub
    End If
    Call ListLoad
    
End Sub

Private Sub Command2_Click()
Dim vItemID As String
Dim i As Integer
Dim j As Integer
Dim R As Integer
Dim BillNO As String
Dim existRs As New ADODB.Recordset
Dim existSql As String

'�ȼ���Ƿ��Ѿ��е����K3
For i = 1 To MSHFlexGrid1.Rows - 1
    If Trim(MSHFlexGrid1.TextMatrix(i, 0)) = "Y" Then
        BillNO = MSHFlexGrid1.TextMatrix(i, 8)
        existSql = "select 1 from ICStockBillEntry where " + billNoSeg + " = '" + BillNO + "'"
        existRs.Open existSql, DBServer, adOpenKeyset
        If existRs.RecordCount > 0 Then
            If MsgBox("���뵥�š�" + BillNO + "����K3������ⵥ�Ѵ��ڣ��Ƿ�������룿", vbYesNo, "��Ϣ������ʾ") = vbNo Then
                existRs.Close
                Exit Sub
            End If
        End If
        existRs.Close
    End If
Next i

'��ͷ����һ��
For i = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(i, 0) = "Y" Then
        For j = 1 To MSHFlexGrid1.Cols - 1
            If headIndex(j) >= 0 Then
                m_bill.SetHead headIndex(j), MSHFlexGrid1.TextMatrix(i, j)
                'MsgBox MSHFlexGrid1.TextMatrix(i, j)
            End If
        Next j
        Exit For  '����һ��֮���˳�
    End If
Next i


'��������Σ��洢���һ�ε��кţ������´μ�������
For i = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(i, 0) = "Y" Then
        vItemID = m_bill.GetGridText(1, ColItemID)
        R = 1
        Do While vItemID <> ""
           R = R + 1
           vItemID = m_bill.GetGridText(R, ColItemID)
        Loop
        If R > 1 Then
            m_bill.BillForm.InsertRow
        End If
        
        For j = 1 To MSHFlexGrid1.Cols - 1
            If entryIndex(j) >= 0 Then
                m_bill.SetGridText R, entryIndex(j), MSHFlexGrid1.TextMatrix(i, j)
            End If
        Next j
    End If
Next i


End Sub

Private Sub Form_Activate()
    BillNO.SetFocus
End Sub

Private Sub Form_Load()
    Dim Rs As New ADODB.Recordset
    Dim i As Integer
    Dim j As Integer
    Dim sql As String
    
    BillNO.Text = Right("" & Year(DateTime.Date), 2) & Right("0" & Month(DateTime.Date), 2)
    
    sql = "select en_name,head_or_entry from dbo.liyh_table_description where bill_type='OtherIN'"
    Rs.Open sql, DBServer, adOpenKeyset
    
    
    ReDim headIndex(1 To Rs.RecordCount)
    ReDim entryIndex(1 To Rs.RecordCount)
    
    '���ñ�ͷ
    For i = 1 To Rs.RecordCount
        headIndex(i) = -1
        entryIndex(i) = -1
        If Rs(1) = "h" Then '��ͷ��0��ʼ
            For j = 0 To UBound(m_bill.HeadCtl)
                If UCase(m_bill.HeadCtl(j).fieldname) = UCase(Rs(0)) Then
                    headIndex(i) = j
                    Exit For
                End If
            Next j
        ElseIf Rs(1) = "e" Then '�����1��ʼ
            For j = 1 To UBound(m_bill.EntryCtl)
                If UCase(m_bill.EntryCtl(j).fieldname) = UCase(Rs(0)) Then
                    entryIndex(i) = j
                    If UCase(m_bill.EntryCtl(j).fieldname) = "FITEMID" Then
                        ColItemID = j
                    End If
                    Exit For
                End If
            Next j
        End If
        Rs.MoveNext
    Next i
    Rs.Close
End Sub


Private Sub MshInit()

'�����б�ı���Ϳ��
Dim Rs As New ADODB.Recordset
Dim sql As String
Dim i As Integer
sql = "select cn_name,width from dbo.liyh_table_description where bill_type='OtherIN'"
Rs.Open sql, DBServer, adOpenKeyset

MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = Rs.RecordCount + 1
MSHFlexGrid1.TextMatrix(0, 0) = ""
MSHFlexGrid1.ColWidth(0) = 400
MSHFlexGrid1.TextMatrix(0, 0) = "��ʶ"

For i = 1 To Rs.RecordCount
    MSHFlexGrid1.TextMatrix(0, i) = Rs(0)
    MSHFlexGrid1.ColWidth(i) = Rs(1)
    Rs.MoveNext
Next i
Rs.Close

End Sub

Private Sub ListLoad()
Dim Rs As New ADODB.Recordset
Dim StrWhere As String
Dim StrSql As String
Dim R As Long
Dim T As Integer
Dim fieldIndex As Integer
fieldIndex = 0
On Error Resume Next
If TxtBillNo.Text <> "" Then
    StrWhere = "where �������뵥��  like '%" + BillNO.Text + "%' and ���� like '%" + BathNo.Text + "%' and ����ͺ� like '%" + Model.Text + "%'"
End If
    
StrSql = "select * from dbo.vw_lyh_chyp_plugin  " + StrWhere + " order by ����ͺ�"

Rs.Open StrSql, DBServer, adOpenKeyset
If Rs.RecordCount > 0 Then
    MSHFlexGrid1.Rows = Rs.RecordCount + 1
Else
'    MSHFlexGrid1.Rows = 2
    MsgBox "���޷���������δ����ļ�¼"
    Rs.Close
    Exit Sub
End If

Call MshInit
For R = 1 To Rs.RecordCount
    MSHFlexGrid1.TextMatrix(R, 0) = ""
    For T = 0 To Rs.Fields.Count - 1
        MSHFlexGrid1.TextMatrix(R, T + 1) = Rs(T)
    Next T
    fieldIndex = fieldIndex + 1
    Rs.MoveNext
Next R
Rs.Close
End Sub

Private Sub MSHFlexGrid1_Click()
    If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0) = "Y" Then
        MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0) = ""
    Else
        MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0) = "Y"
    End If
End Sub
