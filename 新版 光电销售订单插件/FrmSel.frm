VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmSel 
   Caption         =   "FrmSel"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   11475
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame 
      Caption         =   "��������"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      Begin VB.TextBox TxtBillNo 
         Height          =   350
         Left            =   1800
         TabIndex        =   3
         Top             =   235
         Width           =   2415
      End
      Begin VB.CommandButton ComEnter 
         Caption         =   "��ѯ"
         Default         =   -1  'True
         Height          =   350
         Left            =   4320
         TabIndex        =   2
         Top             =   235
         Width           =   1095
      End
      Begin VB.CommandButton ComAdd 
         Caption         =   "��������"
         Height          =   350
         Left            =   5760
         TabIndex        =   1
         Top             =   235
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "��������ϵͳ"
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "��ˮ��/������ţ�"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   320
         Width           =   1575
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3855
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   6800
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "FrmSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public m_bill As k3BillTransfer.Bill
Public myServer As String
Private ColItemID As Integer
Private fieldsNum As Integer
Private headIndex() As Integer
Private entryIndex() As Integer
Private CanNotAdd As Boolean


Private Sub ComAdd_Click()
Dim vItemID As String
Dim R As Integer
Dim S As Integer
Dim VarTmp As Variant
Dim colCount As Integer
Dim sysNo As String
Dim billNo As String
Dim existRs As New ADODB.Recordset
Dim existSql As String

sysNo = MSHFlexGrid1.TextMatrix(1, 1)
If (sysNo = "") Then
    MsgBox "���Ȳ�ѯ��Ҫ����ļ�¼", vbOKOnly, "��Ϣ������ʾ"
    Exit Sub
End If

'�ж��Ƿ��Ѿ���k3�д�����ˮ��
existSql = "select FInterID from SEOrderEntry where FSourceBillNo='" + sysNo + "'"
existRs.Open existSql, DBServer, adOpenKeyset
If (existRs.RecordCount > 0) Then
    If MsgBox("����ˮ���֮ǰ�Ѿ�����ȥ��K3���Ƿ������", vbYesNo, "��Ϣ������ʾ") = vbNo Then
        existRs.Close
        Exit Sub
    End If
End If
existRs.Close

'�ж��Ƿ��Ѿ���K3���ڶ������
If account <> "ele" Then  '����û�ж������
    billNo = MSHFlexGrid1.TextMatrix(1, 2)
    existSql = "select FInterID from SEOrder where FBillNo='" + billNo + "'"
    existRs.Open existSql, DBServer, adOpenKeyset
    If (existRs.RecordCount > 0) Then
        If MsgBox("�ö��������K3�Ѿ����ڣ��Ƿ������", vbYesNo, "��Ϣ������ʾ") = vbNo Then
            existRs.Close
            Exit Sub
        End If
    End If
    existRs.Close
End If

Me.Hide

'��ͷ
For S = 1 To MSHFlexGrid1.Rows - 1
'    If MSHFlexGrid1.TextMatrix(S, 0) = "Y" Then
        For colCount = 1 To MSHFlexGrid1.Cols - 1
            If headIndex(colCount) >= 0 Then
                m_bill.SetHead headIndex(colCount), MSHFlexGrid1.TextMatrix(S, colCount)
                'MsgBox MSHFlexGrid1.TextMatrix(S, colCount)
            End If
        Next
        Exit For
'    End If
Next S
'����
For S = 1 To MSHFlexGrid1.Rows - 1
'    If MSHFlexGrid1.TextMatrix(S, 0) = "Y" Then
        vItemID = m_bill.GetGridText(1, ColItemID)
        R = 1
        Do While vItemID <> ""
           R = R + 1
           vItemID = m_bill.GetGridText(R, ColItemID)
        Loop
        If R > 1 Then
            m_bill.BillForm.InsertRow
        End If
        
        For colCount = 1 To MSHFlexGrid1.Cols - 1
            If entryIndex(colCount) >= 0 Then
                m_bill.SetGridText R, entryIndex(colCount), MSHFlexGrid1.TextMatrix(S, colCount)
            End If
        Next
'    End If
Next S
Unload Me
End Sub

Private Sub ComEnter_Click()
If (TxtBillNo.Text = "") Then
    MsgBox "��ˮ�Ų���Ϊ��!", vbOKOnly, "�����ʾ"
    Exit Sub
End If

Call ListLoad
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
If account <> "ele" Then
    If TxtBillNo.Text <> "" Then
        StrWhere = "where ��ˮ��  = '" + TxtBillNo.Text + "' or ������� = '" + TxtBillNo.Text + "'"
    End If
    StrSql = "select * from" + myServer + "dbo.vw_SaleOrder_plugin  " + StrWhere + " order by Դ���к� "
Else
    '���ӵ��ֶβ�һ��
    If TxtBillNo.Text <> "" Then
        StrWhere = "where Դ������  = '" + TxtBillNo.Text + "'"
    End If
    StrSql = "select * from" + myServer + "dbo.vw_SaleOrder_plugin  " + StrWhere + " order by Դ����¼ "

End If

Rs.Open StrSql, DBServer, adOpenKeyset
If Rs.RecordCount > 0 Then
    MSHFlexGrid1.Rows = Rs.RecordCount + 1
Else
'    MSHFlexGrid1.Rows = 2
    MsgBox "���޴˼�¼����������ȷ����ˮ���"
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

Private Sub Form_Activate()
    TxtBillNo.SetFocus
End Sub

'Private Sub MSHFlexGrid1_Click()
'If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0) = "Y" Then
'    MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0) = ""
'Else
'    MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0) = "Y"
'End If
'End Sub
Private Sub Form_Load()
Dim Rs As New ADODB.Recordset
Dim i As Integer
Dim j As Integer
Dim sql As String

sql = "select field_en_name,head_or_entry from " + myServer + "dbo.Sale_K3_table_description where bill_type='SO'"

If account = "ele" Then '������Ҫ����sort_no����
    sql = sql + " order by sort_no"
End If

Rs.Open sql, DBServer, adOpenKeyset


ReDim headIndex(1 To Rs.RecordCount)
ReDim entryIndex(1 To Rs.RecordCount)

'Call MshInit
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
CanNotAdd = False
End Sub


Private Sub MshInit()

'�����б�ı���Ϳ��
Dim Rs As New ADODB.Recordset
Dim sql As String
Dim i As Integer
sql = "select field_cn_name,field_width from " + myServer + "dbo.Sale_K3_table_description where bill_type='SO'"
If account = "ele" Then '������Ҫ����sort_no����
    sql = sql + " order by sort_no"
End If

Rs.Open sql, DBServer, adOpenKeyset
fieldsNum = Rs.RecordCount

MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = fieldsNum + 1
MSHFlexGrid1.TextMatrix(0, 0) = ""
MSHFlexGrid1.ColWidth(0) = 400

For i = 1 To Rs.RecordCount
    MSHFlexGrid1.TextMatrix(0, i) = Rs(0)
    MSHFlexGrid1.ColWidth(i) = Rs(1)
    Rs.MoveNext
Next i
Rs.Close

End Sub

