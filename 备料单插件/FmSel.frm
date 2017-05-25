VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FmSel 
   Caption         =   "CRM�ͻ���ϵ����ϵͳ"
   ClientHeight    =   3840
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   8775
   StartUpPosition =   2  '��Ļ����
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2775
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4895
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      Caption         =   "CRM-���ϵ�"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin VB.CommandButton insertBt 
         Caption         =   "��������"
         Height          =   375
         Left            =   4920
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton ckBt 
         Caption         =   "��ѯ"
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox billNoTxt 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "��ˮ��/���ϵ��ţ�"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FmSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public m_bill As k3BillTransfer.Bill
Private ColItemID As Integer
Private fieldsNum As Integer
Private headIndex() As Integer
Private entryIndex() As Integer

Private Sub ckBt_Click()
If (billNoTxt.Text = "") Then
    MsgBox "��ˮ��/���ϵ��Ų���Ϊ��!", vbOKOnly, "��Ϣ������ʾ��ʾ"
    Exit Sub
End If

Call ListLoad
End Sub

Private Sub Form_Activate()
billNoTxt.SetFocus
End Sub

Private Sub MshInit()

'�����б�ı���Ϳ��
Dim Rs As New ADODB.Recordset
Dim sql As String
Dim i As Integer
sql = "select field_cn_name,field_width from " + myServer + "dbo.Sale_K3_table_description where bill_type='BL' and account='" + account + "'"
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

Private Sub Form_Load()
Dim Rs As New ADODB.Recordset
Dim i As Integer
Dim j As Integer
Dim sql As String

sql = "select field_en_name,head_or_entry from " + myServer + "dbo.Sale_K3_table_description where bill_type='BL' and account='" + account + "'"
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
    StrWhere = "where (Դ������  = '" + billNoTxt.Text + "' or ԭ���۶��� = '" + billNoTxt.Text + "') and �������� = '" + account + "'"
End If
    
StrSql = "select * from" + myServer + "dbo.vw_BL_plugin " + StrWhere
    

Rs.Open StrSql, DBServer, adOpenKeyset
If Rs.RecordCount > 0 Then
    MSHFlexGrid1.Rows = Rs.RecordCount + 1
Else
'    MSHFlexGrid1.Rows = 2
    MsgBox "���޴˼�¼����������ȷ����ˮ��/���ϵ���"
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

Private Sub insertBt_Click()
Dim vItemID As String
Dim R As Integer
Dim S As Integer
Dim VarTmp As Variant
Dim colCount As Integer
Dim billNo As String
Dim existRs As New ADODB.Recordset
Dim existSql As String

'�ж��Ƿ��Ѿ���K3���ڶ������
billNo = MSHFlexGrid1.TextMatrix(1, 1)
If (billNo = "") Then
    MsgBox "���Ȳ�ѯ��Ҫ����ļ�¼", vbOKOnly, "��Ϣ������ʾ"
    Exit Sub
End If
existSql = "select FInterID from SEOrder where FBillNo='" + billNo + "'"
existRs.Open existSql, DBServer, adOpenKeyset
If (existRs.RecordCount > 0) Then
    If MsgBox("�ö��������K3�Ѿ����ڣ��Ƿ������", vbYesNo, "��Ϣ������ʾ") = vbNo Then
        existRs.Close
        Exit Sub
    End If
End If
existRs.Close
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
