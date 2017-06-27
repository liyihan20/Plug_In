VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FmSel 
   Caption         =   "CRM客户关系管理系统_备料单导入"
   ClientHeight    =   5505
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   10830
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton insertBt 
      Caption         =   "批量插入"
      Height          =   495
      Left            =   4320
      TabIndex        =   9
      Top             =   4920
      Width           =   2175
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3615
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   6376
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      Caption         =   "记录检索"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      Begin VB.CommandButton ckBt 
         Caption         =   "查询"
         Height          =   375
         Left            =   9240
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox orderTxt 
         Height          =   390
         Left            =   7800
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox modelTxt 
         Height          =   375
         Left            =   4800
         TabIndex        =   4
         Top             =   248
         Width           =   2055
      End
      Begin VB.TextBox BillNoTxt 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   248
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "订料员:"
         Height          =   255
         Left            =   6960
         TabIndex        =   5
         Top             =   315
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "规格型号:"
         Height          =   255
         Left            =   3840
         TabIndex        =   3
         Top             =   308
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "流水号/备料单号:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   308
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
If (BillNoTxt.Text = "" Or modelTxt.Text = "" Or orderTxt.Text = "") Then
    MsgBox "检索条件不能为空!需至少填写一个", vbOKOnly, "信息中心提示提示"
    Exit Sub
End If

Call ListLoad
End Sub

Private Sub Form_Activate()
BillNoTxt.SetFocus
End Sub

Private Sub MshInit()

'设置列表的标题和宽度
Dim Rs As New ADODB.Recordset
Dim sql As String
Dim i As Integer
sql = "select field_cn_name,field_width from " + myServer + "dbo.Sale_K3_table_description where bill_type='BL_PO' and account='" + account + "'"
Rs.Open sql, DBServer, adOpenKeyset
fieldsNum = Rs.RecordCount

MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = fieldsNum + 1
MSHFlexGrid1.TextMatrix(0, 0) = ""
MSHFlexGrid1.ColWidth(0) = 400
MSHFlexGrid1.TextMatrix(0, 0) = "标识"

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

sql = "select field_en_name,head_or_entry from " + myServer + "dbo.Sale_K3_table_description where bill_type='BL_PO' and account='" + account + "'"
Rs.Open sql, DBServer, adOpenKeyset

ReDim headIndex(1 To Rs.RecordCount)
ReDim entryIndex(1 To Rs.RecordCount)

'Call MshInit
'设置表头
For i = 1 To Rs.RecordCount
    headIndex(i) = -1
    entryIndex(i) = -1
    If Rs(1) = "h" Then '表头从0开始
        For j = 0 To UBound(m_bill.HeadCtl)
            If UCase(m_bill.HeadCtl(j).fieldname) = UCase(Rs(0)) Then
                headIndex(i) = j
                Exit For
            End If
        Next j
    ElseIf Rs(1) = "e" Then '表体从1开始
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
    StrWhere = "where ([源单单号] like '%" + BillNoTxt.Text + "%' or [订单批号] like '%"
    StrWhere = StrWhere + BillNoTxt.Text + "%') and [物料型号] like '%"
    StrWhere = StrWhere + modelTxt.Text + "%' and [订料员] like '%"
    StrWhere = StrWhere + orderTxt.Text + "%' and [事业部] = '" + account + "'"
End If
    
StrSql = "select * from" + myServer + "dbo.vw_BLPO_plugin " + StrWhere
    

Rs.Open StrSql, DBServer, adOpenKeyset
If Rs.RecordCount > 0 Then
    MSHFlexGrid1.Rows = Rs.RecordCount + 1
Else
'    MSHFlexGrid1.Rows = 2
    MsgBox "查无记录，请检查查询条件"
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
Dim entryId As String
Dim existRs As New ADODB.Recordset
Dim existSql As String

'判断是否已经在K3存在订单编号
For i = 1 To MSHFlexGrid1.Rows - 1
    If Trim(MSHFlexGrid1.TextMatrix(i, 0)) = "Y" Then
        billNo = MSHFlexGrid1.TextMatrix(i, 8)
        entryId = MSHFlexGrid1.TextMatrix(i, 9)
        existSql = "select 1 from POOrder where FSourceBillNo = '" + billNo + "' and FSourceEntryID = " + entryId
        existRs.Open existSql, DBServer, adOpenKeyset
        If (existRs.RecordCount > 0) Then
            If MsgBox("流水号[" + billNo + "]分录号[" + entryId + "]在K3已经存在，是否继续？", vbYesNo, "信息中心提示") = vbNo Then
                existRs.Close
                Exit Sub
            End If
        End If
        existRs.Close
    End If
End If

'表头
For S = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(S, 0) = "Y" Then
        For colCount = 1 To MSHFlexGrid1.Cols - 1
            If headIndex(colCount) >= 0 Then
                m_bill.SetHead headIndex(colCount), MSHFlexGrid1.TextMatrix(S, colCount)
                'MsgBox MSHFlexGrid1.TextMatrix(S, colCount)
            End If
        Next
        Exit For
    End If
Next S
'表体
For S = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(S, 0) = "Y" Then
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
    End If
Next S

End Sub

Private Sub MSHFlexGrid1_Click()
    If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0) = "Y" Then
        MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0) = ""
    Else
        MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0) = "Y"
    End If
End Sub

