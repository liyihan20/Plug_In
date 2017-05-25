VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FmSel 
   Caption         =   "CRM客户关系管理系统"
   ClientHeight    =   3840
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   8775
   StartUpPosition =   2  '屏幕中心
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
      Caption         =   "CRM-备料单"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin VB.CommandButton insertBt 
         Caption         =   "批量插入"
         Height          =   375
         Left            =   4920
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton ckBt 
         Caption         =   "查询"
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
         Caption         =   "流水号/备料单号："
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
    MsgBox "流水号/备料单号不可为空!", vbOKOnly, "信息中心提示提示"
    Exit Sub
End If

Call ListLoad
End Sub

Private Sub Form_Activate()
billNoTxt.SetFocus
End Sub

Private Sub MshInit()

'设置列表的标题和宽度
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
    StrWhere = "where (源单单号  = '" + billNoTxt.Text + "' or 原销售订单 = '" + billNoTxt.Text + "') and 生产部门 = '" + account + "'"
End If
    
StrSql = "select * from" + myServer + "dbo.vw_BL_plugin " + StrWhere
    

Rs.Open StrSql, DBServer, adOpenKeyset
If Rs.RecordCount > 0 Then
    MSHFlexGrid1.Rows = Rs.RecordCount + 1
Else
'    MSHFlexGrid1.Rows = 2
    MsgBox "查无此记录，请输入正确的流水号/备料单号"
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

'判断是否已经在K3存在订单编号
billNo = MSHFlexGrid1.TextMatrix(1, 1)
If (billNo = "") Then
    MsgBox "请先查询需要导入的记录", vbOKOnly, "信息中心提示"
    Exit Sub
End If
existSql = "select FInterID from SEOrder where FBillNo='" + billNo + "'"
existRs.Open existSql, DBServer, adOpenKeyset
If (existRs.RecordCount > 0) Then
    If MsgBox("该订单编号在K3已经存在，是否继续？", vbYesNo, "信息中心提示") = vbNo Then
        existRs.Close
        Exit Sub
    End If
End If
existRs.Close
Me.Hide

'表头
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
'表体
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
