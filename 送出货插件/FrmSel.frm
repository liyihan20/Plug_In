VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmSel 
   Caption         =   "查询送货申请单"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   11565
   StartUpPosition =   3  '窗口缺省
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6615
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   11668
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame 
      Caption         =   "申请单查询"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      Begin VB.CommandButton ComAdd 
         Caption         =   "批量插入"
         Height          =   350
         Left            =   5520
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton ComEnter 
         Caption         =   "查询"
         Height          =   350
         Left            =   4320
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TxtBillNo 
         Height          =   350
         Left            =   1800
         TabIndex        =   1
         Top             =   235
         Width           =   2415
      End
      Begin VB.Label lblSelect 
         Alignment       =   1  'Right Justify
         Caption         =   "送货申请单号"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   315
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FrmSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public m_bill As k3BillTransfer.Bill
Public SelType As String
Public sFTranType As String
Public sFBillNo As String
Public IsSelectAll As Boolean
Public StoreProcName As String
Private ColItemID As Integer
Private ColUnitID As Integer
Private ColQty As Integer
Private ColSBillNo As Integer
Private ColSInderID As Integer
Private ColSTranType As Integer
Private ColSEntryID As Integer
Private ColAMount As Integer
Private ColPrice As Integer
Private ColSNo As Integer
Private ColSName As Integer
Private ColSNote As Integer
Private ColNote As Integer
Private ColMode As Integer
Private ColNo As Integer
Private ColCNo As Integer
Private ColCID As Integer
Private ColCEntry As Integer
Private ColTrda As Integer
Private ColTrade As Integer
Private ColEmp As Integer
Private ColSy As Integer
Private ColEveryQty As Integer
Private ColDOcQty As Integer
Private ColBoxnum As Integer
Private ColDocNum As Integer
Private ColDOcHight As Integer
Private ColDocClear As Integer
Private ColTotalHight As Integer
Private ColTotalClear As Integer
Private Colmustqty As Integer
Private ColBatno As Integer
Private ColOrder As Integer
Private ColOrInter As Integer
Private ColOrEntry As Integer
Private ColSNum As Integer

Private CanNotAdd As Boolean

Private fIndex() As Integer
Private headIndex() As Integer
Private entryIndex() As Integer
Private HeadOrEntry() As Integer
Private fCount As Integer
Dim rCount As Integer '列数

Private Sub ComAdd_Click()
If MSHFlexGrid1.TextMatrix(1, 1) = "" Then
    MsgBox "当前查询结果为空，不能进行插入操作。", vbOKOnly, "金蝶提示"
    Exit Sub
End If

Dim FBillNos As String
Dim MultiCount, SelectedCount, tempCount As Integer
MultiCount = 0
SelectedCount = 1
tempCount = 0
FBillNos = ""

'处理异常
On Error GoTo Catch

For i = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(i, 0) = "Y" Then
        If InStr(1, FBillNos, MSHFlexGrid1.TextMatrix(i, 1), 1) < 1 Then '如果不包含则送货单号个数+1
            FBillNos = FBillNos + MSHFlexGrid1.TextMatrix(i, 1)
            MultiCount = MultiCount + 1
        Else
            SelectedCount = SelectedCount + 1
        End If
    End If
Next

If MultiCount = 1 Then
    For i = 1 To MSHFlexGrid1.Rows - 1
        If MSHFlexGrid1.TextMatrix(i, 1) = FBillNos Then
            tempCount = tempCount + 1
        End If
    Next
Else
    MsgBox "只能选择来自同一申请单号的数据", vbOKOnly, "金蝶提示"
    Exit Sub
End If

If IsSelectAll And Not (SelectedCount = tempCount) Then
    MsgBox "必须选择申请单号为 " + FBillNos + " 的所有数据", vbOKOnly, "金蝶提示"
    Exit Sub
End If

'If m_bill.GetHeadText(fIndex(1)) <> "" And m_bill.GetHeadText(fIndex(1)) <> FBillNos Then
    'MsgBox "当前外购入库单只能插入送货申请单号为" + m_bill.GetHeadText(fIndex(1)) + "的数据。", vbOKOnly, "金蝶提示"
    'Exit Sub
'End If

If m_bill.GetGridText(1, ColItemID) <> "" Then
    If MsgBox("继续执行将清除当前数据后再插入新的数据，是否继续？", vbYesNo) = vbYes Then
        '插入数据之前清除已添加的数据以防止重复插入相同的数据
        For t = 1 To m_bill.BillForm.vsEntrys.MaxRows
            'MsgBox t & " - " & m_bill.GetGridText(1, ColItemID)
            If m_bill.GetGridText(1, ColItemID) <> "" Then
                m_bill.BillForm.DelRow (1)
            End If
        Next
    Else
        Exit Sub
    End If
End If

Dim vItemID As String
Dim R As Integer
Dim S As Integer
Dim VarTmp As Variant
'表头
For S = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(S, 0) = "Y" Then
        For t = 1 To MSHFlexGrid1.Cols
            If headIndex(t) > 0 Then
                m_bill.SetHead headIndex(t), MSHFlexGrid1.TextMatrix(S, t)
            End If
        Next t
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
        For t = 1 To MSHFlexGrid1.Cols
           If entryIndex(t) > 0 Then
                m_bill.SetGridText R, entryIndex(t), MSHFlexGrid1.TextMatrix(S, t)
                'MsgBox Str(entryIndex(t)) + ":" + MSHFlexGrid1.TextMatrix(S, t)
            End If
        Next
    End If
Next S

Unload Me
Exit Sub

Catch:
    MsgBox "批量插入时出错:" & Err.Description
    '插入数据之前清除已添加的数据以防止重复插入相同的数据
    For t = 1 To m_bill.BillForm.vsEntrys.MaxRows
        If m_bill.GetGridText(1, ColItemID) <> "" Then
            m_bill.BillForm.DelRow (1)
        End If
    Next
End Sub

Private Sub ComEnter_Click()
If TxtBillNo.Text = "" Then
    MsgBox "申请单号不可为空!", vbOKOnly, "金蝶提示"
    Exit Sub
End If
Call ListLoad
End Sub

Private Sub ListLoad()
Dim rs As New ADODB.Recordset
Dim StrWhere As String
Dim StrSql As String
StrWhere = " where 1=1 "
'On Error Resume Next

If StoreProcName = "" Then
    MsgBox "加载过程失败,请关闭窗口重试."
    Exit Sub
End If

DBServer.CommandTimeout = 120
rs.Open "Exec " & StoreProcName & " @FBillNo='" + TxtBillNo.Text + "'", DBServer, adOpenKeyset
If rs.RecordCount > 0 Then
    MSHFlexGrid1.Rows = rs.RecordCount + 1
    fCount = rs.Fields.Count
Else
    MSHFlexGrid1.Rows = 2
End If

For R = 1 To rs.RecordCount
    MSHFlexGrid1.TextMatrix(R, 0) = ""
    For S = 1 To rCount
        MSHFlexGrid1.TextMatrix(R, S) = rs(S - 1)
    Next S
    rs.MoveNext
Next R

If UBound(headIndex) <= 0 Then
    ReDim headIndex(1 To fCount + 1)
    ReDim entryIndex(1 To fCount + 1)
    ReDim HeadOrEntry(1 To fCount + 1)
    Dim tName As String
    Dim headCount, entryCount, tCount, i, j As Integer
    headCount = UBound(m_bill.HeadCtl)
    entryCount = UBound(m_bill.EntryCtl)

    '获取K3字段数
    If headCount > entryCount Then tCount = headCount Else tCount = entryCount
    For i = 1 To fCount
        tName = UCase(rs.Fields(i - 1).Name)

        headIndex(i) = 0
        entryIndex(i) = 0
        For j = 1 To tCount
            If j <= headCount Then
                If UCase(m_bill.HeadCtl(j).FieldName) = tName Then
                    headIndex(i) = j
                    HeadOrEntry(i) = 0
                    Exit For
                End If
            Else
                Exit For
            End If
        Next j
    
        For k = 1 To tCount
            If k <= entryCount Then
                If UCase(m_bill.EntryCtl(k).FieldName) = tName Then
                    entryIndex(i) = k
                    HeadOrEntry(i) = 1
                    Exit For
                End If
            Else
                Exit For
            End If
        Next k
    Next
End If


If rs.RecordCount < 1 Then
    MsgBox "您的查询结果为空，请检查申请单号输入是否正确。", vbOKOnly, "金蝶提示"
End If
rs.Close
End Sub

Private Sub MSHFlexGrid1_Click()
If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0) = "Y" Then
    MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0) = ""
Else
    MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0) = "Y"
End If
End Sub

Private Sub Form_Load()

Call MshInit

FrmSel.Caption = SelType + "选择要入库的送货申请单"
Select Case SelType
Case "采购订单"
    FrmSel.Caption = SelType + "选择要入库的送货申请单"
    lblSelect(0).Caption = "送货申请单号"
    
Case "销售订单"
    FrmSel.Caption = SelType + "选择要出库的出货申请单"
    lblSelect(0).Caption = "出货申请单号"
Case Else
    FrmSel.Caption = "选择要出(入)库的出(送)货申请单"
    lblSelect(0).Caption = "送(出)货申请单号"
End Select

For i = 1 To UBound(m_bill.EntryCtl)
    If UCase(m_bill.EntryCtl(i).FieldName) = "FITEMID" Then
        ColItemID = i
    End If
    If UCase(m_bill.EntryCtl(i).FieldName) = "FORDERINTERID" Then
        ColOrInter = i
    End If
    'MsgBox m_bill.EntryCtl(i).FieldName
Next
ReDim headIndex(0 To 0)
CanNotAdd = False
End Sub


Private Sub MshInit()
MSHFlexGrid1.Clear

Dim rs As New ADODB.Recordset
Dim cIndex As Integer
rs.Open "Exec sp_SelectCols @FBillType='" & sFTranType & "'", DBServer, adOpenKeyset
rCount = rs.Fields.Count
MSHFlexGrid1.Cols = (rCount / 2) + 1

MSHFlexGrid1.TextMatrix(0, 0) = "标识"
MSHFlexGrid1.ColWidth(0) = 600

cIndex = 1
For R = 1 To rCount Step 2
    MSHFlexGrid1.TextMatrix(0, cIndex) = rs.Fields(R - 1)
    MSHFlexGrid1.ColWidth(cIndex) = rs.Fields(R)
    cIndex = cIndex + 1
Next R
rCount = rCount / 2
fCount = 0

rs.Close
End Sub

Private Sub TxtBillNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If TxtBillNo.Text = "" Then
            MsgBox "申请单号不可为空!", vbOKOnly, "金蝶提示"
            Exit Sub
        End If
        Call ListLoad
    End If
End Sub
