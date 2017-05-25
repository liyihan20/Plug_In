VERSION 5.00
Begin VB.Form FrmSel 
   Caption         =   "开改模扩展"
   ClientHeight    =   975
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   ScaleHeight     =   975
   ScaleWidth      =   5220
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "流水号/订单号："
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "FrmSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents m_Bill  As BillEvent
Attribute m_Bill.VB_VarHelpID = -1
Public myServer As String
Public tbName As String
Public tbEntryName As String
Public billNoName As String
Private ICTIP As String



Private Sub Command1_Click()
Dim rs_value As New ADODB.Recordset
Dim rs_text As New ADODB.Recordset
Dim rs_exists As New ADODB.Recordset
Dim i As Integer
Dim field_name As String
Dim field_value As String
Dim exchangeRate As Single

'判断流水号和订单号的合法性
If Text1.Text = "" Then
    MsgBox "请输入流水号或订单号", vbOKOnly, ICTIP
    Exit Sub
End If
   
'先读取需要插入表单的值，放在recordset中
Set rs_value = m_Bill.K3Lib.GetData("select * from " + myServer + "[dbo].vw_CM_plugin where " + billNoName + " = '" + Text1.Text + "' or sys_no = '" + Text1.Text + "'")
If rs_value.RecordCount < 1 Then
    MsgBox "查询不到相关单据", vbOKOnly, ICTIP
    Exit Sub
Else
    FrmSel.Hide
End If

'判断订单号或流水号是否已存在
Set rs_exists = m_Bill.K3Lib.GetData("select 1 from " + tbName + " where " + billNoName + " = '" + rs_value(billNoName) + "'")
If rs_exists.RecordCount > 0 Then
    If MsgBox("此订单号之前已导入过K3，是否继续？", vbYesNo, ICTIP) = vbNo Then
        rs_value.Close
        rs_exists.Close
        Exit Sub
    End If
End If

Set rs_exists = m_Bill.K3Lib.GetData("select 1 from " + tbEntryName + " where FExplanation = '" + rs_value("sys_no") + "'")
If rs_exists.RecordCount > 0 Then
    If MsgBox("此流水号之前已导入过K3，是否继续？", vbYesNo, ICTIP) = vbNo Then
        rs_value.Close
        rs_exists.Close
        Exit Sub
    End If
End If

'对应字段名将值设置到表单，通过字段名关联，注意不能用m_bill.BillHeads(1).BOSFields("FDate").Value这种方式设置值。
'因为这个只有表头的非关联类型才适用，表够的关联类型和表体所有字段都不能用这种方法
Set rs_text = m_Bill.K3Lib.GetData("select * from " + myServer + "[dbo].Sale_K3_table_description where bill_type='CM'")
If rs_text.RecordCount > 0 Then
    For i = 0 To rs_text.RecordCount - 1

        field_name = rs_text("field_en_name") '字段名
        field_value = rs_value(field_name) '字段值

        If rs_text("head_or_entry") = "h" Then '表头
            m_Bill.SetFieldValue field_name, rs_value(field_name)
        ElseIf rs_text("head_or_entry") = "e" Then  '表体
            m_Bill.SetFieldValue field_name, rs_value(field_name), 1
        End If
        rs_text.MoveNext
    Next i
End If
    
'读取汇率，设置总金额和总金额（本位币），K3不会自动设置
exchangeRate = m_Bill.GetFieldValue("FExchangerate")
m_Bill.SetFieldValue "FTotalAmountFor", rs_value("FContractAmount") - rs_value("FRebateAmount")
m_Bill.SetFieldValue "FTotalAmount", Math.Round((rs_value("FContractAmount") - rs_value("FRebateAmount")) * exchangeRate, 2)

MsgBox "数据插入完成", vbOKOnly, ICTIP
    
rs_text.Close
rs_value.Close
End Sub

Private Sub Form_Load()
ICTIP = "信息中心提示"
End Sub
