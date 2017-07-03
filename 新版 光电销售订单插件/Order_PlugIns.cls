VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Order_PlugIns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'定义插件对象接口. 必须具有的声明, 以此来获得事件
Private WithEvents m_BillTransfer  As k3BillTransfer.Bill
Attribute m_BillTransfer.VB_VarHelpID = -1
Private HeadEditNumber As Long
Public podate As Long
Private IsEdit As Boolean
 
Public Sub Show(ByVal oBillTransfer As Object)
 
    '接口实现
    '注意: 此方法必须存在, 请勿修改
    Set m_BillTransfer = oBillTransfer
 
End Sub

Private Sub Class_Terminate()
 
    '释放接口对象
    '注意: 此方法必须存在, 请勿修改
    Set m_BillTransfer = Nothing

End Sub
Private Sub m_BillTransfer_GridChange(ByVal Col As Long, ByVal Row As Long, ByVal Value As Variant, ByVal bNewBill As Boolean, Cancel As Boolean)
IsEdit = True
End Sub

Private Sub m_BillTransfer_HeadChange(ByVal CtlIndex As Long, ByVal Value As Variant, ByVal bNewBill As Boolean, Cancel As Boolean)
IsEdit = True
End Sub

Private Sub m_BillTransfer_BillInitialize()
 
'*************** 开始设置菜单 ***************

m_BillTransfer.AddUserMenuItem "CRM_订单导入", "扩展"
DBServer.ConnectionString = m_BillTransfer.Cnnstring
DBServer.Open
'*************** 结束设置菜单 ***************
 
End Sub


Private Sub m_BillTransfer_BillTerminate()
DBServer.Close
End Sub

Private Sub m_BillTransfer_UserMenuClick(ByVal Index As Long, ByVal Caption As String)
Dim R As Integer
'Dim vsEntrys As Variant
Dim THeadCtl As Variant
Dim TmpVal As Variant
Dim connStr As String

connStr = m_BillTransfer.Cnnstring

    Select Case Caption
    Case "CRM_订单导入"
        '此处添加处理 物料清单 菜单对象的 Click 事件
         Set FrmSel.m_bill = m_BillTransfer
        'For R = 1 To UBound(m_BillTransfer.HeadCtl)
        '    If UCase(m_BillTransfer.HeadCtl(R).FieldName) = "FSELTRANTYPE" Then
        '        FrmSel.SelType = m_BillTransfer.Head(R).Text
         '       Exit For
        '    End If
       ' Next R
       
       If InStr(connStr, "AIS20070122151122") > 0 Then
            FrmSel.myServer = "[192.168.100.205].[SaleOrder_platform_semi]."
            account = "semi"
        ElseIf InStr(connStr, "AIS20120508141651") > 0 Or InStr(connStr, "GDZB_CS_JLH20130828") > 0 Then
            FrmSel.myServer = "[192.168.100.205].[SaleOrder_platform]."
            account = "opto"
        ElseIf InStr(connStr, "AIS20060821075019") > 0 Then  '电子
            FrmSel.myServer = "[192.168.100.205].[SaleOrder_platform_ele]."
            account = "ele"
        Else
            MsgBox "没有合适的数据库连接", vbOKOnly, "信息中心提示"
            Exit Sub
        End If
        FrmSel.Show 1
    Case Else
    End Select

End Sub



