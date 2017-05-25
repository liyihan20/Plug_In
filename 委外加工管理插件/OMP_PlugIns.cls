VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OMP_PlugIns"
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
Private trantype As Integer
Private viewName As String
 
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
trantype = m_BillTransfer.SaveVect.Item(1).Value("FTransType")

Select Case trantype
Case 5
    m_BillTransfer.AddUserMenuItem "委外平台_委外入库", "扩展"
    viewName = "VwSendProductPlugin"
Case 6
    m_BillTransfer.AddUserMenuItem "委外平台_虚仓入库", "扩展"
    viewName = "VwReturnLeftOutPlugin"
Case 28
    m_BillTransfer.AddUserMenuItem "委外平台_红字委外出库", "扩展"
    viewName = "VwReturnMaterialsPlugin"
Case Else
End Select

DBServer.ConnectionString = m_BillTransfer.Cnnstring
DBServer.Open
'*************** 结束设置菜单 ***************
 
End Sub


Private Sub m_BillTransfer_BillTerminate()
DBServer.Close
End Sub

Private Sub m_BillTransfer_UserMenuClick(ByVal Index As Long, ByVal Caption As String)

If Caption = "委外平台_红字委外出库" Then
     If m_BillTransfer.SaveVect.Item(1).Value("ISRedBill") = False Then
            MsgBox "请先点击红字按钮", vbOKOnly, "信息中心提示"
            Exit Sub
     End If
End If

Set FrmSel.m_bill = m_BillTransfer
         FrmSel.trantype = trantype
         FrmSel.viewName = viewName
         FrmSel.Show 1
End Sub

Private Sub EndSave(ByVal billNo As String)
    MsgBox billNo
End Sub
