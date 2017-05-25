VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TH_PlugIns"
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

m_BillTransfer.AddUserMenuItem "生成退货红字", "扩展"
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
    Select Case Caption
    Case "生成退货红字"
        If m_BillTransfer.SaveVect.Item(1).Value("ISRedBill") = False Then
            MsgBox "请先点击红字按钮", vbOKOnly, "信息中心提示"
        Else
            Set FrmSel.m_bill = m_BillTransfer
            FrmSel.Show 1
        End If
    Case Else
    End Select

End Sub

Private Sub EndSave(ByVal BillNo As String)
    MsgBox BillNo
End Sub
