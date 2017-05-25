VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHSample_PlugIns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is Industry Bill Interface Class, made by K3BOSPLUGINSWIZAED"
 
'定义插件对象接口. 必须具有的声明, 以此来获得事件
Private WithEvents m_BillTransfer   As k3BillTransfer.Bill
Attribute m_BillTransfer.VB_VarHelpID = -1
 
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

Private Sub m_BillTransfer_BillInitialize()
 
'*************** 开始设置菜单 ***************
 
    m_BillTransfer.AddUserMenuItem "出货样品入库", "扩展"
    DBServer.ConnectionString = m_BillTransfer.Cnnstring
    DBServer.Open
 
'*************** 结束设置菜单 ***************
 

End Sub

Private Sub m_BillTransfer_BillTerminate()
    DBServer.Close
End Sub

Private Sub m_BillTransfer_UserMenuClick(ByVal Index As Long, ByVal Caption As String)
    Dim connStr As String
    connStr = m_BillTransfer.Cnnstring
    
    Set m_bill = m_BillTransfer
    Select Case Caption
    Case "出货样品入库"
        If InStr(connStr, "AIS20070122151122") > 0 Then
            billNoSeg = "FEntrySelfA9744"
        ElseIf InStr(connStr, "AIS20120508141651") > 0 Or InStr(connStr, "GDZB_CS_JLH20130828") > 0 Then
            billNoSeg = "FEntrySelfA9747"
        Else
            MsgBox "没有合适的数据库连接", vbOKOnly, "信息中心提示"
            Exit Sub
        End If
        
        FmSel.Show 1
    Case Else
    End Select

End Sub
