VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICBL2PO_PlugIns"
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
 
    m_BillTransfer.AddUserMenuItem "CRM备料单导入", "扩展"
 
'*************** 结束设置菜单 ***************
 
    DBServer.ConnectionString = m_BillTransfer.Cnnstring
    DBServer.Open

End Sub


Private Sub m_BillTransfer_BillTerminate()
 
    'TODO: 请在此处添加代码响应事件 BillTerminate
    DBServer.Close

End Sub

Private Sub m_BillTransfer_UserMenuClick(ByVal Index As Long, ByVal Caption As String)
 
    'TODO: 请在此处添加代码响应事件 UserMenuClick
    Dim connStr As String
    connStr = m_BillTransfer.Cnnstring
    
    Select Case Caption
    Case "CRM备料单导入"
        
        Set FmSel.m_bill = m_BillTransfer
        myServer = "[192.168.100.205].[SaleOrder_platform_semi]."  '以下是半导体账套
        If InStr(connStr, "AIS20070514084102") > 0 Then
            account = "LTD"
        ElseIf InStr(connStr, "AIS20070508145209") > 0 Then
            account = "TFT"
        ElseIf InStr(connStr, "AIS20070208163242") > 0 Then
            account = "MDS"
        ElseIf InStr(connStr, "AIS20070731165027") > 0 Then
            account = "ADD"
        ElseIf InStr(connStr, "AIS20070927102604") > 0 Then
            account = "OLED"
        ElseIf InStr(connStr, "AIS20070731163322") > 0 Or InStr(connStr, "SMD_CJL20170413") > 0 Then
            account = "SMD"
        ElseIf InStr(connStr, "AIS20070807151717") > 0 Then
            account = "LCD"
        Else
            myServer = "[192.168.100.205].[SaleOrder_platform]."  '以下是光电账套
            If InStr(connStr, "AIS20160729104317") > 0 Then
                account = "AITD"
            ElseIf InStr(connStr, "AIS20120419162024") > 0 Then
                account = "TPD"
            ElseIf InStr(connStr, "AIS20070208163242") > 0 Then
                account = "FPI"
            ElseIf InStr(connStr, "AIS20120419162202") > 0 Then
                account = "CCM"
            ElseIf InStr(connStr, "AIS20120419162510") > 0 Then
                account = "TDD"
            ElseIf InStr(connStr, "AIS20170316164445") > 0 Then
                account = "STC"
            ElseIf InStr(connStr, "AIS20140716104140") > 0 Then
                account = "背光"
            Else
                MsgBox "没有合适的数据库连接", vbOKOnly, "信息中心提示"
                Exit Sub
            End If
        End If
        
        FmSel.Show 1
        
    Case Else
    End Select

End Sub
