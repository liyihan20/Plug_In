VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOS_BillInsert_PlugIns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is BillEvent Interface Class, made by K3BOSPLUGINSWIZAED"
 
'定义 BillEvent 接口. 必须具有的声明, 以此来获得事件
Private WithEvents m_BillInterface  As BillEvent
Attribute m_BillInterface.VB_VarHelpID = -1
 
Public Sub Show(ByVal oBillInterface As Object)
 
    'BillEvent 接口实现
    '注意: 此方法必须存在, 请勿修改
    Set m_BillInterface = oBillInterface
 
End Sub

Private Sub Class_Terminate()
 
    '释放接口对象
    '注意: 此方法必须存在, 请勿修改
    Set m_BillInterface = Nothing

End Sub

Private Sub m_BillInterface_MenuBarClick(ByVal BOSTool As K3ClassEvents.BOSTool, Cancel As Boolean)
    Dim tableName As String

    tableName = UCase(m_BillInterface.tableName)
    
    Select Case BOSTool.ToolName
    Case "Mnu_Insert"
        '此处添加处理Mnu_Insert 菜单对象的 Click 事件
        Set FrmSel.m_Bill = m_BillInterface
        '用表名来判断是光电账套还是半导体账套，因为光电和半导体的表明是不一样的，自定义单据也没有constr这个属性
        If InStr(tableName, UCase("t_RPContract3")) Then
            FrmSel.myServer = "[192.168.100.205].[SaleOrder_platform_semi]."
            FrmSel.tbEntryName = "t_RPContract3Entry"  '表体名称
            FrmSel.billNoName = "FText5"  '单号字段名
        ElseIf InStr(tableName, UCase("t_BosContract")) Then
           FrmSel.myServer = "[192.168.100.205].[SaleOrder_platform]."
           FrmSel.tbEntryName = "t_BOSContractEntry1"
           FrmSel.billNoName = "FOldBillNo"
        Else
            MsgBox "没有合适的数据库连接", vbOKOnly, "信息中心提示"
            Exit Sub
        End If
        FrmSel.tbName = m_BillInterface.tableName '表头名称
        
        FrmSel.Show 1
    Case Else
    End Select

End Sub

Private Sub m_BillInterface_MenuBarInitialize(ByVal oMenuBar As K3ClassEvents.MenuBar)
 
    'TODO: 请在此处添加代码响应事件
 
 
    Dim oTool   As K3ClassEvents.BOSTool
    Dim oBand   As K3ClassEvents.BOSBand
 
'*************** 开始新增 BOS 菜单 ***************
 
    '新增 Mnu_Insert 菜单对象，并设置属性
    Set oTool = oMenuBar.BOSTools.Add("Mnu_Insert")
    With oTool
        .Caption = "开改模扩展"
        .ShortcutKey = 0
        .Visible = True
        .Enabled = True
        .BeginGroup = False
    End With
 
    Set oBand = oMenuBar.BOSBands("mnuEdit")
    oBand.BOSTools.InsertAfter "mnuEditUnMultiCheck", oTool '将菜单对象插入指定一级菜单
 
 
'*************** 结束新增 BOS 菜单 ***************
 

End Sub
