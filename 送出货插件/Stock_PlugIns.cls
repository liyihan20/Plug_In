VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stock_PlugIns"
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

m_BillTransfer.AddUserMenuItem "送(出)货申请单", "扩展"
m_BillTransfer.AddUserMenuItem "生成送货单", "扩展"

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
    Dim BillTypeRS As ADODB.Recordset
    FrmSel.sFTranType = 0
    Dim iFTranType As Integer

'    For R = 1 To UBound(m_BillTransfer.HeadCtl)
'        If InStr(m_BillTransfer.HeadCtl(R).FieldName, "HeadSelfA01") > 0 Then '外购入库
'            iFTranType = 1
'            Exit For
'        ElseIf InStr(m_BillTransfer.HeadCtl(R).FieldName, "HeadSelfB01") > 0 Then '销售出库
'            iFTranType = 21
'            Exit For
'        ElseIf InStr(m_BillTransfer.HeadCtl(R).FieldName, "HeadSelfZIN") > 0 Then '虚仓入库
'            iFTranType = 6
'            Exit For
'        ElseIf InStr(m_BillTransfer.HeadCtl(R).FieldName, "HeadSelfZOU") > 0 Then '虚仓出库
'            iFTranType = 26
'            Exit For
'        ElseIf InStr(m_BillTransfer.HeadCtl(R).FieldName, "HeadSelfA97") > 0 Then '其他入库
'            iFTranType = 10
'            Exit For
'        ElseIf InStr(m_BillTransfer.HeadCtl(R).FieldName, "HeadSelfB09") > 0 Then '其他出库
'            iFTranType = 29
'            Exit For
'        End If
'    Next R
    
    iFTranType = m_BillTransfer.SaveVect.Item(1).Value("FTransType")
    
    'MsgBox Str(iFTranType)
    
    Select Case Caption
    Case "送(出)货申请单"
        Set FrmSel.m_bill = m_BillTransfer
        
        On Error GoTo Catch1
        
        For R = 1 To UBound(m_BillTransfer.HeadCtl)
            If UCase(m_BillTransfer.HeadCtl(R).FieldName) = "FSELTRANTYPE" Then
                FrmSel.SelType = m_BillTransfer.Head(R).Text
                Exit For
            End If
        Next R
        
        FrmSel.sFTranType = iFTranType
        FrmSel.IsSelectAll = False
        If iFTranType = 21 Or iFTranType = 29 Or iFTranType = 41 Then
            FrmSel.TxtBillNo.Text = Right("" & Year(DateTime.Date), 2) & Right("0" & Month(DateTime.Date), 2)
        End If

        Set BillTypeRS = New ADODB.Recordset
        BillTypeRS.Open "Exec sp_SelectBillType @FBillNo='', @FTranType=" & iFTranType, DBServer, adOpenKeyset
        If BillTypeRS.RecordCount > 0 Then
            FrmSel.sFTranType = BillTypeRS(0)
            FrmSel.StoreProcName = BillTypeRS(1)
            If BillTypeRS(2) = 1 Then
                FrmSel.IsSelectAll = True
            Else
                FrmSel.IsSelectAll = False
            End If
        End If
        BillTypeRS.Close
                    
        If FrmSel.StoreProcName = "" Then
            MsgBox "当前单据不能使用[送(出)申请单]功能", vbOKOnly, "信息中心提示"
            Exit Sub
        Else
            FrmSel.Show 1
        End If
        
    Case "生成送货单"
        Dim FBillNo As String
        Dim FSRBillNo As String
        
        Dim ProcName As String
        Dim SelfFieldName As String
        Dim BillNoFieldName As String
        Dim GenDeliveryProcRS As ADODB.Recordset
        
        On Error GoTo Catch2
        
        Set GenDeliveryProcRS = New ADODB.Recordset
        If DBServer.State = adStateClosed Then
            MsgBox "Open connection"
            DBServer.Open
        End If
        GenDeliveryProcRS.Open "Exec sp_SelectDeliveryProc @FTranType=" & iFTranType, DBServer, adOpenKeyset
        If GenDeliveryProcRS.RecordCount > 0 Then
            ProcName = GenDeliveryProcRS(0)
            BillNoFieldName = GenDeliveryProcRS(1)
            SelfFieldName = GenDeliveryProcRS(2)
        End If
        GenDeliveryProcRS.Close
        
        For R = 1 To UBound(m_BillTransfer.HeadCtl)
            '自定义字段名
            If UCase(m_BillTransfer.HeadCtl(R).FieldName) = UCase(SelfFieldName) Then
                FSRBillNo = m_BillTransfer.Head(R).Text
                Exit For
            End If
        Next R
        For R = 1 To UBound(m_BillTransfer.HeadCtl)
            '单据编号字段名
            If UCase(m_BillTransfer.HeadCtl(R).FieldName) = UCase(BillNoFieldName) Then
                FBillNo = m_BillTransfer.Head(R).Text
                Exit For
            End If
        Next R

        If ProcName = "" Or FBillNo = "" Then
            MsgBox "此类单据不能生成送货单", vbOKOnly, "信息中心提示"
        Else
            GenDelivery ProcName, FBillNo, FSRBillNo
        End If
    Case Else
    End Select
    Exit Sub
Catch1:
    MsgBox "调用申请数据查询窗口失败:" & Err.Description
    Exit Sub
Catch2:
    MsgBox "生成送货单失败,出现异常:" & Err.Description
    Exit Sub
End Sub

Private Sub GenDelivery(ProcName As String, FBillNo As String, FSRBillNo As String)
    Dim bSucc As Boolean
    Dim FDeliveryNo As String
    Dim strMsg As String
    Dim GenRS As ADODB.Recordset

    Dim cmd As New ADODB.Command
    With cmd
        .ActiveConnection = DBServer
        If .ActiveConnection.State = adStateClosed Then
            .ActiveConnection.Open
        End If
        .CommandType = adCmdStoredProc
        .CommandText = ProcName
        .Parameters.Append cmd.CreateParameter("FBillNo", adVarChar, adParamInput, 80, FBillNo)
        .Parameters.Append cmd.CreateParameter("FSRBillNo", adVarChar, adParamInput, 80, FSRBillNo)
        .Parameters.Append cmd.CreateParameter("FDeliveryNo", adVarChar, adParamOutput, 80, "")
        .Parameters.Append cmd.CreateParameter("FEmail", adVarChar, adParamOutput, 80, "")
        .Parameters.Append cmd.CreateParameter("FMsg", adVarChar, adParamOutput, 500, "")
        .Parameters.Append cmd.CreateParameter("FSuccess", adInteger, adParamOutput, 8, 0)
    End With
    cmd.Execute
    If cmd.Parameters("FSuccess").Value = 1 Then
        FDeliveryNo = cmd.Parameters("FDeliveryNo").Value
        '发送邮件时出现异常
        On Error GoTo Catch
        'SendMail cmd.Parameters("FEmail").Value, FBillNo, FDeliveryNo, FSRBillNo
        MsgBox cmd.Parameters("FMsg").Value & ". 单号:" & FDeliveryNo
    Else
        MsgBox cmd.Parameters("FMsg").Value
    End If
    Exit Sub
Catch:
    MsgBox "发送邮件失败:" & Err.Description
End Sub

Sub SendMail(EmailAddr As String, FBillNo As String, FDeliveryNo As String, FSRBillNo As String)
    Dim strHTML As String
    Dim jmail As Object
    Set jmail = CreateObject("jmail.Message")
    jmail.Charset = "gb2312"
    jmail.ContentType = "text/html"
    jmail.Priority = 3 '邮件状态,1-5 1为最高
    jmail.MailServerUserName = "pur@truly.com.cn" 'Email帐号
    jmail.MailServerPassWord = "Truly3006" 'Email密码
    
    jmail.FromName = "出货管理系统" '发信人姓名
    jmail.From = "pur@truly.com.cn" '发邮件地址地址
    
    jmail.Subject = "已生成送货单" & FBillNo '主题
    Dim arrEmail() As String
    arrEmail = Split(EmailAddr, ";")
    For i = 0 To UBound(arrEmail)
        If arrEmail(i) <> "" Then
            jmail.AddRecipient arrEmail(i) '收信人地址
        End If
    Next
    strHTML = "各位好!<br />"
    strHTML = strHTML & "&nbsp;&nbsp;&nbsp;&nbsp;已通过销售(其他)出库单" & FBillNo & "生成送货单" & FDeliveryNo & ",对应的出货申请单号为" & FSRBillNo
    strHTML = strHTML & "<br />&nbsp;&nbsp;&nbsp;&nbsp;请物流中心工作人员及时打印送货单."
    strHTML = strHTML & "<br /><br /><br /><br />注：本邮件由出货管理系统自动发送，请勿直接回复。您之所以收到本邮件，是因为您正在使用信利国际有限公司出货管理系统。"
    strHTML = strHTML & "<br /><br /><br />Truly International Holdings Limited"
    strHTML = strHTML & "<br />信利国际有限公司"
    strHTML = strHTML & "<br /><STRONG><FONT face=Arial size=4 color=#ff0000><EM>Truly</EM></FONT>"
    strHTML = strHTML & "<FONT size=3>&nbsp;&nbsp;</FONT><FONT face=Arial size=4 color=#008000><EM>Truview</EM></FONT>"
    strHTML = strHTML & "<FONT size=3>&nbsp;&nbsp;</FONT><FONT face=Arial size=4 color=#0000ff><EM>Trutouch</EM></FONT></STRONG><br />"
    jmail.Body = strHTML
    jmail.Send ("smtp.truly.com.cn")
    Set jmail = Nothing
End Sub
