Attribute VB_Name = "Module1"
Option Explicit

Public DBServer As New ADODB.Connection
Public m_bill As k3BillTransfer.Bill
Public billNoSeg As String  '其它入库单的出库申请单号自定义字段名
