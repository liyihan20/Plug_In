VERSION 5.00
Begin VB.Form FrmSel 
   Caption         =   "����ģ��չ"
   ClientHeight    =   975
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   ScaleHeight     =   975
   ScaleWidth      =   5220
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
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
      Caption         =   "��ˮ��/�����ţ�"
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

'�ж���ˮ�źͶ����ŵĺϷ���
If Text1.Text = "" Then
    MsgBox "��������ˮ�Ż򶩵���", vbOKOnly, ICTIP
    Exit Sub
End If
   
'�ȶ�ȡ��Ҫ�������ֵ������recordset��
Set rs_value = m_Bill.K3Lib.GetData("select * from " + myServer + "[dbo].vw_CM_plugin where " + billNoName + " = '" + Text1.Text + "' or sys_no = '" + Text1.Text + "'")
If rs_value.RecordCount < 1 Then
    MsgBox "��ѯ������ص���", vbOKOnly, ICTIP
    Exit Sub
Else
    FrmSel.Hide
End If

'�ж϶����Ż���ˮ���Ƿ��Ѵ���
Set rs_exists = m_Bill.K3Lib.GetData("select 1 from " + tbName + " where " + billNoName + " = '" + rs_value(billNoName) + "'")
If rs_exists.RecordCount > 0 Then
    If MsgBox("�˶�����֮ǰ�ѵ����K3���Ƿ������", vbYesNo, ICTIP) = vbNo Then
        rs_value.Close
        rs_exists.Close
        Exit Sub
    End If
End If

Set rs_exists = m_Bill.K3Lib.GetData("select 1 from " + tbEntryName + " where FExplanation = '" + rs_value("sys_no") + "'")
If rs_exists.RecordCount > 0 Then
    If MsgBox("����ˮ��֮ǰ�ѵ����K3���Ƿ������", vbYesNo, ICTIP) = vbNo Then
        rs_value.Close
        rs_exists.Close
        Exit Sub
    End If
End If

'��Ӧ�ֶ�����ֵ���õ�����ͨ���ֶ���������ע�ⲻ����m_bill.BillHeads(1).BOSFields("FDate").Value���ַ�ʽ����ֵ��
'��Ϊ���ֻ�б�ͷ�ķǹ������Ͳ����ã����Ĺ������ͺͱ��������ֶζ����������ַ���
Set rs_text = m_Bill.K3Lib.GetData("select * from " + myServer + "[dbo].Sale_K3_table_description where bill_type='CM'")
If rs_text.RecordCount > 0 Then
    For i = 0 To rs_text.RecordCount - 1

        field_name = rs_text("field_en_name") '�ֶ���
        field_value = rs_value(field_name) '�ֶ�ֵ

        If rs_text("head_or_entry") = "h" Then '��ͷ
            m_Bill.SetFieldValue field_name, rs_value(field_name)
        ElseIf rs_text("head_or_entry") = "e" Then  '����
            m_Bill.SetFieldValue field_name, rs_value(field_name), 1
        End If
        rs_text.MoveNext
    Next i
End If
    
'��ȡ���ʣ������ܽ����ܽ���λ�ң���K3�����Զ�����
exchangeRate = m_Bill.GetFieldValue("FExchangerate")
m_Bill.SetFieldValue "FTotalAmountFor", rs_value("FContractAmount") - rs_value("FRebateAmount")
m_Bill.SetFieldValue "FTotalAmount", Math.Round((rs_value("FContractAmount") - rs_value("FRebateAmount")) * exchangeRate, 2)

MsgBox "���ݲ������", vbOKOnly, ICTIP
    
rs_text.Close
rs_value.Close
End Sub

Private Sub Form_Load()
ICTIP = "��Ϣ������ʾ"
End Sub
