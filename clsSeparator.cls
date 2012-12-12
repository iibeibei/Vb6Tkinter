VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSeparator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'�ָ����࣬�������úͱ�����Ӧ�����Ժ�ֵ�������ɶ�ӦPYTHON��TKinter����

Private m_dicTotal As Dictionary                                                '����ȫ�������ԣ�����Ĭ��ֵ
Private m_Base As clsBaseControl                                                '�����ؼ���
Private m_CanbeOutByMainForm As Boolean

'���PYTHON����,
'sCmdFunc: ����������¼������ص�����;
'rel:�Ƿ�ʹ���������,
'oop:�Ƿ�ʹ�����������
'usettk:�Ƿ�ʹ��TTK������չ
Public Function toString(ByRef sCmdFunc As cStrBuilder, rel As Boolean, oop As Boolean, usettk As Boolean) As String
    If Len(m_Base("orient")) = 0 Then m_Base("orient") = "'horizontal'"
    toString = m_Base.toString(sCmdFunc, rel, oop, usettk)
End Function


'���������Ҫ���ϵ������������ʼ��������
Public Sub InitConfig(o As Object, parentWidth As Long, parentHeight As Long, Optional dMethods As Dictionary)
    
    Dim W As Long, H As Long, X1 As Long, X2 As Long, Y1 As Long, Y2 As Long
    
    m_Base.Name = o.Properties("Name")
    
    If o.Properties("X1") <= o.Properties("X2") Then
        X1 = o.Properties("X1")
        Y1 = o.Properties("Y1")
        X2 = o.Properties("X2")
        Y2 = o.Properties("Y2")
    Else
        X1 = o.Properties("X2")
        Y1 = o.Properties("Y2")
        X2 = o.Properties("X1")
        Y2 = o.Properties("Y1")
    End If
    
    m_dicTotal("x") = m_Base.toPixelX(X1)
    m_dicTotal("y") = m_Base.toPixelY(Y1)
    
    '�ж���ˮƽ���Ǵ�ֱ�ģ����LINE�ؼ���б�ģ������ˮƽͶӰ�ϳ�����Ϊˮƽ����ֱ���ж���Ȼ
    W = Abs(X2 - X1)
    H = Abs(Y2 - Y1)
    If W >= H Then
        m_dicTotal("width") = m_Base.toPixelX(W)
        m_dicTotal("height") = o.Properties("BorderWidth")
        m_dicTotal("relx") = Format(X1 / parentWidth, "0.###")
        m_dicTotal("rely") = Format(Y1 / parentHeight, "0.###")
        m_dicTotal("relwidth") = Format(W / parentWidth, "0.###")
        m_dicTotal("relheight") = Format(o.Properties("BorderWidth") / m_Base.toPixelY(parentHeight), "0.####")
        m_dicTotal("orient") = "'horizontal'"
    Else
        m_dicTotal("width") = o.Properties("BorderWidth")
        m_dicTotal("height") = m_Base.toPixelX(H)
        m_dicTotal("relx") = Format(X1 / parentWidth, "0.###")
        m_dicTotal("rely") = Format(Y1 / parentHeight, "0.###")
        m_dicTotal("relwidth") = Format(o.Properties("BorderWidth") / m_Base.toPixelX(parentWidth), "0.####")
        m_dicTotal("relheight") = Format(H / parentHeight, "0.###")
        m_dicTotal("orient") = "'vertical'"
    End If
    
    m_dicTotal("bg") = TranslateColor(o.Properties("BorderColor"))
    
    '��Щ��Ĭ��ѡ���
    m_Base("x") = m_dicTotal("x")
    m_Base("y") = m_dicTotal("y")
    m_Base("width") = m_dicTotal("width")
    m_Base("height") = m_dicTotal("height")
    m_Base("relx") = m_dicTotal("relx")
    m_Base("rely") = m_dicTotal("rely")
    m_Base("relwidth") = m_dicTotal("relwidth")
    m_Base("relheight") = m_dicTotal("relheight")
    m_Base("orient") = m_dicTotal("orient")
    'If o.Properties("BorderColor") <> vbWindowText Then
    m_Base("bg") = m_dicTotal("bg")
    
End Sub

'��������ֵ�Ŀ���ֵ�б�
'����ֵ��0-û�п�ѡֵ��1-��һ���ϸ����ƵĿ�ѡֵ�б���2-���ṩ�Ŀ�ѡֵ�б��⣬�������ֶ���������ֵ
'�����sa()��ѡֵ�б�����
Public Function GetAttrValueList(sAttr As String, ByRef sa() As String) As Long
    GetAttrValueList = m_Base.GetAttrValueList(sAttr, sa)
End Function

Public Function Tips(sAttr As String) As String
    Tips = m_Base.Tips(sAttr)
End Function

'���û�ѡ������ø��µ������У�����Ϊʹ��"|"�ָ�ĺܶ������/ֵ��
Public Sub SetConfig(sAttrs As String)
    m_Base.SetConfig sAttrs
End Sub

'�޸Ļ����ӵ������������/ֵ��"|"�ָ�
Public Sub SetSingleConfig(sAttr As String)
    m_Base.SetSingleConfig sAttr
End Sub

'����һ�����ϣ�ÿ����Ŀ��Ԫ��"������|ֵ|�Ƿ�Ĭ��ѡ��"
'�����������������������Բ����б���
Public Function Allitems() As Collection
    Dim re As Collection, k As Variant, ks As Collection
    
    Set re = New Collection
    
    '��׼����
    Set ks = m_dicTotal.Keys
    For Each k In ks
        If Len(m_Base(k)) Then
            re.Add k & "|" & m_Base(k) & "|1"
        Else
            re.Add k & "|" & m_dicTotal(k) & "|0"
        End If
    Next
    
    '�û����ӵ��Զ������������еĻ���
    Set ks = m_Base.Keys
    For Each k In ks
        If Not m_dicTotal.Exists(k) Then
            re.Add k & "|" & m_Base(k) & "|1"
        End If
    Next
    
    Set Allitems = re
    
End Function

Private Sub Class_Initialize()
    Set m_dicTotal = New Dictionary
    Set m_Base = New clsBaseControl
    m_Base.ctlType = "Separator"
    m_Base.StyleName = "TSeparator"
    m_CanbeOutByMainForm = True
End Sub

Private Sub Class_Terminate()
    Set m_dicTotal = Nothing
    Set m_Base = Nothing
End Sub

Public Property Let Parent(s As String)
    m_Base.Parent = s
End Property

Public Property Get Parent() As String
    Parent = m_Base.Parent
End Property

Public Property Get Name() As String
    Name = m_Base.Name
End Property

'���ڸı���Ĭ�϶�Ӧ��widget���ͣ��޸�widget���ͺ�ע�������б��ĺϷ���
Public Function SetWidgetType(sType As String, sStyleName As String)
    m_Base.ctlType = sType
    m_Base.StyleName = sStyleName
End Function

'ȷ�������������ܷ������toString()���������룬Ĭ��ΪTrue������ΪFalse˵�����������������ô���
Public Property Get EnableOutByMainForm() As Boolean
    EnableOutByMainForm = m_CanbeOutByMainForm
End Property
Public Property Let EnableOutByMainForm(bEnable As Boolean)
    m_CanbeOutByMainForm = bEnable
End Property

'�������л�����
Public Function Serializer(vSer As clsSerialization)
    vSer.Serializer m_Base
End Function

Public Function Deserializer(vSer As clsSerialization)
    vSer.Deserializer m_Base
End Function

Public Property Get Description() As String
    Description = L("l_DescSeparator", "�ָ������󣬶�Ӧ��Tkinter.ttk��Separator�ؼ���ֻ֧��ˮƽ�ߺʹ�ֱ�ߡ�")
End Property

Public Property Let ScaleMode(nV As Long)
    m_Base.ScaleMode = nV
End Property

'����ģ��Ƚ�����ĺ�����ʵ�������ж����������������ϵ
'�ñ��������һ������Ƚϣ��߼����Ϊ'������-��һ������'
'����ֵ���壺
'<0����ʾ��������Ҫ����һ������ǰ�������
'=0����ʾ���߽�û��������ϵ������ǰ��˳����Ӱ��
'>0����һ������Ҫ��������롣
'������߼�����������صĳ���
Public Function Compare(ByRef Obj As Object) As Long
    
    If Parent = Obj.Name Then '���ؼ����������
        Compare = 1
    ElseIf Obj.Parent = Name Then
        Compare = -1
    ElseIf Parent = WTOP And Obj.Parent <> WTOP Then '����ؼ������
        Compare = -1
    ElseIf Parent <> WTOP And Obj.Parent = WTOP Then
        Compare = 1
    Else
        Compare = 0
    End If
End Function
