VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmBgRpt 
   Caption         =   "�Ƽ�Ԥ�㱨�����ɹ���"
   ClientHeight    =   6816
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   10176
   OleObjectBlob   =   "FrmBgRpt.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "FrmBgRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim bookMarkName() As String
Dim dic As Object
Dim dic2 As Object
Dim dic3 As Object
Dim dic4 As Object
Dim dic51 As Object
Dim dic52 As Object
Dim dic53 As Object
Dim dic7 As Object
Dim dic8 As Object
Dim Egg As String

Dim oRegEgg As Object

Public dataSource As String
Public KeyConts As Object



Public Sub DimBMN()
 
    ReDim bookMarkName(30) As String
    
    bookMarkName(1) = "��Ŀ����"   'ΪbookMarkName��ֵ
    bookMarkName(2) = "ί�е�λ"
    bookMarkName(3) = "��ʼʱ��"
    bookMarkName(4) = "��������"
    bookMarkName(5) = "��˾�����"
    bookMarkName(6) = "���̸ſ�"
    bookMarkName(7) = "���Ʒ�Χ"
    bookMarkName(8) = "��������"
    bookMarkName(9) = "���Ʒ���"
    bookMarkName(10) = "���ƽ��"
    bookMarkName(11) = "����˵��"
    bookMarkName(12) = "����"
    bookMarkName(13) = "���ű����"
    
    
    Set oRegEgg = CreateObject("VBscript.RegExp")
    oRegEgg.Global = True
    oRegEgg.Pattern = "521314"
    
'    'Ԥ��д��������ݣ����ڳ���5�����
'    Set bkMark = ActiveDocument.Bookmarks(bookMarkName(8)).Range  '������ǩ��bookmark1������ֵ��bkMark
'    bkMark.Select 'ѡ��bkMark��ǩ��Ӧ���ı�
'    TBRptAcc.Text = bkMark.Text  '��������
        
    
End Sub



Private Sub CheckBox2_Click()
    Call CheckBoxClickNew(CheckBox2, TBRptRg, dic3, "����24", Page3 + "." + "����24", TextValue3, MultiPage1, 2)
End Sub

Private Sub modBreak4_Click()
     Call addLineBreak(TextValue4, dic4, TBRptAcc)
End Sub

Private Sub modBreak51_Click()
     Call addLineBreak(TextValue51, dic51, TBRptMtd)
End Sub

Private Sub modBreak52_Click()
     Call addLineBreak(TextValue52, dic52, TBRptMtd)
End Sub

Private Sub modBreak53_Click()
     Call addLineBreak(TextValue53, dic53, TBRptMtd)
End Sub

Private Sub modBreak7_Click()
     Call addLineBreak(TextValue7, dic7, TBRptOth)
End Sub

Private Sub modBreak8_Click()
     Call addLineBreak(TextValue8, dic8, AuditorText, True)
End Sub



'������Ŀ
Private Sub CommandButton95_Click()

    Call ModfStringFromCustomAttrib(ListBox1.Text, TextBox37.Text)

End Sub

Private Sub CommandButton96_Click()
    Call ModfAllCustomAttrib
    
End Sub



Private Sub ListBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    'Call HookListBoxScroll

End Sub

Private Sub ListBox1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Dim attr As String
    attr = ReadStringFromCustomAttrib(ListBox1.Text)
    
    TextBox37.Text = attr
    Label35.Caption = ListBox1.Text
    
    
End Sub



Private Sub UserForm_Terminate()

    Dim bkMark As Range '����һ��bkMark��range�����һ��bookmarkName�ַ�������


    Call userWriteBookMarkName(bookMarkName(1), ��Ŀ����, True)
    Call userWriteBookMarkName(bookMarkName(2), ί�е�λ, True)
    Call userWriteBookMarkName(bookMarkName(3), ��ʼʱ��, True)
    Call userWriteBookMarkName(bookMarkName(4), ��������, True)
    Call userWriteBookMarkName(bookMarkName(5), ��˾�����, True)
    Call userWriteBookMarkName(bookMarkName(13), ���ű����, True)

    Call userWriteBookMarkName(bookMarkName(6), TxtBasicInfo, True)
    Call ReplaceTextwithCrossRef(bookMarkName(6), bookMarkName(1), TxtBasicInfo)
    Call ReplaceTextwithCrossRef(bookMarkName(6), bookMarkName(2), TxtBasicInfo)

    Call userWriteBookMarkName(bookMarkName(7), TBRptRg, True)
    Call ReplaceTextwithCrossRef(bookMarkName(7), bookMarkName(1), TBRptRg)

    Call userWriteBookMarkName(bookMarkName(8), TBRptAcc, True)
    Call ReplaceTextwithCrossRef(bookMarkName(8), bookMarkName(1), TBRptAcc)

    Call userWriteBookMarkName(bookMarkName(9), TBRptMtd, True)
    Call userWriteBookMarkName(bookMarkName(10), TBRptCkl, True)
    Call userWriteBookMarkName(bookMarkName(11), TBRptOth, True)

    Call userWriteBookMarkName(bookMarkName(12), AuditorText, True)
    Call ReplaceTextwithCrossRef(bookMarkName(12), bookMarkName(1), AuditorText)

    UpAF


End Sub


Private Sub UserForm_Initialize()


    '��ʼ��������д������ݣ�
    Set dic = CreateObject("Scripting.Dictionary")
    Set dic2 = CreateObject("Scripting.Dictionary")
    Set dic3 = CreateObject("Scripting.Dictionary")
    Set dic4 = CreateObject("Scripting.Dictionary")
    Set dic51 = CreateObject("Scripting.Dictionary")
    Set dic52 = CreateObject("Scripting.Dictionary")
    Set dic53 = CreateObject("Scripting.Dictionary")
    Set dic7 = CreateObject("Scripting.Dictionary")
    Set dic8 = CreateObject("Scripting.Dictionary")
    
    DimBMN

    
    
    '������Ϣ
    Label29.FontSize = 6
    
    '��������� True �򿪣�False �ر�
    
    MultiPage1.Pages.Item(8).Visible = False
    
    '������ϢԤд��'
    
'    Call bookMarkNameCheckExistWrite(bookMarkName(1), ��Ŀ����)
'    Call bookMarkNameCheckExistWrite(bookMarkName(2), ί�е�λ)
'    Call bookMarkNameCheckExistWrite(bookMarkName(3), ��ʼʱ��)
'    Call bookMarkNameCheckExistWrite(bookMarkName(4), ��������)
'    Call bookMarkNameCheckExistWrite(bookMarkName(5), ��˾�����)
'    Call bookMarkNameCheckExistWrite(bookMarkName(13), ���ű����)
    
   
    
    'д����һ�α༭״̬
    Call ReadLast
    
    
    '���̸ſ�
    
    'Call ModfAllCustomAttrib
    
    
    '��ʼ������������ȷ˳��
    Call formatCheckBoxInLine(Check2Box7, dic2, "����λ��", Page2 + "." + "����λ��")
    Call formatCheckBoxInLine(Check2Box5, dic2, "���赥λ", Page2 + "." + "���赥λ")
    Call formatCheckBoxInLine(Check2Box6, dic2, "�����ṹ", Page2 + "." + "�����ṹ")
    Call formatCheckBoxInLine(Check2Box8, dic2, "�������", Page2 + "." + "�������")
    Call formatCheckBoxInLine(Check2Box9, dic2, "¥����ϸ", Page2 + "." + "¥����ϸ")
    Call formatCheckBoxInLine(Check2Box10, dic2, "Ͷ����Ŀ", Page2 + "." + "Ͷ����Ŀ")
    Call formatCheckBoxInLine(Check2Box11, dic2, "����", Page2 + "." + "����")
    
    
    '���Ʒ�Χ

    Call formatCheckBoxInLine(Check3Box1, dic3, "����", Page3 + "." + "����")
    Call formatCheckBoxInLine(Check3Box2, dic3, "������", Page3 + "." + "������")
    Call formatCheckBoxInLine(Check3Box3, dic3, "����", Page3 + "." + "����")
    
    'Call AddBooleanFromCustomAttrib(Page3 + "." + "����234", False)

    
    '��������

    Call formatCheckBoxWrap(Check4Box1, dic4, "��1��", Page4 + "." + "�б��ļ�")
    Call formatCheckBoxWrap(Check4Box2, dic4, "��2��", Page4 + "." + "ͼֽ")
    Call formatCheckBoxWrap(Check4Box3, dic4, "��3��", Page4 + "." + "��ϵ��")
    Call formatCheckBoxWrap(Check4Box4, dic4, "��4��", Page4 + "." + "��������")
    Call formatCheckBoxWrap(Check4Box5, dic4, "��5��", Page4 + "." + "��������")



    '���Ʒ���


    '���ϵ��۷�

    Call formatCheckBoxWrap(Check5Box11, dic51, "��1��", Page5 + "." + "���ϵ��۷�.��������")
    Call formatCheckBoxWrap(Check5Box12, dic51, "��2��", Page5 + "." + "���ϵ��۷�.�۸���Դ")
    Call formatCheckBoxWrap(Check5Box13, dic51, "��3��", Page5 + "." + "���ϵ��۷�.�˹�����")
    Call formatCheckBoxWrap(Check5Box14, dic51, "��4��", Page5 + "." + "���ϵ��۷�.����")



    '�ۺϵ��۷�

    Call formatCheckBoxWrap(Check5Box21, dic52, "��1��", Page5 + "." + "�ۺϵ��۷�.��������")
    Call formatCheckBoxWrap(Check5Box22, dic52, "��2��", Page5 + "." + "�ۺϵ��۷�.�۸���Դ")
    Call formatCheckBoxWrap(Check5Box23, dic52, "��3��", Page5 + "." + "�ۺϵ��۷�.�˹�����")
    Call formatCheckBoxWrap(Check5Box24, dic52, "��4��", Page5 + "." + "�ۺϵ��۷�.����")



    'ȫ�����ۺϵ��۷�

    Call formatCheckBoxWrap(Check5Box31, dic53, "��1��", Page5 + "." + "ȫ�����ۺϵ��۷�.��������")
    Call formatCheckBoxWrap(Check5Box32, dic53, "��2��", Page5 + "." + "ȫ�����ۺϵ��۷�.�۸���Դ")
    Call formatCheckBoxWrap(Check5Box33, dic53, "��3��", Page5 + "." + "ȫ�����ۺϵ��۷�.�˹�����")
    Call formatCheckBoxWrap(Check5Box34, dic53, "��4��", Page5 + "." + "ȫ�����ۺϵ��۷�.����")


    '����˵��

    Call formatCheckBoxWrap(Check7Box1, dic7, "��1��", Page7 + "." + "ˮ��")
    Call formatCheckBoxWrap(Check7Box2, dic7, "��2��", Page7 + "." + "���ڼ�����")
    Call formatCheckBoxWrap(Check7Box3, dic7, "��3��", Page7 + "." + "���⴦��")
    Call formatCheckBoxWrap(Check7Box4, dic7, "��4��", Page7 + "." + "Ԥ�������")


    '����

    Call formatCheckBoxInLine(Check8Box1, dic8, "��1��", Page8 + "." + "Ԥ����ϸ��")
    Call formatCheckBoxInLine(Check8Box2, dic8, "��2��", Page8 + "." + "Ԥ�����Ҫ��")
    Call formatCheckBoxInLine(Check8Box3, dic8, "��3��", Page8 + "." + "��ϵ��")

    '������Ϣ
    
    Call readPage1(��Ŀ����, Page1 + "." + "��Ŀ����")
    Call readPage1(ί�е�λ, Page1 + "." + "ί�е�λ")
    Call readPage1(��ʼʱ��, Page1 + "." + "��ʼʱ��")
    Call readPage1(��������, Page1 + "." + "��������")
    Call readPage1(��˾�����, Page1 + "." + "��˾�����")
    Call readPage1(���ű����, Page1 + "." + "���ű����")
    
    Call bookMarkNameCheckExistWrite(bookMarkName(10), TBRptCkl)
    
    
    
    Call UpComboBox(ComboBox1)
    Call UpComboBox(ComboBox2, True)
    
    Call UpListBox(ListBox1)
    
End Sub



Private Sub MultiPage1_Change()
    
    
    Select Case MultiPage1.value
        Case 0
             Egg = Egg + "1"
        Case 1
            'Call CheckBoxClickNew(Check2Box5, TxtBasicInfo, dic2, "���赥λ", Page2 + "." + "���赥λ", TextValue2, MultiPage1, 1)
            'Call CheckBoxClickNew(Check2Box7, TxtBasicInfo, dic2, "����λ��", Page2 + "." + "����λ��", TextValue2, MultiPage1, 1)
            Egg = Egg + "2"
        Case 2
            'Call CheckBoxClickNew(Check3Box1, TBRptRg, dic3, "����", Page3 + "." + "����", TextValue3, MultiPage1, 2)

            Egg = Egg + "3"

        Case 3
            'Call CheckBoxClickNum(Check4Box1, TBRptAcc, dic4, "��1��", Page4 + "." + "�б��ļ�", TextValue4, MultiPage1, 3)
            Egg = Egg + "4"

        Case 4
             Egg = Egg + "5"

        Case 5
             Egg = Egg + "6"

        Case 6
             Egg = Egg + "7"

        Case 7
             'Call CheckBoxClickNum(Check8Box1, AuditorText, dic8, "��1��", Page8 + "." + "Ԥ����ϸ��", TextValue8, MultiPage1, 7, True)

             Egg = Egg + "8"

    End Select

            If oRegEgg.test(Egg) = True Then

                MultiPage1.Pages.Item(8).Visible = True


            End If
End Sub

Public Function ReadLast()
        
      For Each prop In ActiveDocument.CustomDocumentProperties
                       
        If prop.Type = msoPropertyTypeBoolean Then
            If prop.value = True Then
                Controls(prop.Name).value = True
            End If
        Else
        End If
    Next
    
End Function

Private Sub Check3Box4_Click()
    
    Call CheckBoxClickNew(Check3Box4, TBRptRg, dic3, "����23", Page3 + "." + "����23", TextValue3, MultiPage1, 2)
    
End Sub


Private Sub ComboBox1_Change()
    If ComboBox1 = Page5 Then
        ComboBox2.Visible = True
    Else
        ComboBox2.Visible = False
    End If
End Sub

Private Sub CommandButton51_Click()
     Call bookMarkNameReadFromDic(dic51, TBRptMtd, True)
     Call CheckedAllBox(MultiPage2, 0)
End Sub

Private Sub CommandButton90_Click()
     Call bookMarkNameReadFromDic(dic52, TBRptMtd, True)
     Call CheckedAllBox(MultiPage2, 1)
End Sub

Private Sub CommandButton91_Click()
     Call bookMarkNameReadFromDic(dic53, TBRptMtd, True)
     Call CheckedAllBox(MultiPage2, 2)
End Sub

Private Sub CommandButton94_Click()
   
    Call DelStringFromCustomAttrib(ListBox1.Text)
    
    ListBox1.Clear
    
    Call UpListBox(ListBox1)
    
End Sub

Private Sub CommandButton92_Click()
    Dim Combo1Str As String, Combo2Str, TextBox1Str, TextBox2Str, Name
    Dim Comb As Control
 
    TextBox1Str = TextBox35.Text
    TextBox2Str = TextBox36.Text
    
    If ComboBox1 = Page5 Then
        Name = ComboBox1 + "." + ComboBox2 + "." + TextBox1Str
    Else
        Name = ComboBox1 + "." + TextBox1Str
    End If

    If TextBox2Str <> "" Then
    
        For Each prop In ActiveDocument.CustomDocumentProperties
                           
            If prop.Name = Name Then
                Label34.Caption = "������ͬ�Զ����������ƣ��뻻��������"
                Label34.ForeColor = &HFF&
                MsgBox "������ͬ�Զ����������ƣ��뻻��������"
                
                Exit For
            Else
                Label34.ForeColor = &H80000012
                Call AddStringFromCustomAttrib(CStr(Name), CStr(TextBox2Str))
                'MsgBox "����Զ������� " + Name + "�ɹ�"
                
                Call UpListBox(ListBox1)
            End If
        Next
    Else
        MsgBox "����д�Զ�����������"
    End If
    
    
End Sub

Private Sub CommandButton93_Click()
    Dim Combo1Str As String, Combo2Str, TextBox1Str, TextBox2Str, TempName
    Dim Comb As Control
    Dim Name As String
    
    TextBox1Str = TextBox35.Text
    
    If ComboBox1 = Page5 Then
        Name = ComboBox1 + "." + ComboBox2 + "." + TextBox1Str
    Else
        Name = ComboBox1 + "." + TextBox1Str
    End If

    If TextBox1Str <> "" Then
        'Debug.Print Join(dic.Items(), "")
        
'        For Each objProperty In ActiveDocument.CustomDocumentProperties
'            Debug.Print objProperty.Name, objProperty.Type, objProperty.value
'        Next
        For Each prop In ActiveDocument.CustomDocumentProperties
                           
            If prop.Name = Name Then
                Label34.Caption = "������ͬ�Զ����������ƣ��뻻��������"
                Label34.ForeColor = &HFF&
                Exit For
            Else
                Label34.Caption = "���ƿ���"
            End If
        Next
    Else
        MsgBox "����д�Զ�����������"
    End If

    
End Sub



'Public Function TextValue2_Change(showHideText, dicCont)
'    Debug.Print showHideText
'    Debug.Print dicCont
'
'End Function


Private Sub ��Ŀ����_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    Call SyncEditSingleText(��Ŀ����, FrmBgRpt, dic2, dic3, dic4, dic8)
    
End Sub

Private Sub ί�е�λ_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    Call SyncEditSingleText(ί�е�λ, FrmBgRpt, dic2)
    
End Sub

Private Sub ��ʼʱ��_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    Call SyncEditSingleText(��ʼʱ��)
    
End Sub

Private Sub ��������_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    Call SyncEditSingleText(��������)
    
End Sub

Private Sub ��˾�����_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    Call SyncEditSingleText(��˾�����)
    
End Sub

Private Sub ���ű����_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    Call SyncEditSingleText(���ű����)
    
End Sub

Private Sub TextValue2_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    Call SyncEditTextLine(TextValue2, dic2, TxtBasicInfo)
    
End Sub

Private Sub TextValue3_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    Call SyncEditTextLine(TextValue3, dic3, TBRptRg)
    
End Sub


Private Sub TextValue4_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    Call SyncEditTextWrap(TextValue4, dic4, TBRptAcc)
    
End Sub



Private Sub TextValue7_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    Call SyncEditTextWrap(TextValue7, dic7, TBRptOth)
    
End Sub

Private Sub TextValue8_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    Call SyncEditTextWrap(TextValue8, dic8, AuditorText, True)
    
End Sub

Private Sub TextValue51_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    Call SyncEditTextWrap(TextValue51, dic51, TBRptMtd)
    
End Sub

Private Sub TextValue52_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    Call SyncEditTextWrap(TextValue52, dic52, TBRptMtd)
    
End Sub

Private Sub TextValue53_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    Call SyncEditTextWrap(TextValue53, dic53, TBRptMtd)
    
End Sub



Private Sub Check3Box1_Click()
    
     'Call CheckBoxClick(Check3Box1, TBRptRg, dic3, TextBox31, "����")
     Call CheckBoxClickNew(Check3Box1, TBRptRg, dic3, "����", Page3 + "." + "����", TextValue3, MultiPage1, 2)
    
End Sub

Private Sub Check3Box2_Click()
    'Call CheckBoxClick(Check3Box2, TBRptRg, dic3, TextBox32, "������")
    Call CheckBoxClickNew(Check3Box2, TBRptRg, dic3, "������", Page3 + "." + "������", TextValue3, MultiPage1, 2)
End Sub

Private Sub Check3Box3_Click()
    'Call CheckBoxClick(Check3Box3, TBRptRg, dic3, TextBox33, "����")
    Call CheckBoxClickNew(Check3Box3, TBRptRg, dic3, "����", Page3 + "." + "����", TextValue3, MultiPage1, 2)
End Sub

Private Sub Check8Box1_Click()
    'Call CheckBoxClick(Check8Box1, AuditorText, dic8, Text8Box1, "��1��")
    Call CheckBoxClickNum(Check8Box1, AuditorText, dic8, "��1��", Page8 + "." + "Ԥ����ϸ��", TextValue8, MultiPage1, 7, True)
End Sub

Private Sub Check8Box2_Click()
    'Call CheckBoxClick(Check8Box2, AuditorText, dic8, Text8Box2, "��2��")
    Call CheckBoxClickNum(Check8Box2, AuditorText, dic8, "��2��", Page8 + "." + "Ԥ�����Ҫ��", TextValue8, MultiPage1, 7, True)
End Sub

Private Sub Check8Box3_Click()
    'Call CheckBoxClick(Check8Box3, AuditorText, dic8, Text8Box3, "��3��")
    Call CheckBoxClickNum(Check8Box3, AuditorText, dic8, "��3��", Page8 + "." + "��ϵ��", TextValue8, MultiPage1, 7, True)
End Sub


Private Sub CheckBox1_Click()
    If CheckBox1.value = False Then
       Fit = False
    End If
    If CheckBox1.value = True Then
        Fit = True
    End If

End Sub

'���۸�
Private Sub CommandButton54_Click()
    
    If IsNumeric(TextBox9.Text) = True Then
        TextBox34.value = CChinese(TextBox9.Text)
    Else
        MsgBox "�����봿����"
    End If
    
End Sub


'���ɱ��ƽ������
Private Sub CommandButton61_Click()
    Dim MyStr As String, MyStr2, MyStr3
    MyStr2 = TextBox9.Text
    TBRptCkl.Text = Chr(32) + Chr(32) + Chr(32) + Chr(32) + "������Ԥ����Ƽ� " + MyStr2 + "Ԫ��"

End Sub


Private Sub CommandButton83_Click() '���̸ſ�
    
    'Call bookMarkNameCheckExistWrite(bookMarkName(6), TxtBasicInfo)
    
     Call bookMarkNameReadFromDic(dic2, TxtBasicInfo)
     Call CheckedAllBox(MultiPage1, 1)
     
End Sub



Private Sub CommandButton84_Click()
    'Call bookMarkNameCheckExistWrite(bookMarkName(8), TBRptAcc)
     Call bookMarkNameReadFromDic(dic4, TBRptAcc, True)
     Call CheckedAllBox(MultiPage1, 3)

End Sub

Private Sub CommandButton85_Click()
    Call bookMarkNameCheckExistWrite(bookMarkName(9), TBRptMtd)
End Sub


Private Sub CommandButton86_Click()
'    Call bookMarkNameCheckExistWrite(bookMarkName(11), TBRptOth)
     Call bookMarkNameReadFromDic(dic7, TBRptOth, True)
     Call CheckedAllBox(MultiPage1, 6)

End Sub

Private Sub CommandButton87_Click()

'    Call bookMarkNameCheckExistWrite(bookMarkName(12), AuditorText)
    Call bookMarkNameReadFromDic(dic8, AuditorText)
    Call CheckedAllBox(MultiPage1, 7)

End Sub


Private Sub CommandButton88_Click()
    'Call bookMarkNameCheckExistWrite(bookMarkName(7), TBRptRg)
    
     Call bookMarkNameReadFromDic(dic3, TBRptRg)
     Call CheckedAllBox(MultiPage1, 2)
End Sub

Private Sub CommandButton89_Click()

    Call bookMarkNameCheckExistWrite(bookMarkName(1), ��Ŀ����)
    Call bookMarkNameCheckExistWrite(bookMarkName(2), ί�е�λ)
    Call bookMarkNameCheckExistWrite(bookMarkName(3), ��ʼʱ��)
    Call bookMarkNameCheckExistWrite(bookMarkName(4), ��������)
    Call bookMarkNameCheckExistWrite(bookMarkName(5), ��˾�����)
    Call bookMarkNameCheckExistWrite(bookMarkName(13), ���ű����)
    
End Sub


'����˵�� checkbox--begin
Private Sub Check7Box1_Click()
'    Call CheckBoxClick(Check7Box1, TBRptOth, dic7, Text7Box1, "��1��")
    Call CheckBoxClickNum(Check7Box1, TBRptOth, dic7, "��1��", Page7 + "." + "ˮ��", TextValue7, MultiPage1, 6)
End Sub


Private Sub Check7Box2_Click()
'    Call CheckBoxClick(Check7Box2, TBRptOth, dic7, Text7Box2, "��2��")
     Call CheckBoxClickNum(Check7Box2, TBRptOth, dic7, "��2��", Page7 + "." + "���ڼ�����", TextValue7, MultiPage1, 6)
End Sub


Private Sub Check7Box3_Click()
    'Call CheckBoxClick(Check7Box3, TBRptOth, dic7, Text7Box3, "��3��")
     Call CheckBoxClickNum(Check7Box3, TBRptOth, dic7, "��3��", Page7 + "." + "���⴦��", TextValue7, MultiPage1, 6)
End Sub

Private Sub Check7Box4_Click()
'    Call CheckBoxClick(Check7Box4, TBRptOth, dic7, Text7Box4, "��4��")
    
    Call CheckBoxClickNum(Check7Box4, TBRptOth, dic7, "��4��", Page7 + "." + "Ԥ�������", TextValue7, MultiPage1, 6)

End Sub


Private Sub Check7Box5_Click()
    Call CheckBoxClickNum(Check7Box5, TBRptOth, dic7, "��5��", Page7 + "." + "��չ3", TextValue7, MultiPage1, 6)
End Sub

'����˵��checkbox--end


'���Ʒ������ϵ��۷�checkbox--begin
Private Sub Check5Box11_Click()
'    Call CheckBoxClick(Check5Box11, TBRptMtd, dic51, Text5Box11, "��1��")
    Call CheckBoxClickNum(Check5Box11, TBRptMtd, dic51, "��1��", Page5 + "." + "���ϵ��۷�.��������", TextValue51, MultiPage2, 0)
End Sub

Private Sub check5Box12_Click()
'    Call CheckBoxClick(Check5Box12, TBRptMtd, dic51, Text5Box12, "��2��")
    Call CheckBoxClickNum(Check5Box12, TBRptMtd, dic51, "��2��", Page5 + "." + "���ϵ��۷�.�۸���Դ", TextValue51, MultiPage2, 0)
End Sub

Private Sub check5Box13_Click()
'   Call CheckBoxClick(Check5Box13, TBRptMtd, dic51, Text5Box13, "��3��")
    Call CheckBoxClickNum(Check5Box13, TBRptMtd, dic51, "��3��", Page5 + "." + "���ϵ��۷�.�˹�����", TextValue51, MultiPage2, 0)
End Sub

Private Sub check5Box14_Click()
'   Call CheckBoxClick(Check5Box14, TBRptMtd, dic51, Text5Box14, "��4��")
    Call CheckBoxClickNum(Check5Box14, TBRptMtd, dic51, "��4��", Page5 + "." + "���ϵ��۷�.����", TextValue51, MultiPage2, 0)
End Sub


'���Ʒ������ϵ��۷�checkbox--end


'���Ʒ����ۺϵ��۷�checkbox--begin
Private Sub Check5Box21_Click()
    'Call CheckBoxClick(Check5Box21, TBRptMtd, dic52, Text5Box21, "��1��")
    Call CheckBoxClickNum(Check5Box21, TBRptMtd, dic52, "��1��", Page5 + "." + "�ۺϵ��۷�.��������", TextValue52, MultiPage2, 1)
End Sub


Private Sub check5Box22_Click()
   'Call CheckBoxClick(Check5Box22, TBRptMtd, dic52, Text5Box22, "��2��")
    Call CheckBoxClickNum(Check5Box22, TBRptMtd, dic52, "��2��", Page5 + "." + "�ۺϵ��۷�.�۸���Դ", TextValue52, MultiPage2, 1)
End Sub


Private Sub check5Box23_Click()
    'Call CheckBoxClick(Check5Box23, TBRptMtd, dic52, Text5Box23, "��3��")
    Call CheckBoxClickNum(Check5Box23, TBRptMtd, dic52, "��3��", Page5 + "." + "�ۺϵ��۷�.�˹�����", TextValue52, MultiPage2, 1)
End Sub

Private Sub check5Box24_Click()
    'Call CheckBoxClick(Check5Box24, TBRptMtd, dic52, Text5Box24, "��4��")
    Call CheckBoxClickNum(Check5Box24, TBRptMtd, dic52, "��4��", Page5 + "." + "�ۺϵ��۷�.����", TextValue52, MultiPage2, 1)
End Sub

'���Ʒ����ۺϵ��۷�checkbox--end


'���Ʒ���ȫ�����ۺϵ��۷�checkbox--begin
Private Sub Check5Box31_Click()
    'Call CheckBoxClick(Check5Box31, TBRptMtd, dic53, Text5Box31, "��1��")
    Call CheckBoxClickNum(Check5Box31, TBRptMtd, dic53, "��1��", Page5 + "." + "ȫ�����ۺϵ��۷�.��������", TextValue53, MultiPage2, 2)
End Sub


Private Sub check5Box32_Click()
    'Call CheckBoxClick(Check5Box32, TBRptMtd, dic53, Text5Box32, "��2��")
    Call CheckBoxClickNum(Check5Box32, TBRptMtd, dic53, "��2��", Page5 + "." + "ȫ�����ۺϵ��۷�.�۸���Դ", TextValue53, MultiPage2, 2)
End Sub


Private Sub check5Box33_Click()
    'Call CheckBoxClick(Check5Box33, TBRptMtd, dic53, Text5Box33, "��3��")
    Call CheckBoxClickNum(Check5Box33, TBRptMtd, dic53, "��3��", Page5 + "." + "ȫ�����ۺϵ��۷�.�˹�����", TextValue53, MultiPage2, 2)
End Sub

Private Sub check5Box34_Click()
    'Call CheckBoxClick(Check5Box34, TBRptMtd, dic53, Text5Box34, "��4��")
    Call CheckBoxClickNum(Check5Box34, TBRptMtd, dic53, "��4��", Page5 + "." + "ȫ�����ۺϵ��۷�.����", TextValue53, MultiPage2, 2)
End Sub

'���Ʒ���ȫ�����ۺϵ��۷�checkbox--end



'��������checkbox--begin
Private Sub Check4Box1_Click()
    'Call CheckBoxClick(Check4Box1, TBRptAcc, dic4, ��1��, "��1��")
    
    Call CheckBoxClickNum(Check4Box1, TBRptAcc, dic4, "��1��", Page4 + "." + "�б��ļ�", TextValue4, MultiPage1, 3)
End Sub

Private Sub Check4Box1_MouseOver()
    
End Sub

Private Sub Check4Box2_Click()
    'Call CheckBoxClick(Check4Box2, TBRptAcc, dic4, ��2��, "��2��")
    
    Call CheckBoxClickNum(Check4Box2, TBRptAcc, dic4, "��2��", Page4 + "." + "ͼֽ", TextValue4, MultiPage1, 3)
End Sub

Private Sub Check4Box3_Click()
    'Call CheckBoxClick(Check4Box3, TBRptAcc, dic4, ��3��, "��3��")
    
    Call CheckBoxClickNum(Check4Box3, TBRptAcc, dic4, "��3��", Page4 + "." + "��ϵ��", TextValue4, MultiPage1, 3)
End Sub

Private Sub Check4Box4_Click()
    'Call CheckBoxClick(Check4Box4, TBRptAcc, dic4, ��4��, "��4��")
    
    Call CheckBoxClickNum(Check4Box4, TBRptAcc, dic4, "��4��", Page4 + "." + "��������", TextValue4, MultiPage1, 3)
End Sub

Private Sub Check4Box5_Click()
    'Call CheckBoxClick(Check4Box5, TBRptAcc, dic4, ��5��, "��5��")

    Call CheckBoxClickNum(Check4Box5, TBRptAcc, dic4, "��5��", Page4 + "." + "��������", TextValue4, MultiPage1, 3)
End Sub

'��������checkbox---end


'���̸ſ�checkbox---begin
Private Sub Check2Box5_Click()
    Call CheckBoxClickNew(Check2Box5, TxtBasicInfo, dic2, "���赥λ", Page2 + "." + "���赥λ", TextValue2, MultiPage1, 1)
End Sub

Private Sub Check2Box6_Click()
    Call CheckBoxClickNew(Check2Box6, TxtBasicInfo, dic2, "�����ṹ", Page2 + "." + "�����ṹ", TextValue2, MultiPage1, 1)
End Sub

Private Sub Check2Box7_Click() '����λ��

    'MultiPage1.Pages.Item(2).Visible = False
    
    
    Call CheckBoxClickNew(Check2Box7, TxtBasicInfo, dic2, "����λ��", Page2 + "." + "����λ��", TextValue2, MultiPage1, 1)
    
    'UserForm1.Show 0

End Sub


Private Sub Check2Box8_Click()
    Call CheckBoxClickNew(Check2Box8, TxtBasicInfo, dic2, "�������", Page2 + "." + "�������", TextValue2, MultiPage1, 1)
End Sub

Private Sub Check2Box10_Click()
    Call CheckBoxClickNew(Check2Box10, TxtBasicInfo, dic2, "Ͷ����Ŀ", Page2 + "." + "Ͷ����Ŀ", TextValue2, MultiPage1, 1)
End Sub

Private Sub Check2Box11_Click()
    Call CheckBoxClickNew(Check2Box11, TxtBasicInfo, dic2, "����", Page2 + "." + "����", TextValue2, MultiPage1, 1)
End Sub

Private Sub Check2Box9_Click()
    Call CheckBoxClickNew(Check2Box9, TxtBasicInfo, dic2, "¥����ϸ", Page2 + "." + "¥����ϸ", TextValue2, MultiPage1, 1)
End Sub

'���̸ſ�checkbox---end

Private Sub CmbClup_Click()
    TBcash.Text = TBRptAks.Text
    TBRptAks.Text = ""
    CmbRst.Enabled = True
    CmbClup.Enabled = False
End Sub

Private Sub CmbRst_Click()
    TBRptAks.Text = TBcash.Text
    CmbRst.Enabled = False
    CmbClup.Enabled = True
End Sub

Private Sub CmdChange_Click() '��ʼ����
    Dim bkMark As Range '����һ��bkMark��range�����һ��bookmarkName�ַ�������
    

    Call userWriteBookMarkName(bookMarkName(1), ��Ŀ����, True)
    Call userWriteBookMarkName(bookMarkName(2), ί�е�λ, True)
    Call userWriteBookMarkName(bookMarkName(3), ��ʼʱ��, True)
    Call userWriteBookMarkName(bookMarkName(4), ��������, True)
    Call userWriteBookMarkName(bookMarkName(5), �����, True)

    Call userWriteBookMarkName(bookMarkName(6), TxtBasicInfo, True)
    Call ReplaceTextwithCrossRef(bookMarkName(6), bookMarkName(1), TxtBasicInfo)
    Call ReplaceTextwithCrossRef(bookMarkName(6), bookMarkName(2), TxtBasicInfo)
    
    Call userWriteBookMarkName(bookMarkName(7), TBRptRg, True)
    Call ReplaceTextwithCrossRef(bookMarkName(7), bookMarkName(1), TBRptRg)
    
    Call userWriteBookMarkName(bookMarkName(8), TBRptAcc, True)
    Call ReplaceTextwithCrossRef(bookMarkName(8), bookMarkName(1), TBRptAcc)
    
    Call userWriteBookMarkName(bookMarkName(9), TBRptMtd, True)
    Call userWriteBookMarkName(bookMarkName(10), TBRptCkl, True)
    Call userWriteBookMarkName(bookMarkName(11), TBRptOth, True)
    
    Call userWriteBookMarkName(bookMarkName(12), AuditorText, True)
    Call ReplaceTextwithCrossRef(bookMarkName(12), bookMarkName(1), AuditorText)
      
    UpAF
    
    MsgBox "����д����ɣ�"
        
   
    
End Sub



Private Sub CmdExit_Click()
    FrmBgRpt.Hide
    End
End Sub


Private Sub CmdRead_Click() '��ȡ����

    Dim bkMark As Range '����һ��bkMark��range�����һ��bookmarkName�ַ�������

    DimBMN

    
    Call bookMarkNameCheckExistWrite(bookMarkName(1), ��Ŀ����)
    Call bookMarkNameCheckExistWrite(bookMarkName(2), ί�е�λ)
    Call bookMarkNameCheckExistWrite(bookMarkName(3), ��ʼʱ��)
    Call bookMarkNameCheckExistWrite(bookMarkName(4), ��������)
    Call bookMarkNameCheckExistWrite(bookMarkName(5), �����)
    
'    Call bookMarkNameCheckExistWrite(bookMarkName(6), TxtBasicInfo)
'    Call bookMarkNameCheckExistWrite(bookMarkName(7), TBRptRg)
'    Call bookMarkNameCheckExistWrite(bookMarkName(8), TBRptAcc)
'    Call bookMarkNameCheckExistWrite(bookMarkName(9), TBRptMtd)
'    Call bookMarkNameCheckExistWrite(bookMarkName(10), TBRptCkl)
'    Call bookMarkNameCheckExistWrite(bookMarkName(11), TBRptOth)
'    Call bookMarkNameCheckExistWrite(bookMarkName(12), AuditorText)

    Call bookMarkNameReadFromDic(dic, TxtBasicInfo)
    Call CheckedAllBox(MultiPage1, 1)

    Call bookMarkNameReadFromDic(dic3, TBRptRg)
    Call CheckedAllBox(MultiPage1, 2)

    Call bookMarkNameReadFromDic(dic4, TBRptAcc, True)
    Call CheckedAllBox(MultiPage1, 3)

     Call bookMarkNameReadFromDic(dic51, TBRptMtd, True)
     Call CheckedAllBox(MultiPage2, 0)
     
     Call bookMarkNameReadFromDic(dic7, TBRptOth, True)
     Call CheckedAllBox(MultiPage1, 6)

    Call bookMarkNameReadFromDic(dic8, AuditorText)
    Call CheckedAllBox(MultiPage1, 7)

    If ((DateValue(��������.Text) < DateValue(��ʼʱ��.Text))) Then
        MsgBox "���������������飡", vbOKOnly, "���ڹ�ϵ������ʾ"
        Exit Sub
    End If



End Sub


Private Sub Command2Button1_Click() 'д�빤�̸ſ�
    
    Call userWriteBookMarkName(bookMarkName(6), TxtBasicInfo)
    
    Call ReplaceTextwithCrossRef(bookMarkName(6), bookMarkName(1), TxtBasicInfo)
    
    Call ReplaceTextwithCrossRef(bookMarkName(6), bookMarkName(2), TxtBasicInfo)

End Sub


Private Sub Command4Button1_Click() 'д���������
    
     Call userWriteBookMarkName(bookMarkName(8), TBRptAcc)
    
     Call ReplaceTextwithCrossRef(bookMarkName(8), bookMarkName(1), TBRptAcc)
     
End Sub


Private Sub Command5Button1_Click() '���Ʒ���
    Call userWriteBookMarkName(bookMarkName(9), TBRptMtd)
End Sub

Private Sub Command1Button3_Click() '1-������Ϣд��


' Dim mColButtons As New Collection
' Dim ctl As MSForms.Control
' Dim myText As MSForms.Control
''
''Set ctl = Me.Controls.Add("Forms.CommandButton.1")
''
''  With ctl
''  .Caption = "XYZ"
''  .Name = "AButton"
''  End With
''
''
''  Set myText = Me.Controls.Add("Forms.TextBox.1")
''
''  With myText
''
''  .Visible = True
''
''  .Text = "���Ǽ��صĶ�̬�ؼ� "
''  .Name = "xxxcvdsfds"
''  .Width = 320
''
''  End With
    
    Call userWriteBookMarkName(bookMarkName(1), ��Ŀ����, True)
    Call userWriteBookMarkName(bookMarkName(2), ί�е�λ, True)
    Call userWriteBookMarkName(bookMarkName(3), ��ʼʱ��, True)
    Call userWriteBookMarkName(bookMarkName(4), ��������, True)
    Call userWriteBookMarkName(bookMarkName(5), ��˾�����, True)
    Call userWriteBookMarkName(bookMarkName(13), ���ű����, True)

    Dim My1Str1 As String, My1Str2, My1Str3, My1Str4, My1Str5

    My1Str1 = ��Ŀ����.Text
    My1Str5 = ί�е�λ.Text
    My1Str2 = ��ʼʱ��.Text
    My1Str3 = ��������.Text
    My1Str4 = ��˾�����.Text

    
    Call ModfStringFromCustomAttrib(Page2 + ".����λ��", My1Str1 + "λ��________________��")
    
    Call ModfStringFromCustomAttrib(Page2 + ".���赥λ", "�����̽��赥λΪ" + ί�е�λ.Text + ",")
    
    Call ModfStringFromCustomAttrib(Page3 + ".����", "���α��Ʒ�Χ��" + My1Str1 + "����׮�����̡��������̡���ƺ���̡���������ṹ���̵ȡ�")
    
    Call ModfStringFromCustomAttrib(Page4 + ".�б��ļ�", "��" + My1Str1 + "Ԥ�����Ҫ�󡷡��б��ļ���" & vbCr)
    
    Call ModfStringFromCustomAttrib(Page8 + ".Ԥ����ϸ��", "��" + My1Str1 + "Ԥ�������ϸ��" & vbCr)
    
    UpAF


End Sub


'д����ƽ����ǩ
Private Sub CommandButton62_Click()

    Call userWriteBookMarkName(bookMarkName(10), TBRptCkl)
    
End Sub



'д������˵����ǩ
Private Sub CommandButton72_Click()
    
    Call userWriteBookMarkName(bookMarkName(11), TBRptOth)
    
    
End Sub


'д�븽����ǩ
Private Sub CommandButton82_Click()
    
    Call userWriteBookMarkName(bookMarkName(12), AuditorText)
    
    Call ReplaceTextwithCrossRef(bookMarkName(12), bookMarkName(1), AuditorText)
    
End Sub


'д����Ʒ�Χ��ǩ
Private Sub CommandButton32_Click()
   
    Call userWriteBookMarkName(bookMarkName(7), TBRptRg)
    
    Call ReplaceTextwithCrossRef(bookMarkName(7), bookMarkName(1), TBRptRg)
    
End Sub

Private Sub TextBox9_Change()
    Call setFormat(TextBox9, TBRptCkl)
End Sub


Private Sub setFormat(objTxt As textBox, textCont As textBox)
    Dim v, strNew$, s1$, intSelStart%
    Dim reg As Object
    Set reg = CreateObject("vbscript.regExp")
    reg.Global = True
    If objTxt.Text = "" Then Exit Sub
    v = Split(objTxt.Text, ".")
    s1 = v(0)
    v(0) = Replace(v(0), ",", "")
    reg.Pattern = "(\d{3})"
    v(0) = StrReverse(reg.Replace(StrReverse(v(0)), "$1,")) 'ÿ��3���ּӶ���
    reg.Pattern = "^([^\d]*?),"
    v(0) = reg.Replace(v(0), "$1")
    intSelStart = objTxt.SelStart
    objTxt.Text = Join(v, ".")
    objTxt.SelStart = intSelStart + Len(v(0)) - Len(s1)
    
    textCont.Text = "������Ԥ����Ƽ� " & Join(v, ".")
    
End Sub


