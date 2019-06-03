Attribute VB_Name = "model"



Public Function ReadStringFromCustomAttrib(ByRef key As String) As String
    On Error GoTo MyErr
        ReadStringFromCustomAttrib = ActiveDocument.CustomDocumentProperties(key).value
    
    Exit Function
MyErr:

    Msg = " ��ȡ���ݳ���!�����Զ������� " & key & " �Ƿ����"

    MsgBox Msg
    
    
End Function



Public Function AddStringFromCustomAttrib(ByVal key As String, ByVal values As String)
    
    ActiveDocument.CustomDocumentProperties.add Name:=key, Type:=msoPropertyTypeString, LinkToContent:=False, value:=values
    
    'MsgBox "����Զ������� " + key + " �ɹ�"
    
End Function

Public Function ModfStringFromCustomAttrib(ByRef key As String, ByRef value As String)
    On Error GoTo MyErr
        ActiveDocument.CustomDocumentProperties(key).value = value
   Exit Function
MyErr:

    Msg = " �޸�����ʱ����!�����Զ������� " & key & " �Ƿ����"

    MsgBox Msg
End Function


Public Function DelStringFromCustomAttrib(ByVal key As String)

    ActiveDocument.CustomDocumentProperties(key).Delete
    
End Function



Public Function AddBooleanFromCustomAttrib(ByVal key As String, ByVal values As Boolean)
    
    ActiveDocument.CustomDocumentProperties.add Name:=key, Type:=msoPropertyTypeBoolean, LinkToContent:=False, value:=values
    
End Function


Public Function ModfBooleanFromCustomAttrib(ByRef key As String, ByRef values As Boolean)

    ActiveDocument.CustomDocumentProperties(key) = values

End Function

Public Function ModfAllCustomAttrib()
    
    Call ModfStringFromCustomAttrib(Page1 + ".��Ŀ����", "________��������")
    Call ModfStringFromCustomAttrib(Page1 + ".ί�е�λ", "________��˾")
    Call ModfStringFromCustomAttrib(Page1 + ".��˾�����", "��Ƽ���[2018]��14-395-E48")
    Call ModfStringFromCustomAttrib(Page1 + ".���ű����", "(045)��")
    Call ModfStringFromCustomAttrib(Page1 + ".��ʼʱ��", "2018��8��21��")
    Call ModfStringFromCustomAttrib(Page1 + ".��������", "2018��9��13��")
    
    
    Call ModfStringFromCustomAttrib(Page2 + ".����λ��", "________��������λ��________________��")
    Call ModfStringFromCustomAttrib(Page2 + ".���赥λ", "�����̽��赥λΪ________��˾,")
    Call ModfStringFromCustomAttrib(Page2 + ".�����ṹ", "�����ṹΪ________��")
    Call ModfStringFromCustomAttrib(Page2 + ".�������", "�ܽ������____ƽ���ף�")
    Call ModfStringFromCustomAttrib(Page2 + ".¥����ϸ", "����__��,����__��,��__�㡣")
    Call ModfStringFromCustomAttrib(Page2 + ".Ͷ����Ŀ", "��Ŀ��Ͷ�ʣ�______��Ԫ��")
    Call ModfStringFromCustomAttrib(Page2 + ".����", "������ע���ݡ�")
    
    Call ModfStringFromCustomAttrib(Page3 + ".����", "���α��Ʒ�Χ��________�������̣���׮�����̡��������̡���ƺ���̡���������ṹ���̵ȡ�")
    Call ModfStringFromCustomAttrib(Page3 + ".������", "�����ֽṹ���̡�ˮ�簲װ���̡�������װ�޹��̡�")
    Call ModfStringFromCustomAttrib(Page3 + ".����", "��������б��ļ���")

    Call ModfStringFromCustomAttrib(Page4 + ".�б��ļ�", "��________��������Ԥ�����Ҫ�󡷡��б��ļ���" & vbCr)
    Call ModfStringFromCustomAttrib(Page4 + ".ͼֽ", "ͼֽ�������ļ���Ϊ��_RDC�������ͼֽ��A��ͼ��" & vbCr)
    Call ModfStringFromCustomAttrib(Page4 + ".��ϵ��", "Ԥ����ƹ����еĹ�����ϵ��������������ϡ�" & vbCr)
    Call ModfStringFromCustomAttrib(Page4 + ".��������", "���㽭ʡ��������Ԥ�㶨���2010�棩�����㽭ʡ���蹤��ʩ�����ö����2010�棩�����㽭ʡʩ����е̨����ö����2010�棩�������蹤�̹������嵥�Ƽ۹淶����GB50500-2013�������㽭ʡ���蹤�̼Ƽ۹��򡷣�2010�棩 ����ز����ļ��ȡ�" & vbCr)
    Call ModfStringFromCustomAttrib(Page4 + ".��������", "�����й����ϡ�" & vbCr)
    
    Call ModfStringFromCustomAttrib(Page5 + ".���ϵ��۷�.��������", "���������ݹ���������淶�����ͼֽ���м��㡣" & vbCr)
    Call ModfStringFromCustomAttrib(Page5 + ".���ϵ��۷�.�۸���Դ", "�ۺϵ������ݹ������Ƽ۹淶���㽭ʡ������ع���Ԥ�㶨����мƼۣ����ϼ۰�������______���__����Ϣ�ۣ��޼۲��ϰ��ݹ��ۻ��г��ۼ�ȡ��" & vbCr)
    Call ModfStringFromCustomAttrib(Page5 + ".���ϵ��۷�.�˹�����", "�����˹����۸���____��һ��__Ԫ/���� ������__Ԫ/���ա�����__Ԫ/���ա�" & vbCr)
    Call ModfStringFromCustomAttrib(Page5 + ".���ϵ��۷�.����", "ʩ�����ð���_____ʡ���蹤��ʩ�����ö����2010�棩��ȡ����ȫ����ʩ���Ѱ�____�ƣ���ҵ����ѡ�����____��ȡ����Ѱ�______�ơ�ũ�񹤹��˱��հ�____�ơ�˰��____�ơ�" & vbCr)
    
    Call ModfStringFromCustomAttrib(Page5 + ".�ۺϵ��۷�.��������", "�ۺϵ������ݹ������Ƽ۹淶���㽭ʡ������ع���Ԥ�㶨����мƼۣ����ϼ۰�������______���__����Ϣ�ۣ��޼۲��ϰ��ݹ��ۻ��г��ۼ�ȡ��" & vbCr)
    Call ModfStringFromCustomAttrib(Page5 + ".�ۺϵ��۷�.�۸���Դ", "�����˹����۸���____��һ��__Ԫ/���� ������__Ԫ/���ա�����__Ԫ/���ա�" & vbCr)
    Call ModfStringFromCustomAttrib(Page5 + ".�ۺϵ��۷�.�˹�����", "ʩ�����ð���_____ʡ���蹤��ʩ�����ö����2010�棩��ȡ����ȫ����ʩ���Ѱ�____�ƣ���ҵ����ѡ�����____��ȡ����Ѱ�______�ơ�ũ�񹤹��˱��հ�____�ơ�˰��____�ơ�" & vbCr)
    Call ModfStringFromCustomAttrib(Page5 + ".�ۺϵ��۷�.����", "���������ݹ���������淶�����ͼֽ���м��㡣" & vbCr)
    
    Call ModfStringFromCustomAttrib(Page5 + ".ȫ�����ۺϵ��۷�.��������", "����������____ʡ������ع���Ԥ�㶨���й����������������ͼֽ���м��㡣" & vbCr)
    Call ModfStringFromCustomAttrib(Page5 + ".ȫ�����ۺϵ��۷�.�۸���Դ", "��������____ʡ������ع���Ԥ�㶨����мƼۣ����ϼ۰�������____���__����Ϣ�ۣ��޼۲��ϰ��г��ۼ�ȡ��" & vbCr)
    Call ModfStringFromCustomAttrib(Page5 + ".ȫ�����ۺϵ��۷�.�˹�����", "�����˹����۸���____��һ��__Ԫ/���� ������__Ԫ/���ա�����__Ԫ/���ա�" & vbCr)
    Call ModfStringFromCustomAttrib(Page5 + ".ȫ�����ۺϵ��۷�.����", "ʩ�����ð���______ʡ���蹤��ʩ�����ö����2010�棩��ȡ����ȫ����ʩ���Ѱ�__�ƣ���ҵ����ѡ�����__��ȡ����Ѱ�____�ơ�ũ�񹤹��˱��հ�____�ơ�˰��____�ơ�" & vbCr)
    
    Call ModfStringFromCustomAttrib(Page7 + ".ˮ��", "�������������ʱ��ˮ���õ�Ӳ����ɷ�����Э���ܳа����ṩ�����б������и���Ӳ����е�һ����ط��ã�ʩ����ˮ���õ�������б���֧�����ܳа��ˣ��˲��ַ����Ѱ�����Ͷ�걨���С�" & vbCr)
    Call ModfStringFromCustomAttrib(Page7 + ".���ڼ�����", "�����̼ƻ�����Ϊ____�졣������ǰ�������ڣ�һ�쿢��������������������۵����֮����ȡ����������ı���Ҫ�Եȣ����ܶ�ó�����������۵İٷ�֮����" & vbCr)
    Call ModfStringFromCustomAttrib(Page7 + ".���⴦��", "���Ԥ����Ƶ��е����⼰ͼֽ����֮���Ѱ����跽�Ļظ����������������������ϵ���ȡ�" & vbCr)
    Call ModfStringFromCustomAttrib(Page7 + ".Ԥ�������", "���ι���Ԥ����еȼ۲���Ϊ���̼ۿ�������ݡ�" & vbCr)

    Call ModfStringFromCustomAttrib(Page8 + ".Ԥ����ϸ��", "��________�������� Ԥ�������ϸ��" & vbCr)
    Call ModfStringFromCustomAttrib(Page8 + ".Ԥ�����Ҫ��", "��Ԥ�����Ҫ��" & vbCr)
    Call ModfStringFromCustomAttrib(Page8 + ".��ϵ��", "��������ϵ����" & vbCr)
    
End Function
