VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmBgRpt 
   Caption         =   "科佳预算报告生成工具"
   ClientHeight    =   6816
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   10176
   OleObjectBlob   =   "FrmBgRpt.frx":0000
   StartUpPosition =   1  '所有者中心
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
    
    bookMarkName(1) = "项目名称"   '为bookMarkName赋值
    bookMarkName(2) = "委托单位"
    bookMarkName(3) = "开始时间"
    bookMarkName(4) = "报告日期"
    bookMarkName(5) = "公司报告号"
    bookMarkName(6) = "工程概况"
    bookMarkName(7) = "编制范围"
    bookMarkName(8) = "编制依据"
    bookMarkName(9) = "编制方法"
    bookMarkName(10) = "编制结果"
    bookMarkName(11) = "其他说明"
    bookMarkName(12) = "附件"
    bookMarkName(13) = "部门报告号"
    
    
    Set oRegEgg = CreateObject("VBscript.RegExp")
    oRegEgg.Global = True
    oRegEgg.Pattern = "521314"
    
'    '预先写入编制依据，存在超出5条情况
'    Set bkMark = ActiveDocument.Bookmarks(bookMarkName(8)).Range  '查找书签“bookmark1”并赋值给bkMark
'    bkMark.Select '选中bkMark书签对应的文本
'    TBRptAcc.Text = bkMark.Text  '编制依据
        
    
End Sub



Private Sub CheckBox2_Click()
    Call CheckBoxClickNew(CheckBox2, TBRptRg, dic3, "其他24", Page3 + "." + "其他24", TextValue3, MultiPage1, 2)
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



'更新条目
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

    Dim bkMark As Range '定义一个bkMark的range对象和一个bookmarkName字符串对象


    Call userWriteBookMarkName(bookMarkName(1), 项目名称, True)
    Call userWriteBookMarkName(bookMarkName(2), 委托单位, True)
    Call userWriteBookMarkName(bookMarkName(3), 开始时间, True)
    Call userWriteBookMarkName(bookMarkName(4), 报告日期, True)
    Call userWriteBookMarkName(bookMarkName(5), 公司报告号, True)
    Call userWriteBookMarkName(bookMarkName(13), 部门报告号, True)

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


    '初始化本地已写入的数据，
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

    
    
    '基本信息
    Label29.FontSize = 6
    
    '打开配置面板 True 打开，False 关闭
    
    MultiPage1.Pages.Item(8).Visible = False
    
    '基本信息预写入'
    
'    Call bookMarkNameCheckExistWrite(bookMarkName(1), 项目名称)
'    Call bookMarkNameCheckExistWrite(bookMarkName(2), 委托单位)
'    Call bookMarkNameCheckExistWrite(bookMarkName(3), 开始时间)
'    Call bookMarkNameCheckExistWrite(bookMarkName(4), 报告日期)
'    Call bookMarkNameCheckExistWrite(bookMarkName(5), 公司报告号)
'    Call bookMarkNameCheckExistWrite(bookMarkName(13), 部门报告号)
    
   
    
    '写入上一次编辑状态
    Call ReadLast
    
    
    '工程概况
    
    'Call ModfAllCustomAttrib
    
    
    '初始化调用排列正确顺序
    Call formatCheckBoxInLine(Check2Box7, dic2, "工程位置", Page2 + "." + "工程位置")
    Call formatCheckBoxInLine(Check2Box5, dic2, "建设单位", Page2 + "." + "建设单位")
    Call formatCheckBoxInLine(Check2Box6, dic2, "建筑结构", Page2 + "." + "建筑结构")
    Call formatCheckBoxInLine(Check2Box8, dic2, "建筑面积", Page2 + "." + "建筑面积")
    Call formatCheckBoxInLine(Check2Box9, dic2, "楼层详细", Page2 + "." + "楼层详细")
    Call formatCheckBoxInLine(Check2Box10, dic2, "投资数目", Page2 + "." + "投资数目")
    Call formatCheckBoxInLine(Check2Box11, dic2, "其他", Page2 + "." + "其他")
    
    
    '编制范围

    Call formatCheckBoxInLine(Check3Box1, dic3, "包括", Page3 + "." + "包括")
    Call formatCheckBoxInLine(Check3Box2, dic3, "不包括", Page3 + "." + "不包括")
    Call formatCheckBoxInLine(Check3Box3, dic3, "其他", Page3 + "." + "其他")
    
    'Call AddBooleanFromCustomAttrib(Page3 + "." + "其他234", False)

    
    '编制依据

    Call formatCheckBoxWrap(Check4Box1, dic4, "第1条", Page4 + "." + "招标文件")
    Call formatCheckBoxWrap(Check4Box2, dic4, "第2条", Page4 + "." + "图纸")
    Call formatCheckBoxWrap(Check4Box3, dic4, "第3条", Page4 + "." + "联系函")
    Call formatCheckBoxWrap(Check4Box4, dic4, "第4条", Page4 + "." + "编制依据")
    Call formatCheckBoxWrap(Check4Box5, dic4, "第5条", Page4 + "." + "其他资料")



    '编制方法


    '工料单价法

    Call formatCheckBoxWrap(Check5Box11, dic51, "第1条", Page5 + "." + "工料单价法.计算依据")
    Call formatCheckBoxWrap(Check5Box12, dic51, "第2条", Page5 + "." + "工料单价法.价格来源")
    Call formatCheckBoxWrap(Check5Box13, dic51, "第3条", Page5 + "." + "工料单价法.人工单价")
    Call formatCheckBoxWrap(Check5Box14, dic51, "第4条", Page5 + "." + "工料单价法.费率")



    '综合单价法

    Call formatCheckBoxWrap(Check5Box21, dic52, "第1条", Page5 + "." + "综合单价法.计算依据")
    Call formatCheckBoxWrap(Check5Box22, dic52, "第2条", Page5 + "." + "综合单价法.价格来源")
    Call formatCheckBoxWrap(Check5Box23, dic52, "第3条", Page5 + "." + "综合单价法.人工单价")
    Call formatCheckBoxWrap(Check5Box24, dic52, "第4条", Page5 + "." + "综合单价法.费率")



    '全费用综合单价法

    Call formatCheckBoxWrap(Check5Box31, dic53, "第1条", Page5 + "." + "全费用综合单价法.计算依据")
    Call formatCheckBoxWrap(Check5Box32, dic53, "第2条", Page5 + "." + "全费用综合单价法.价格来源")
    Call formatCheckBoxWrap(Check5Box33, dic53, "第3条", Page5 + "." + "全费用综合单价法.人工单价")
    Call formatCheckBoxWrap(Check5Box34, dic53, "第4条", Page5 + "." + "全费用综合单价法.费率")


    '其他说明

    Call formatCheckBoxWrap(Check7Box1, dic7, "第1条", Page7 + "." + "水电")
    Call formatCheckBoxWrap(Check7Box2, dic7, "第2条", Page7 + "." + "工期及奖罚")
    Call formatCheckBoxWrap(Check7Box3, dic7, "第3条", Page7 + "." + "问题处理")
    Call formatCheckBoxWrap(Check7Box4, dic7, "第4条", Page7 + "." + "预算非依据")


    '附件

    Call formatCheckBoxInLine(Check8Box1, dic8, "第1条", Page8 + "." + "预算明细表")
    Call formatCheckBoxInLine(Check8Box2, dic8, "第2条", Page8 + "." + "预算编制要求")
    Call formatCheckBoxInLine(Check8Box3, dic8, "第3条", Page8 + "." + "联系函")

    '基本信息
    
    Call readPage1(项目名称, Page1 + "." + "项目名称")
    Call readPage1(委托单位, Page1 + "." + "委托单位")
    Call readPage1(开始时间, Page1 + "." + "开始时间")
    Call readPage1(报告日期, Page1 + "." + "报告日期")
    Call readPage1(公司报告号, Page1 + "." + "公司报告号")
    Call readPage1(部门报告号, Page1 + "." + "部门报告号")
    
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
            'Call CheckBoxClickNew(Check2Box5, TxtBasicInfo, dic2, "建设单位", Page2 + "." + "建设单位", TextValue2, MultiPage1, 1)
            'Call CheckBoxClickNew(Check2Box7, TxtBasicInfo, dic2, "工程位置", Page2 + "." + "工程位置", TextValue2, MultiPage1, 1)
            Egg = Egg + "2"
        Case 2
            'Call CheckBoxClickNew(Check3Box1, TBRptRg, dic3, "包括", Page3 + "." + "包括", TextValue3, MultiPage1, 2)

            Egg = Egg + "3"

        Case 3
            'Call CheckBoxClickNum(Check4Box1, TBRptAcc, dic4, "第1条", Page4 + "." + "招标文件", TextValue4, MultiPage1, 3)
            Egg = Egg + "4"

        Case 4
             Egg = Egg + "5"

        Case 5
             Egg = Egg + "6"

        Case 6
             Egg = Egg + "7"

        Case 7
             'Call CheckBoxClickNum(Check8Box1, AuditorText, dic8, "第1条", Page8 + "." + "预算明细表", TextValue8, MultiPage1, 7, True)

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
    
    Call CheckBoxClickNew(Check3Box4, TBRptRg, dic3, "其他23", Page3 + "." + "其他23", TextValue3, MultiPage1, 2)
    
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
                Label34.Caption = "存在相同自定义属性名称，请换其他名称"
                Label34.ForeColor = &HFF&
                MsgBox "存在相同自定义属性名称，请换其他名称"
                
                Exit For
            Else
                Label34.ForeColor = &H80000012
                Call AddStringFromCustomAttrib(CStr(Name), CStr(TextBox2Str))
                'MsgBox "添加自定义属性 " + Name + "成功"
                
                Call UpListBox(ListBox1)
            End If
        Next
    Else
        MsgBox "请填写自定义属性内容"
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
                Label34.Caption = "存在相同自定义属性名称，请换其他名称"
                Label34.ForeColor = &HFF&
                Exit For
            Else
                Label34.Caption = "名称可用"
            End If
        Next
    Else
        MsgBox "请填写自定义属性名称"
    End If

    
End Sub



'Public Function TextValue2_Change(showHideText, dicCont)
'    Debug.Print showHideText
'    Debug.Print dicCont
'
'End Function


Private Sub 项目名称_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    Call SyncEditSingleText(项目名称, FrmBgRpt, dic2, dic3, dic4, dic8)
    
End Sub

Private Sub 委托单位_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    Call SyncEditSingleText(委托单位, FrmBgRpt, dic2)
    
End Sub

Private Sub 开始时间_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    Call SyncEditSingleText(开始时间)
    
End Sub

Private Sub 报告日期_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    Call SyncEditSingleText(报告日期)
    
End Sub

Private Sub 公司报告号_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    Call SyncEditSingleText(公司报告号)
    
End Sub

Private Sub 部门报告号_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    Call SyncEditSingleText(部门报告号)
    
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
    
     'Call CheckBoxClick(Check3Box1, TBRptRg, dic3, TextBox31, "包括")
     Call CheckBoxClickNew(Check3Box1, TBRptRg, dic3, "包括", Page3 + "." + "包括", TextValue3, MultiPage1, 2)
    
End Sub

Private Sub Check3Box2_Click()
    'Call CheckBoxClick(Check3Box2, TBRptRg, dic3, TextBox32, "不包括")
    Call CheckBoxClickNew(Check3Box2, TBRptRg, dic3, "不包括", Page3 + "." + "不包括", TextValue3, MultiPage1, 2)
End Sub

Private Sub Check3Box3_Click()
    'Call CheckBoxClick(Check3Box3, TBRptRg, dic3, TextBox33, "其他")
    Call CheckBoxClickNew(Check3Box3, TBRptRg, dic3, "其他", Page3 + "." + "其他", TextValue3, MultiPage1, 2)
End Sub

Private Sub Check8Box1_Click()
    'Call CheckBoxClick(Check8Box1, AuditorText, dic8, Text8Box1, "第1条")
    Call CheckBoxClickNum(Check8Box1, AuditorText, dic8, "第1条", Page8 + "." + "预算明细表", TextValue8, MultiPage1, 7, True)
End Sub

Private Sub Check8Box2_Click()
    'Call CheckBoxClick(Check8Box2, AuditorText, dic8, Text8Box2, "第2条")
    Call CheckBoxClickNum(Check8Box2, AuditorText, dic8, "第2条", Page8 + "." + "预算编制要求", TextValue8, MultiPage1, 7, True)
End Sub

Private Sub Check8Box3_Click()
    'Call CheckBoxClick(Check8Box3, AuditorText, dic8, Text8Box3, "第3条")
    Call CheckBoxClickNum(Check8Box3, AuditorText, dic8, "第3条", Page8 + "." + "联系函", TextValue8, MultiPage1, 7, True)
End Sub


Private Sub CheckBox1_Click()
    If CheckBox1.value = False Then
       Fit = False
    End If
    If CheckBox1.value = True Then
        Fit = True
    End If

End Sub

'检查价格
Private Sub CommandButton54_Click()
    
    If IsNumeric(TextBox9.Text) = True Then
        TextBox34.value = CChinese(TextBox9.Text)
    Else
        MsgBox "请输入纯数字"
    End If
    
End Sub


'生成编制结果文字
Private Sub CommandButton61_Click()
    Dim MyStr As String, MyStr2, MyStr3
    MyStr2 = TextBox9.Text
    TBRptCkl.Text = Chr(32) + Chr(32) + Chr(32) + Chr(32) + "本工程预算编制价 " + MyStr2 + "元。"

End Sub


Private Sub CommandButton83_Click() '工程概况
    
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

    Call bookMarkNameCheckExistWrite(bookMarkName(1), 项目名称)
    Call bookMarkNameCheckExistWrite(bookMarkName(2), 委托单位)
    Call bookMarkNameCheckExistWrite(bookMarkName(3), 开始时间)
    Call bookMarkNameCheckExistWrite(bookMarkName(4), 报告日期)
    Call bookMarkNameCheckExistWrite(bookMarkName(5), 公司报告号)
    Call bookMarkNameCheckExistWrite(bookMarkName(13), 部门报告号)
    
End Sub


'其他说明 checkbox--begin
Private Sub Check7Box1_Click()
'    Call CheckBoxClick(Check7Box1, TBRptOth, dic7, Text7Box1, "第1条")
    Call CheckBoxClickNum(Check7Box1, TBRptOth, dic7, "第1条", Page7 + "." + "水电", TextValue7, MultiPage1, 6)
End Sub


Private Sub Check7Box2_Click()
'    Call CheckBoxClick(Check7Box2, TBRptOth, dic7, Text7Box2, "第2条")
     Call CheckBoxClickNum(Check7Box2, TBRptOth, dic7, "第2条", Page7 + "." + "工期及奖罚", TextValue7, MultiPage1, 6)
End Sub


Private Sub Check7Box3_Click()
    'Call CheckBoxClick(Check7Box3, TBRptOth, dic7, Text7Box3, "第3条")
     Call CheckBoxClickNum(Check7Box3, TBRptOth, dic7, "第3条", Page7 + "." + "问题处理", TextValue7, MultiPage1, 6)
End Sub

Private Sub Check7Box4_Click()
'    Call CheckBoxClick(Check7Box4, TBRptOth, dic7, Text7Box4, "第4条")
    
    Call CheckBoxClickNum(Check7Box4, TBRptOth, dic7, "第4条", Page7 + "." + "预算非依据", TextValue7, MultiPage1, 6)

End Sub


Private Sub Check7Box5_Click()
    Call CheckBoxClickNum(Check7Box5, TBRptOth, dic7, "第5条", Page7 + "." + "扩展3", TextValue7, MultiPage1, 6)
End Sub

'其他说明checkbox--end


'编制方法工料单价法checkbox--begin
Private Sub Check5Box11_Click()
'    Call CheckBoxClick(Check5Box11, TBRptMtd, dic51, Text5Box11, "第1条")
    Call CheckBoxClickNum(Check5Box11, TBRptMtd, dic51, "第1条", Page5 + "." + "工料单价法.计算依据", TextValue51, MultiPage2, 0)
End Sub

Private Sub check5Box12_Click()
'    Call CheckBoxClick(Check5Box12, TBRptMtd, dic51, Text5Box12, "第2条")
    Call CheckBoxClickNum(Check5Box12, TBRptMtd, dic51, "第2条", Page5 + "." + "工料单价法.价格来源", TextValue51, MultiPage2, 0)
End Sub

Private Sub check5Box13_Click()
'   Call CheckBoxClick(Check5Box13, TBRptMtd, dic51, Text5Box13, "第3条")
    Call CheckBoxClickNum(Check5Box13, TBRptMtd, dic51, "第3条", Page5 + "." + "工料单价法.人工单价", TextValue51, MultiPage2, 0)
End Sub

Private Sub check5Box14_Click()
'   Call CheckBoxClick(Check5Box14, TBRptMtd, dic51, Text5Box14, "第4条")
    Call CheckBoxClickNum(Check5Box14, TBRptMtd, dic51, "第4条", Page5 + "." + "工料单价法.费率", TextValue51, MultiPage2, 0)
End Sub


'编制方法工料单价法checkbox--end


'编制方法综合单价法checkbox--begin
Private Sub Check5Box21_Click()
    'Call CheckBoxClick(Check5Box21, TBRptMtd, dic52, Text5Box21, "第1条")
    Call CheckBoxClickNum(Check5Box21, TBRptMtd, dic52, "第1条", Page5 + "." + "综合单价法.计算依据", TextValue52, MultiPage2, 1)
End Sub


Private Sub check5Box22_Click()
   'Call CheckBoxClick(Check5Box22, TBRptMtd, dic52, Text5Box22, "第2条")
    Call CheckBoxClickNum(Check5Box22, TBRptMtd, dic52, "第2条", Page5 + "." + "综合单价法.价格来源", TextValue52, MultiPage2, 1)
End Sub


Private Sub check5Box23_Click()
    'Call CheckBoxClick(Check5Box23, TBRptMtd, dic52, Text5Box23, "第3条")
    Call CheckBoxClickNum(Check5Box23, TBRptMtd, dic52, "第3条", Page5 + "." + "综合单价法.人工单价", TextValue52, MultiPage2, 1)
End Sub

Private Sub check5Box24_Click()
    'Call CheckBoxClick(Check5Box24, TBRptMtd, dic52, Text5Box24, "第4条")
    Call CheckBoxClickNum(Check5Box24, TBRptMtd, dic52, "第4条", Page5 + "." + "综合单价法.费率", TextValue52, MultiPage2, 1)
End Sub

'编制方法综合单价法checkbox--end


'编制方法全费用综合单价法checkbox--begin
Private Sub Check5Box31_Click()
    'Call CheckBoxClick(Check5Box31, TBRptMtd, dic53, Text5Box31, "第1条")
    Call CheckBoxClickNum(Check5Box31, TBRptMtd, dic53, "第1条", Page5 + "." + "全费用综合单价法.计算依据", TextValue53, MultiPage2, 2)
End Sub


Private Sub check5Box32_Click()
    'Call CheckBoxClick(Check5Box32, TBRptMtd, dic53, Text5Box32, "第2条")
    Call CheckBoxClickNum(Check5Box32, TBRptMtd, dic53, "第2条", Page5 + "." + "全费用综合单价法.价格来源", TextValue53, MultiPage2, 2)
End Sub


Private Sub check5Box33_Click()
    'Call CheckBoxClick(Check5Box33, TBRptMtd, dic53, Text5Box33, "第3条")
    Call CheckBoxClickNum(Check5Box33, TBRptMtd, dic53, "第3条", Page5 + "." + "全费用综合单价法.人工单价", TextValue53, MultiPage2, 2)
End Sub

Private Sub check5Box34_Click()
    'Call CheckBoxClick(Check5Box34, TBRptMtd, dic53, Text5Box34, "第4条")
    Call CheckBoxClickNum(Check5Box34, TBRptMtd, dic53, "第4条", Page5 + "." + "全费用综合单价法.费率", TextValue53, MultiPage2, 2)
End Sub

'编制方法全费用综合单价法checkbox--end



'编制依据checkbox--begin
Private Sub Check4Box1_Click()
    'Call CheckBoxClick(Check4Box1, TBRptAcc, dic4, 第1条, "第1条")
    
    Call CheckBoxClickNum(Check4Box1, TBRptAcc, dic4, "第1条", Page4 + "." + "招标文件", TextValue4, MultiPage1, 3)
End Sub

Private Sub Check4Box1_MouseOver()
    
End Sub

Private Sub Check4Box2_Click()
    'Call CheckBoxClick(Check4Box2, TBRptAcc, dic4, 第2条, "第2条")
    
    Call CheckBoxClickNum(Check4Box2, TBRptAcc, dic4, "第2条", Page4 + "." + "图纸", TextValue4, MultiPage1, 3)
End Sub

Private Sub Check4Box3_Click()
    'Call CheckBoxClick(Check4Box3, TBRptAcc, dic4, 第3条, "第3条")
    
    Call CheckBoxClickNum(Check4Box3, TBRptAcc, dic4, "第3条", Page4 + "." + "联系函", TextValue4, MultiPage1, 3)
End Sub

Private Sub Check4Box4_Click()
    'Call CheckBoxClick(Check4Box4, TBRptAcc, dic4, 第4条, "第4条")
    
    Call CheckBoxClickNum(Check4Box4, TBRptAcc, dic4, "第4条", Page4 + "." + "编制依据", TextValue4, MultiPage1, 3)
End Sub

Private Sub Check4Box5_Click()
    'Call CheckBoxClick(Check4Box5, TBRptAcc, dic4, 第5条, "第5条")

    Call CheckBoxClickNum(Check4Box5, TBRptAcc, dic4, "第5条", Page4 + "." + "其他资料", TextValue4, MultiPage1, 3)
End Sub

'编制依据checkbox---end


'工程概况checkbox---begin
Private Sub Check2Box5_Click()
    Call CheckBoxClickNew(Check2Box5, TxtBasicInfo, dic2, "建设单位", Page2 + "." + "建设单位", TextValue2, MultiPage1, 1)
End Sub

Private Sub Check2Box6_Click()
    Call CheckBoxClickNew(Check2Box6, TxtBasicInfo, dic2, "建筑结构", Page2 + "." + "建筑结构", TextValue2, MultiPage1, 1)
End Sub

Private Sub Check2Box7_Click() '工程位置

    'MultiPage1.Pages.Item(2).Visible = False
    
    
    Call CheckBoxClickNew(Check2Box7, TxtBasicInfo, dic2, "工程位置", Page2 + "." + "工程位置", TextValue2, MultiPage1, 1)
    
    'UserForm1.Show 0

End Sub


Private Sub Check2Box8_Click()
    Call CheckBoxClickNew(Check2Box8, TxtBasicInfo, dic2, "建筑面积", Page2 + "." + "建筑面积", TextValue2, MultiPage1, 1)
End Sub

Private Sub Check2Box10_Click()
    Call CheckBoxClickNew(Check2Box10, TxtBasicInfo, dic2, "投资数目", Page2 + "." + "投资数目", TextValue2, MultiPage1, 1)
End Sub

Private Sub Check2Box11_Click()
    Call CheckBoxClickNew(Check2Box11, TxtBasicInfo, dic2, "其他", Page2 + "." + "其他", TextValue2, MultiPage1, 1)
End Sub

Private Sub Check2Box9_Click()
    Call CheckBoxClickNew(Check2Box9, TxtBasicInfo, dic2, "楼层详细", Page2 + "." + "楼层详细", TextValue2, MultiPage1, 1)
End Sub

'工程概况checkbox---end

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

Private Sub CmdChange_Click() '开始更改
    Dim bkMark As Range '定义一个bkMark的range对象和一个bookmarkName字符串对象
    

    Call userWriteBookMarkName(bookMarkName(1), 项目名称, True)
    Call userWriteBookMarkName(bookMarkName(2), 委托单位, True)
    Call userWriteBookMarkName(bookMarkName(3), 开始时间, True)
    Call userWriteBookMarkName(bookMarkName(4), 报告日期, True)
    Call userWriteBookMarkName(bookMarkName(5), 报告号, True)

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
    
    MsgBox "数据写入完成！"
        
   
    
End Sub



Private Sub CmdExit_Click()
    FrmBgRpt.Hide
    End
End Sub


Private Sub CmdRead_Click() '读取数据

    Dim bkMark As Range '定义一个bkMark的range对象和一个bookmarkName字符串对象

    DimBMN

    
    Call bookMarkNameCheckExistWrite(bookMarkName(1), 项目名称)
    Call bookMarkNameCheckExistWrite(bookMarkName(2), 委托单位)
    Call bookMarkNameCheckExistWrite(bookMarkName(3), 开始时间)
    Call bookMarkNameCheckExistWrite(bookMarkName(4), 报告日期)
    Call bookMarkNameCheckExistWrite(bookMarkName(5), 报告号)
    
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

    If ((DateValue(报告日期.Text) < DateValue(开始时间.Text))) Then
        MsgBox "日期数据有误，请检查！", vbOKOnly, "日期关系错误提示"
        Exit Sub
    End If



End Sub


Private Sub Command2Button1_Click() '写入工程概况
    
    Call userWriteBookMarkName(bookMarkName(6), TxtBasicInfo)
    
    Call ReplaceTextwithCrossRef(bookMarkName(6), bookMarkName(1), TxtBasicInfo)
    
    Call ReplaceTextwithCrossRef(bookMarkName(6), bookMarkName(2), TxtBasicInfo)

End Sub


Private Sub Command4Button1_Click() '写入编制依据
    
     Call userWriteBookMarkName(bookMarkName(8), TBRptAcc)
    
     Call ReplaceTextwithCrossRef(bookMarkName(8), bookMarkName(1), TBRptAcc)
     
End Sub


Private Sub Command5Button1_Click() '编制方法
    Call userWriteBookMarkName(bookMarkName(9), TBRptMtd)
End Sub

Private Sub Command1Button3_Click() '1-基本信息写入


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
''  .Text = "这是加载的动态控件 "
''  .Name = "xxxcvdsfds"
''  .Width = 320
''
''  End With
    
    Call userWriteBookMarkName(bookMarkName(1), 项目名称, True)
    Call userWriteBookMarkName(bookMarkName(2), 委托单位, True)
    Call userWriteBookMarkName(bookMarkName(3), 开始时间, True)
    Call userWriteBookMarkName(bookMarkName(4), 报告日期, True)
    Call userWriteBookMarkName(bookMarkName(5), 公司报告号, True)
    Call userWriteBookMarkName(bookMarkName(13), 部门报告号, True)

    Dim My1Str1 As String, My1Str2, My1Str3, My1Str4, My1Str5

    My1Str1 = 项目名称.Text
    My1Str5 = 委托单位.Text
    My1Str2 = 开始时间.Text
    My1Str3 = 报告日期.Text
    My1Str4 = 公司报告号.Text

    
    Call ModfStringFromCustomAttrib(Page2 + ".工程位置", My1Str1 + "位于________________。")
    
    Call ModfStringFromCustomAttrib(Page2 + ".建设单位", "本工程建设单位为" + 委托单位.Text + ",")
    
    Call ModfStringFromCustomAttrib(Page3 + ".包括", "本次编制范围：" + My1Str1 + "，含桩基工程、基础工程、地坪工程、辅房主体结构工程等。")
    
    Call ModfStringFromCustomAttrib(Page4 + ".招标文件", "《" + My1Str1 + "预算编制要求》、招标文件。" & vbCr)
    
    Call ModfStringFromCustomAttrib(Page8 + ".预算明细表", "《" + My1Str1 + "预算编制明细表》" & vbCr)
    
    UpAF


End Sub


'写入编制结果标签
Private Sub CommandButton62_Click()

    Call userWriteBookMarkName(bookMarkName(10), TBRptCkl)
    
End Sub



'写入其他说明标签
Private Sub CommandButton72_Click()
    
    Call userWriteBookMarkName(bookMarkName(11), TBRptOth)
    
    
End Sub


'写入附件标签
Private Sub CommandButton82_Click()
    
    Call userWriteBookMarkName(bookMarkName(12), AuditorText)
    
    Call ReplaceTextwithCrossRef(bookMarkName(12), bookMarkName(1), AuditorText)
    
End Sub


'写入编制范围标签
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
    v(0) = StrReverse(reg.Replace(StrReverse(v(0)), "$1,")) '每隔3数字加逗号
    reg.Pattern = "^([^\d]*?),"
    v(0) = reg.Replace(v(0), "$1")
    intSelStart = objTxt.SelStart
    objTxt.Text = Join(v, ".")
    objTxt.SelStart = intSelStart + Len(v(0)) - Len(s1)
    
    textCont.Text = "本工程预算编制价 " & Join(v, ".")
    
End Sub


