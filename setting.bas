Attribute VB_Name = "setting"

Public Fit As Boolean

Public Const Page1 As String = "基本信息"
Public Const Page2 As String = "工程概况"
Public Const Page3 As String = "编制范围"
Public Const Page4 As String = "编制依据"
Public Const Page5 As String = "编制方法"
Public Const Page51 As String = "工料单价法"
Public Const Page52 As String = "综合单价法"
Public Const Page53 As String = "全费用综合单价法"
Public Const Page6 As String = "编制结果"
Public Const Page7 As String = "其他说明"
Public Const Page8 As String = "附件"



Public Function DelNumberReg(ByVal s As String) As String
 
    Dim oRegExp As Object
    Dim strDest As String, strSource As String
    
    strSource = s
    Set oRegExp = CreateObject("VBscript.RegExp")
    oRegExp.Global = True
    oRegExp.Pattern = "^([0-9])(、)"
     
    strDest = oRegExp.Replace(strSource, "")
   ' strDest = Replace(strDest, "、", "")
    
    DelNumberReg = strDest
       
End Function



Function UpComboBox(ByRef ComboBox As Control, Optional ByRef dicKeyName As Boolean)
    If dicKeyName = True Then
        ComboBox.AddItem Page51
        ComboBox.AddItem Page52
        ComboBox.AddItem Page53
    Else
        ComboBox.AddItem Page1
        ComboBox.AddItem Page2
        ComboBox.AddItem Page3
        ComboBox.AddItem Page4
        ComboBox.AddItem Page5
        ComboBox.AddItem Page6
        ComboBox.AddItem Page7
        ComboBox.AddItem Page8
    End If
End Function

Function UpListBox(ByRef ListBox As Control)
    
    
    For Each prop In ActiveDocument.CustomDocumentProperties
                           
        ListBox.AddItem prop.Name
        
    Next
End Function

Public Function CChinese(ByVal Number As Currency) As String
    Number = Val(Trim(Number))
    If Number = 0 Then CurrencyToStr = "": Exit Function
    Dim str1Ary As Variant, str2Ary As Variant
    str1Ary = Split("零 壹 贰 叁 肆 伍 陆 柒 捌 玖")
    str2Ary = Split("分 角 元 拾 佰 仟 万 拾 佰 仟 亿 拾 佰 仟 万 拾 佰")
    Dim A As Long, B As Long '循环基数
    Dim tmp1 As String '临时转换
    Dim tmp2 As String '临时转换结果
    Dim Point As Long '小数点位置
    tmp1 = Round(Number, 2)
    tmp1 = Replace(tmp1, "-", "") '先去掉“-”号
    Point = InStr(tmp1, ".") '取得小数点位置
    If Point = 0 Then '如果有小数点，最大佰万亿
    B = Len(tmp1) + 2 '加2位小数
    Else
    B = Len(Left(tmp1, Point + 1)) '包括点加2位小数
    End If
    ''先将所有数字替换为中文
    For A = 9 To 0 Step -1
    tmp1 = Replace(Replace(tmp1, A, str1Ary(A)), ".", "")
    Next
    For A = 1 To B
    B = B - 1
    If Mid(tmp1, A, 1) <> "" Then
    If B > UBound(str2Ary) Then Exit For
    tmp2 = tmp2 & Mid(tmp1, A, 1) & str2Ary(B)
    End If
    Next
    If tmp2 = "" Then CurrencyToStr = "": Exit Function
    '去掉多余的零
    For A = 1 To Len(tmp2)
    tmp2 = Replace(tmp2, "零亿", "亿零")
    tmp2 = Replace(tmp2, "零万", "万零")
    tmp2 = Replace(tmp2, "零仟", "零")
    tmp2 = Replace(tmp2, "零佰", "零")
    tmp2 = Replace(tmp2, "零拾", "零")
    tmp2 = Replace(tmp2, "零元", "元")
    tmp2 = Replace(tmp2, "零零", "零")
    tmp2 = Replace(tmp2, "亿万", "亿")
    Next
    
    
    If Point = 1 Then tmp2 = "零元" + tmp2
    If Number < 0 Then tmp2 = "负" + tmp2
    If Point = 0 Then tmp2 = tmp2 + "整"
        
    CChinese = tmp2
    'MsgBox CurrencyToStr
        
End Function

Public Function addLineBreak(ByRef TextBoxCont As Control, ByRef dicCont As Object, ByRef outBoxCont As Control, Optional ByRef idiot As Boolean)
    TextBoxCont.Text = TextBoxCont.Text + vbCr
    If idiot = True Then
    
        Call SyncEditTextWrap(TextBoxCont, dicCont, outBoxCont, True)
    Else
    
        Call SyncEditTextWrap(TextBoxCont, dicCont, outBoxCont)
    End If
    
End Function

Public Function SyncEditTextLine(ByRef TextBoxCont As Control, ByRef dicCont As Object, ByRef outBoxCont As Control)
   
    Dim customAttribKey As String, dicKeyName
    If TextBoxCont.Tag <> "" Then
        v = Split(TextBoxCont.Tag, "$")
        
        customAttribKey = v(0)
        dicKeyName = v(1)
        
        dicCont.Item(dicKeyName) = TextBoxCont.Text
        outBoxCont.Text = "    " + LTrim(Join(dicCont.Items(), ""))
        
        Call ModfStringFromCustomAttrib(customAttribKey, TextBoxCont.Text)

    End If

End Function


Public Function readPage1(ByRef TextBoxCont As Control, ByRef customAttribKey As String)
    TextBoxCont.Text = ReadStringFromCustomAttrib(customAttribKey)
    TextBoxCont.Tag = customAttribKey
    
    'Call AddBooleanFromCustomAttrib(CheckBoxCont.Name, False)
End Function

Public Function SyncEditSingleText(ByRef TextBoxCont As Control, ByRef FrmBgRpt As FrmBgRpt, Optional ByRef dic2 As Object, _
Optional ByRef dic3 As Object, Optional ByRef dic4 As Object, Optional ByRef dic8 As Object)
   
    Dim oRegExp As Object, oRegExp2
    Dim TempReplaceString As String, TempReplaceString2, TempReplaceString3, TempReplaceString4, TempReplaceString5, TempReg, TempReg2
    
    
    If TextBoxCont.Tag <> "" Then
       
        'Debug.Print ReadStringFromCustomAttrib(Page1 + "." + "项目名称")
        
        If TextBoxCont.Tag = Page1 + "." + "项目名称" Then
    
             '取上一次textBox内容 作为正则，读取本次修改的内容。在总字符串里替换
            TempReg = ReadStringFromCustomAttrib(Page1 + "." + "项目名称")

            TempReplaceString = ReadStringFromCustomAttrib(Page2 + ".工程位置")
            TempReplaceString2 = ReadStringFromCustomAttrib(Page3 + ".包括")
            TempReplaceString3 = ReadStringFromCustomAttrib(Page4 + ".招标文件")
            TempReplaceString4 = ReadStringFromCustomAttrib(Page8 + ".预算明细表")
            
            
            Set oRegExp = CreateObject("VBscript.RegExp")
            oRegExp.Global = True
            oRegExp.Pattern = TempReg
            
            
            Call ModfStringFromCustomAttrib(Page2 + ".工程位置", oRegExp.Replace(TempReplaceString, TextBoxCont.Text)) '更新工程位置
            Call ModfStringFromCustomAttrib(Page3 + ".包括", oRegExp.Replace(TempReplaceString2, TextBoxCont.Text)) '更新编制范围
            Call ModfStringFromCustomAttrib(Page4 + ".招标文件", oRegExp.Replace(TempReplaceString3, TextBoxCont.Text)) '更新招标文件
            Call ModfStringFromCustomAttrib(Page8 + ".预算明细表", oRegExp.Replace(TempReplaceString4, TextBoxCont.Text)) '更新工程位置
            
            Call ModfStringFromCustomAttrib(TextBoxCont.Tag, TextBoxCont.Text) '更新工程名称
            
            
            Call CheckBoxClickNew(FrmBgRpt.Check2Box7, FrmBgRpt.TxtBasicInfo, dic2, "工程位置", Page2 + "." + "工程位置", FrmBgRpt.TextValue2, FrmBgRpt.MultiPage1, 1)
            Call CheckBoxClickNew(FrmBgRpt.Check3Box1, FrmBgRpt.TBRptRg, dic3, "包括", Page3 + "." + "包括", FrmBgRpt.TextValue3, FrmBgRpt.MultiPage1, 2)
            Call CheckBoxClickNum(FrmBgRpt.Check4Box1, FrmBgRpt.TBRptAcc, dic4, "第1条", Page4 + "." + "招标文件", FrmBgRpt.TextValue4, FrmBgRpt.MultiPage1, 3)
            Call CheckBoxClickNum(FrmBgRpt.Check8Box1, FrmBgRpt.AuditorText, dic8, "第1条", Page8 + "." + "预算明细表", FrmBgRpt.TextValue8, FrmBgRpt.MultiPage1, 7, True)
                
        ElseIf TextBoxCont.Tag = Page1 + "." + "委托单位" Then
            
            TempReg2 = ReadStringFromCustomAttrib(Page1 + "." + "委托单位")
            TempReplaceString5 = ReadStringFromCustomAttrib(Page2 + ".建设单位")
             
            Set oRegExp2 = CreateObject("VBscript.RegExp")
            oRegExp2.Global = True
            oRegExp2.Pattern = TempReg2
            
            Call ModfStringFromCustomAttrib(Page2 + ".建设单位", oRegExp2.Replace(TempReplaceString5, TextBoxCont.Text)) '更新建设单位
            
            Call ModfStringFromCustomAttrib(TextBoxCont.Tag, TextBoxCont.Text) '更新工程名称
            
            Call CheckBoxClickNew(FrmBgRpt.Check2Box5, FrmBgRpt.TxtBasicInfo, dic2, "建设单位", Page2 + "." + "建设单位", FrmBgRpt.TextValue2, FrmBgRpt.MultiPage1, 1)
        
        Else
            
            Call ModfStringFromCustomAttrib(TextBoxCont.Tag, TextBoxCont.Text) '更新其他
            
        End If

    End If

End Function


Public Function SyncEditTextWrap(ByRef TextBoxCont As Control, ByRef dicCont As Object, ByRef outBoxCont As Control, Optional ByRef idiot As Boolean)
   
    Dim customAttribKey As String, dicKeyName
    
    Dim ArrayTemp(10) As String
    Dim TempCol As New Collection

    
    If TextBoxCont.Tag <> "" Then
        v = Split(TextBoxCont.Tag, "$")
        
        customAttribKey = v(0)
        dicKeyName = v(1)
        
        dicCont.Item(dicKeyName) = TextBoxCont.Text
        
        
        'outBoxCont.Text = Join(dicCont.Items(), "")
        
        
        For i = 1 To UBound(dicCont.Items) + 1
            If (dicCont.Item("第" & i & "条") <> "") Then
                TempCol.add DelNumberReg(dicCont.Item("第" & i & "条"))
            End If
        Next i
        
        For j = 1 To TempCol.count
            If idiot = True Then
                    ArrayTemp(j - 1) = "    " & TempCol(j)
                Else
                    ArrayTemp(j - 1) = "    " & CStr(j) & "、" & TempCol(j)
            End If
        Next j
        
        outBoxCont.Text = Join(ArrayTemp, "")
                        
        Call ModfStringFromCustomAttrib(customAttribKey, TextBoxCont.Text)
               
    End If

End Function


Public Function bindDataInTextBox(ByRef richBoxCont As Control, Optional ByRef customAttribKey As String, _
Optional ByRef dicKeyName As String)
   
   richBoxCont.Tag = customAttribKey + "$" + dicKeyName
    
   
End Function

Public Function checkDicKeyValueExist(ByRef dicCont As Object, ByVal key As String) As Boolean
    If dicCont.Item(key) <> "" Then
        checkDicKeyValueExsit = True
    Else
        checkDicKeyValueExsit = False
    
End Function


Public Function bookMarkNameReadFromDic(ByRef dicCont As Object, ByRef richBoxCont As Control, Optional ByRef No As Boolean)
    
    If No = True Then
        Dim ArrayTemp(10) As String
        Dim TempCol As New Collection
        
        For i = 1 To UBound(dicCont.Items) + 1
            If (dicCont.Item("第" & i & "条") <> "") Then
                TempCol.add DelNumberReg(dicCont.Item("第" & i & "条"))
            End If
        Next i
        
        For j = 1 To TempCol.count
            ArrayTemp(j - 1) = CStr(j) & "、" & TempCol(j)
        Next j
        
        richBoxCont.Text = Join(ArrayTemp, "")
    
    Else
        richBoxCont.Text = Join(dicCont.Items(), "")
    
    End If
    
End Function

Public Function CheckBoxClick(ByRef CheckBoxCont As Control, ByRef richBoxCont As Control, _
ByRef dicCont As Object, ByRef TextBoxCont As Control, ByRef textName As String)
    Dim TempCont As String
    TempCont = richBoxCont.Text
    If CheckBoxCont.value = False Then
        TempCont = Replace(TempCont, dicCont.Item(textName), "")
        dicCont.Item(textName) = ""
        richBoxCont.Text = TempCont
    End If
    If CheckBoxCont.value = True Then
        dicCont.Item(textName) = TextBoxCont.Text
        'Debug.Print textBoxCont.Text
        richBoxCont.Text = TempCont + TextBoxCont.Text
    End If
End Function

Public Function CheckedAllBox(ByRef MultiPage As Control, count As Integer)
    Dim Ctr As Control
    For Each Ctr In MultiPage.Pages.Item(count).Controls
        If TypeName(Ctr) = "CheckBox" Then
            Ctr.ForeColor = &H80000012
            Ctr.value = True
        End If
    Next
End Function

Public Function CheckedAllBoxColorRestore(ByRef MultiPage As Control, count As Integer)
    Dim Ctr As Control
    For Each Ctr In MultiPage.Pages.Item(count).Controls
        If TypeName(Ctr) = "CheckBox" Then
            Ctr.ForeColor = &H80000012
        End If
    Next
End Function

Public Function checkValueState(ByRef MultiPage As Control, count As Integer, ByRef dicCont As Object)

    Dim Ctr As Control
    For Each Ctr In MultiPage.Pages.Item(count).Controls
        If TypeName(Ctr) = "CheckBox" Then
           Ctr.ForeColor = &H80000012
           If Ctr.value = False Then
                dicCont.Item(Ctr.Tag) = ""
           Else
           
           End If
        End If
    Next
    
End Function

Public Function checkExist(ByRef key As String) As Boolean

    For Each prop In ActiveDocument.CustomDocumentProperties
                       
        If prop.Name = key Then
            checkExist = True
            
            Exit For
        Else
            checkExist = False
        End If
    Next

End Function


Public Function CheckBoxClickNew(ByRef CheckBoxCont As Control, ByRef richBoxCont As Control, _
ByRef dicCont As Object, ByRef dicKeyName As String, _
ByRef customAttribKey As String, ByRef showHideText As Control, Optional MultiPage As Control, Optional count As Integer)
    
    Dim CustomAttribValue As String
    CustomAttribValue = ReadStringFromCustomAttrib(customAttribKey)
    
        
    'Debug.Print "dicKeyName" & dicKeyName
    'Debug.Print "CustomAttribValue" & CustomAttribValue
    
    Call CheckedAllBoxColorRestore(MultiPage, count)
    
    If CheckBoxCont.value = False Then
    
        If checkExist(CheckBoxCont.Name) = True Then
           Call ModfBooleanFromCustomAttrib(CheckBoxCont.Name, False)
        Else
           Call AddBooleanFromCustomAttrib(CheckBoxCont.Name, False)
           'Call AddStringFromCustomAttrib(customAttribKey, "此自定义属性需要扩展，请在此框中编辑内容")
        End If
        

        dicCont.Item(dicKeyName) = ""
        
        showHideText.Text = ""
        showHideText.Visible = False
        richBoxCont.Text = "    " + Join(dicCont.Items(), "")
                
    End If
    
    If CheckBoxCont.value = True Then
    
         dicCont.Item(dicKeyName) = CustomAttribValue
         If checkExist(CheckBoxCont.Name) = True Then
            Call ModfBooleanFromCustomAttrib(CheckBoxCont.Name, True)
         Else
            Call AddBooleanFromCustomAttrib(CheckBoxCont.Name, True)
            'Call AddStringFromCustomAttrib(customAttribKey, "此自定义属性需要扩展，请在此框中编辑内容")
         End If
                       
         Call checkValueState(MultiPage, count, dicCont)
         CheckBoxCont.ForeColor = &HFF&
         showHideText.Visible = True
         showHideText.Text = CustomAttribValue
         Call bindDataInTextBox(showHideText, customAttribKey, dicKeyName)
         richBoxCont.Text = "    " + LTrim(Join(dicCont.Items(), ""))
            
    End If
    
End Function




Public Function CheckBoxClickNum(ByRef CheckBoxCont As Control, ByRef richBoxCont As Control, _
ByRef dicCont As Object, ByRef dicKeyName As String, _
ByRef customAttribKey As String, ByRef showHideText As Control, Optional MultiPage As Control, Optional count As Integer, Optional ByRef idiot As Boolean)
    
    Dim CustomAttribValue As String
    CustomAttribValue = ReadStringFromCustomAttrib(customAttribKey)
    
    'Debug.Print "dicKeyName" & dicKeyName
    'Debug.Print "CustomAttribValue" & CustomAttribValue
    
    Dim StrTemp As String
    Dim ArrayTemp(10) As String
    Dim ArrayTemp2(10) As String
    Dim TempCol As New Collection
    Dim TempCol2 As New Collection
    
    Call CheckedAllBoxColorRestore(MultiPage, count)
    
    If CheckBoxCont.value = False Then
        
        If checkExist(CheckBoxCont.Name) = True Then
           Call ModfBooleanFromCustomAttrib(CheckBoxCont.Name, False)
        Else
           Call AddBooleanFromCustomAttrib(CheckBoxCont.Name, False)
        End If
        
        dicCont.Item(dicKeyName) = ""
        showHideText.Text = ""
        showHideText.Visible = False
        richBoxCont.Text = Join(dicCont.Items(), "")
        
        For i = 1 To UBound(dicCont.Items) + 1
            If (dicCont.Item("第" & i & "条") <> "") Then
                TempCol.add DelNumberReg(dicCont.Item("第" & i & "条"))
            End If
        Next i
        
        For j = 1 To TempCol.count
            If idiot = True Then
                ArrayTemp(j - 1) = "    " & TempCol(j)
            Else
                ArrayTemp(j - 1) = "    " & CStr(j) & "、" & TempCol(j)
            End If
        Next j
        
        richBoxCont.Text = Join(ArrayTemp, "")
        
        
    End If
    
    If CheckBoxCont.value = True Then
    
         If checkExist(CheckBoxCont.Name) = True Then
            Call ModfBooleanFromCustomAttrib(CheckBoxCont.Name, True)
         Else
            Call AddBooleanFromCustomAttrib(CheckBoxCont.Name, True)
         End If
               
        
        dicCont.Item(dicKeyName) = CustomAttribValue
        Call checkValueState(MultiPage, count, dicCont)
        CheckBoxCont.ForeColor = &HFF&
        showHideText.Visible = True
        showHideText.Text = CustomAttribValue
        
        
        Call bindDataInTextBox(showHideText, customAttribKey, dicKeyName)
        
        For i = 1 To UBound(dicCont.Items) + 1
            If (dicCont.Item("第" & i & "条") <> "") Then
                TempCol2.add DelNumberReg(dicCont.Item("第" & i & "条"))
            End If
        Next i
        
        For j = 1 To TempCol2.count
            If idiot = True Then
                ArrayTemp2(j - 1) = "    " & TempCol2(j)
            Else
                ArrayTemp2(j - 1) = "    " & CStr(j) & "、" & TempCol2(j)
            End If
        Next j
        richBoxCont.Text = Join(ArrayTemp2, "")
    End If
    
End Function


Function createEditBox(MultiPage As Control, count As Integer, ByRef CustomAttribValue As String, _
ByRef customAttribKey As String, ByRef dicKeyName As String, ByRef dicCont As Object, ByRef richBoxCont As Control)
    
    Dim TextValue3 As Control
    Set TextValue3 = MultiPage.Pages.Item(count).Controls.add("Forms.TextBox.1")
    
    With TextValue3
        .Visible = True
        .MultiLine = True
        .Top = 55
        .Width = 445
        .Left = 12
        .Height = 16
        .Text = CustomAttribValue
        .Tag = customAttribKey + "$" + dicKeyName
        '.OnChange = "SyncEditText("&TextValue3&, dic, TxtBasicInfo)"
    End With
    
    'Call SyncEditText(TextValue3, dicCont, richBoxCont)
    '"'showORhideRows " & iRowStart & "," & iRowEnd & "'"
    'showORhideRows(iRowStart, iRowEnd)
    
    'Call bindDataInTextBox(TextValue3, customAttribKey, dicKeyName)
    
    'Call TextValue3_Change(TextValue3, dicCont, richBoxCont)
    
    
End Function



Public Function formatCheckBoxInLine(ByRef CheckBoxCont As Control, ByRef dicCont As Object, _
ByRef dicKey As String, ByRef customAttribKey As String)
    dicCont.Item(dicKey) = ReadStringFromCustomAttrib(customAttribKey)
    CheckBoxCont.Tag = dicKey
    
    'Call AddBooleanFromCustomAttrib(CheckBoxCont.Name, False)
End Function


Public Function formatCheckBoxWrap(ByRef CheckBoxCont As Control, ByRef dicCont As Object, _
ByRef dicKey As String, ByRef customAttribKey As String)
    dicCont.Item(dicKey) = ReadStringFromCustomAttrib(customAttribKey)
    CheckBoxCont.Tag = dicKey
    
    'Call AddBooleanFromCustomAttrib(CheckBoxCont.Name, False)
    
End Function


Public Function formatTextBoxWrap(ByRef TextBoxCont As Control, ByRef KeyName As String)
        
    Dim TempCont As String
    TempCont = Replace(TextBoxCont.Text, vbCrLf, "")
    
    TextBoxCont.Text = Chr(32) + Chr(32) + Chr(32) + Chr(32) + "#" + TempCont + vbCr
    TextBoxCont.Visible = False
    
    'Call AddStringFromCustomAttrib(KeyName, TempCont)
      
End Function

Public Function formatTextBoxInLine(ByRef TextBoxCont As Control, ByRef KeyName As String)
    
    Dim TempCont As String
    TempCont = Replace(TextBoxCont.Text, vbCrLf, "")
    
    TextBoxCont.Text = "#" + TempCont
    TextBoxCont.Visible = False
    
    'Call AddStringFromCustomAttrib(KeyName, TempCont)
    
End Function


Public Function bookMarkNameCheckExistWrite(ByRef bookMarkName As String, ByRef richBoxCont As Control)
    If ActiveDocument.Bookmarks.Exists(bookMarkName) = True Then
        Set bkMark = ActiveDocument.Bookmarks(bookMarkName).Range  '查找书签“bookmark1”并赋值给bkMark
        bkMark.Select '选中bkMark书签对应的文本
        richBoxCont.Text = bkMark.Text  '工程概况
    Else
        MsgBox "不存在" + bookMarkName + "书签内容，请添加" + bookMarkName + "书签"
    End If

End Function




Public Function userWriteBookMarkName(ByRef bookMarkName As String, ByRef richBoxCont As Control, Optional ByRef wait As Boolean)
    
    If Replace(richBoxCont.Text, vbCrLf, "") <> "" Then
        If ActiveDocument.Bookmarks.Exists(bookMarkName) = True Then
            Set bkMark = ActiveDocument.Bookmarks(bookMarkName).Range
            bkMark.Select
            bkMark.Text = richBoxCont.Text
            ActiveDocument.Bookmarks.add bookMarkName, bkMark
            If wait = True Then
            Else
               UpAF
            End If
        Else
            MsgBox "不存在" + bookMarkName + "书签内容，请添加" + bookMarkName + "书签"
    End If
    Else
        MsgBox "注意：" + bookMarkName + "内容为空，请写入内容"
    End If
    
End Function


Public Function UpAF()

    Dim aField As Field
    Dim aStory As Range
    ''' Update all fields in the document
    For Each aStory In ActiveDocument.StoryRanges
       For Each aField In aStory.Fields
          aField.Update
       Next aField
    Next aStory
        
    If Fit = True Then
        If ActiveWindow.View.DisplayPageBoundaries = True Then
            ActiveWindow.View.DisplayPageBoundaries = False
        Else
            ActiveWindow.View.DisplayPageBoundaries = True
        
        End If
    End If
    
End Function


Public Function ReplaceTextwithCrossRef(ByRef parentBookMarkName As String, ByRef bookMarkName As String, ByRef richBoxCont As Control)

    If ActiveDocument.Bookmarks.Exists(bookMarkName) = True Then
        Set childBkMark = ActiveDocument.Bookmarks(bookMarkName).Range  '查找书签“bookmark1”并赋值给bkMark
        Set parentBkMark = ActiveDocument.Bookmarks(parentBookMarkName).Range
        parentBkMark.Select
            
            
        Selection.CopyFormat
        'Debug.Print Selection
        
         With Selection.Find
          .ClearFormatting
          .Text = childBkMark
          .Replacement.Text = ""
          .Format = False
          .MatchWildcards = False
          .Wrap = wdFindStop
          .Execute
         End With
         If Selection.Find.Found Then
           If Selection.Bookmarks.Exists(bookMarkName) Then
           Else
            Selection.InsertCrossReference ReferenceType:=wdRefTypeBookmark, ReferenceKind:=wdContentText, ReferenceItem:=bookMarkName
           End If
         End If
      
         parentBkMark.Select
         With Selection.Find
          .ClearFormatting
          .Text = childBkMark
          .Replacement.Text = ""
          .Format = False
          .MatchWildcards = False
          .Wrap = wdFindStop
          .Execute
         End With
         If Selection.Find.Found Then
            If Selection.Bookmarks.Exists(bookMarkName) Then
           
            Else
                Selection.PasteFormat
            End If
        End If
        
    Else
        MsgBox "不存在" + bookMarkName + "书签内容，请添加" + bookMarkName + "书签"
    End If

    
End Function
