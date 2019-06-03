Attribute VB_Name = "model"



Public Function ReadStringFromCustomAttrib(ByRef key As String) As String
    On Error GoTo MyErr
        ReadStringFromCustomAttrib = ActiveDocument.CustomDocumentProperties(key).value
    
    Exit Function
MyErr:

    Msg = " 读取数据出错!请检查自定义属性 " & key & " 是否存在"

    MsgBox Msg
    
    
End Function



Public Function AddStringFromCustomAttrib(ByVal key As String, ByVal values As String)
    
    ActiveDocument.CustomDocumentProperties.add Name:=key, Type:=msoPropertyTypeString, LinkToContent:=False, value:=values
    
    'MsgBox "添加自定义属性 " + key + " 成功"
    
End Function

Public Function ModfStringFromCustomAttrib(ByRef key As String, ByRef value As String)
    On Error GoTo MyErr
        ActiveDocument.CustomDocumentProperties(key).value = value
   Exit Function
MyErr:

    Msg = " 修改数据时出错!请检查自定义属性 " & key & " 是否存在"

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
    
    Call ModfStringFromCustomAttrib(Page1 + ".项目名称", "________土建工程")
    Call ModfStringFromCustomAttrib(Page1 + ".委托单位", "________公司")
    Call ModfStringFromCustomAttrib(Page1 + ".公司报告号", "浙科佳咨[2018]造14-395-E48")
    Call ModfStringFromCustomAttrib(Page1 + ".部门报告号", "(045)号")
    Call ModfStringFromCustomAttrib(Page1 + ".开始时间", "2018年8月21日")
    Call ModfStringFromCustomAttrib(Page1 + ".报告日期", "2018年9月13日")
    
    
    Call ModfStringFromCustomAttrib(Page2 + ".工程位置", "________土建工程位于________________。")
    Call ModfStringFromCustomAttrib(Page2 + ".建设单位", "本工程建设单位为________公司,")
    Call ModfStringFromCustomAttrib(Page2 + ".建筑结构", "建筑结构为________，")
    Call ModfStringFromCustomAttrib(Page2 + ".建筑面积", "总建筑面积____平方米，")
    Call ModfStringFromCustomAttrib(Page2 + ".楼层详细", "地下__层,地上__层,共__层。")
    Call ModfStringFromCustomAttrib(Page2 + ".投资数目", "项目总投资：______万元。")
    Call ModfStringFromCustomAttrib(Page2 + ".其他", "其他备注内容。")
    
    Call ModfStringFromCustomAttrib(Page3 + ".包括", "本次编制范围：________土建工程，含桩基工程、基础工程、地坪工程、辅房主体结构工程等。")
    Call ModfStringFromCustomAttrib(Page3 + ".不包括", "不含钢结构工程、水电安装工程、辅房内装修工程。")
    Call ModfStringFromCustomAttrib(Page3 + ".其他", "具体详见招标文件。")

    Call ModfStringFromCustomAttrib(Page4 + ".招标文件", "《________土建工程预算编制要求》、招标文件。" & vbCr)
    Call ModfStringFromCustomAttrib(Page4 + ".图纸", "图纸：电子文件名为《_RDC车间更新图纸》A版图。" & vbCr)
    Call ModfStringFromCustomAttrib(Page4 + ".联系函", "预算编制过程中的工作联系函及其他相关资料。" & vbCr)
    Call ModfStringFromCustomAttrib(Page4 + ".编制依据", "《浙江省建筑工程预算定额》（2010版）、《浙江省建设工程施工费用定额》（2010版）、《浙江省施工机械台班费用定额》（2010版）、《建设工程工程量清单计价规范》（GB50500-2013）、《浙江省建设工程计价规则》（2010版） 及相关补充文件等。" & vbCr)
    Call ModfStringFromCustomAttrib(Page4 + ".其他资料", "其他有关资料。" & vbCr)
    
    Call ModfStringFromCustomAttrib(Page5 + ".工料单价法.计算依据", "工程量依据工程量计算规范和设计图纸进行计算。" & vbCr)
    Call ModfStringFromCustomAttrib(Page5 + ".工料单价法.价格来源", "综合单价依据工程量计价规范和浙江省现行相关工程预算定额进行计价，材料价按杭州市______年第__期信息价，无价材料按暂估价或市场价计取。" & vbCr)
    Call ModfStringFromCustomAttrib(Page5 + ".工料单价法.人工单价", "定额人工单价根据____市一类__元/工日 、二类__元/工日、三类__元/工日。" & vbCr)
    Call ModfStringFromCustomAttrib(Page5 + ".工料单价法.费率", "施工费用按《_____省建设工程施工费用定额》（2010版）计取。安全文明施工费按____计，企业管理费、利润按____计取。规费按______计、农民工工伤保险按____计、税金按____计。" & vbCr)
    
    Call ModfStringFromCustomAttrib(Page5 + ".综合单价法.计算依据", "综合单价依据工程量计价规范和浙江省现行相关工程预算定额进行计价，材料价按杭州市______年第__期信息价，无价材料按暂估价或市场价计取。" & vbCr)
    Call ModfStringFromCustomAttrib(Page5 + ".综合单价法.价格来源", "定额人工单价根据____市一类__元/工日 、二类__元/工日、三类__元/工日。" & vbCr)
    Call ModfStringFromCustomAttrib(Page5 + ".综合单价法.人工单价", "施工费用按《_____省建设工程施工费用定额》（2010版）计取。安全文明施工费按____计，企业管理费、利润按____计取。规费按______计、农民工工伤保险按____计、税金按____计。" & vbCr)
    Call ModfStringFromCustomAttrib(Page5 + ".综合单价法.费率", "工程量依据工程量计算规范和设计图纸进行计算。" & vbCr)
    
    Call ModfStringFromCustomAttrib(Page5 + ".全费用综合单价法.计算依据", "工程量依据____省现行相关工程预算定额中工程量计算规则和设计图纸进行计算。" & vbCr)
    Call ModfStringFromCustomAttrib(Page5 + ".全费用综合单价法.价格来源", "单价依据____省现行相关工程预算定额进行计价，材料价按杭州市____年第__期信息价，无价材料按市场价计取。" & vbCr)
    Call ModfStringFromCustomAttrib(Page5 + ".全费用综合单价法.人工单价", "定额人工单价根据____市一类__元/工日 、二类__元/工日、三类__元/工日。" & vbCr)
    Call ModfStringFromCustomAttrib(Page5 + ".全费用综合单价法.费率", "施工费用按《______省建设工程施工费用定额》（2010版）计取。安全文明施工费按__计，企业管理费、利润按__计取。规费按____计、农民工工伤保险按____计、税金按____计。" & vbCr)
    
    Call ModfStringFromCustomAttrib(Page7 + ".水电", "本工程所需的临时用水、用电接驳点由发包人协调总承包人提供，由中标人自行负责接驳并承担一切相关费用，施工用水及用电费用由中标人支付给总承包人，此部分费用已包含于投标报价中。" & vbCr)
    Call ModfStringFromCustomAttrib(Page7 + ".工期及奖罚", "本工程计划工期为____天。工程提前（或拖期）一天竣工奖（罚）金额按工程总造价的万分之二计取，奖罚数额的比例要对等，但总额不得超过工程总造价的百分之三。" & vbCr)
    Call ModfStringFromCustomAttrib(Page7 + ".问题处理", "相关预算编制等中等问题及图纸疑问之处已按建设方的回复处理，具体详见附件工作联系函等。" & vbCr)
    Call ModfStringFromCustomAttrib(Page7 + ".预算非依据", "本次工程预算等中等价不作为工程价款结算依据。" & vbCr)

    Call ModfStringFromCustomAttrib(Page8 + ".预算明细表", "《________土建工程 预算编制明细表》" & vbCr)
    Call ModfStringFromCustomAttrib(Page8 + ".预算编制要求", "《预算编制要求》" & vbCr)
    Call ModfStringFromCustomAttrib(Page8 + ".联系函", "《工作联系函》" & vbCr)
    
End Function
