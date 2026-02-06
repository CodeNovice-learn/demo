import os
import win32com.client as win32

def write_action_learning_plan():
    # 打开 Word 应用
    word = win32.Dispatch("Word.Application")
    word.Visible = True

    # 打开文档
    desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
    doc_path = os.path.join(desktop, '新建文档.docx')
    doc = word.Documents.Open(doc_path)

    # 清除内容
    doc.Content.Delete()

    # 写入内容
    content = """行动学习策划方案

一、项目背景

行动学习(Action Learning)是一种通过解决实际业务问题来促进学习与发展的创新方法。本项目旨在通过行动学习模式，提升团队的协作能力、问题解决能力和创新能力，同时推动实际业务问题的有效解决。

二、项目目标

1. 提升团队成员的批判性思维和问题解决能力
2. 增强团队协作和沟通效率
3. 解决当前面临的关键业务问题
4. 培养团队的学习型组织文化
5. 建立可持续的行动学习机制

三、参与人员

1. 项目发起人：负责项目整体指导和资源协调
2. 行动学习教练：负责引导学习过程和反思环节
3. 学习小组：4-6人组成，负责实际问题的研讨和解决
4. 业务专家：提供专业支持和咨询

四、学习议题

1. 当前业务瓶颈分析及解决方案
2. 跨部门协作优化
3. 客户满意度提升策略
4. 创新思维培养
5. 领导力发展实践

五、实施流程

1. 问题识别：各小组识别并确定要解决的实际问题
2. 方案设计：小组讨论并制定解决方案
3. 行动实施：在实际行动中验证和完善方案
4. 反思学习：定期回顾和反思学习成果
5. 成果分享：向其他小组和领导层汇报成果

六、时间安排

第一阶段（第1周）：项目启动，组建学习小组，确定议题
第二阶段（第2-3周）：问题分析和方案设计
第三阶段（第4-6周）：方案实施和持续优化
第四阶段（第7-8周）：反思总结和成果分享

七、预期成果

1. 完成2-3个关键业务问题的解决方案
2. 提升团队成员的综合能力
3. 建立行动学习的方法论和工具库
4. 形成可持续的学习机制
5. 产出可复用的行动学习案例集

八、评估机制

1. 过程评估：定期检查各小组的进展情况
2. 结果评估：评估问题解决的实际效果
3. 学习评估：通过问卷和访谈评估学习成果
4. 同伴评估：小组成员之间的互评
5. 专家评估：邀请业务专家对成果进行评审

九、风险控制

1. 成员参与度不足风险：建立激励机制，定期沟通
2. 议题选择不当风险：教练引导，确保议题的实用性
3. 资源投入不足风险：提前规划和争取必要资源
4. 效果不明显风险：设置阶段性目标，及时调整方向
5. 知识流失风险：建立知识管理体系，做好文档归档

十、资源配置

1. 人力资源：教练、业务专家、学习小组成员
2. 时间资源：每周至少4小时的学习和讨论时间
3. 场地资源：会议室、学习空间
4. 技术资源：在线协作工具、学习平台
5. 预算资源：培训费用、材料费用、奖励费用
"""

    # 分行处理
    lines = content.split('\n')
    for line in lines:
        if line.strip():
            para = doc.Paragraphs.Add()
            para.Range.Text = line
            para.Range.Font.Size = 12
            if line.startswith(('一、', '二、', '三、', '四、', '五、', '六、', '七、', '八、', '九、', '十、')):
                para.Range.Font.Size = 16
                para.Range.Font.Bold = True
            elif line == '行动学习策划方案':
                para.Range.Font.Size = 22
                para.Range.Font.Bold = True
                para.Range.ParagraphFormat.Alignment = 1  # 居中

    # 保存并关闭
    doc.Save()
    doc.Close()
    word.Quit()

    print("Action learning plan has been written to the document!")

if __name__ == "__main__":
    write_action_learning_plan()
