from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import cm

# 注册中文字体
pdfmetrics.registerFont(TTFont('SimSun', 'C:\\Windows\\Fonts\\simsun.ttc'))
pdfmetrics.registerFont(TTFont('SimHei', 'C:\\Windows\\Fonts\\simhei.ttf'))

def create_pdf():
    c = canvas.Canvas("C:\\Users\\zoufeng\\Desktop\\行动学习策划方案.pdf", pagesize=A4)
    width, height = A4

    # 设置边距
    margin_left = 2 * cm
    margin_top = 2 * cm
    line_height = 0.6 * cm

    # 内容
    content = [
        ("行动学习策划方案", "SimHei", 24, True),
        ("", "SimSun", 12, False),
        ("一、项目背景", "SimHei", 16, True),
        ("行动学习是一种通过解决实际业务问题来促进学习与发展的创新方法。本项目旨在通过行动学习模式，提升团队的协作能力、问题解决能力和创新能力，同时推动实际业务问题的有效解决。", "SimSun", 12, False),
        ("", "SimSun", 12, False),
        ("二、项目目标", "SimHei", 16, True),
        ("1. 提升团队成员的批判性思维和问题解决能力", "SimSun", 12, False),
        ("2. 增强团队协作和沟通效率", "SimSun", 12, False),
        ("3. 解决当前面临的关键业务问题", "SimSun", 12, False),
        ("4. 培养团队的学习型组织文化", "SimSun", 12, False),
        ("5. 建立可持续的行动学习机制", "SimSun", 12, False),
        ("", "SimSun", 12, False),
        ("三、参与人员", "SimHei", 16, True),
        ("1. 项目发起人：负责项目整体指导和资源协调", "SimSun", 12, False),
        ("2. 行动学习教练：负责引导学习过程和反思环节", "SimSun", 12, False),
        ("3. 学习小组：4-6人组成，负责实际问题的研讨和解决", "SimSun", 12, False),
        ("4. 业务专家：提供专业支持和咨询", "SimSun", 12, False),
        ("", "SimSun", 12, False),
        ("四、学习议题", "SimHei", 16, True),
        ("1. 当前业务瓶颈分析及解决方案", "SimSun", 12, False),
        ("2. 跨部门协作优化", "SimSun", 12, False),
        ("3. 客户满意度提升策略", "SimSun", 12, False),
        ("4. 创新思维培养", "SimSun", 12, False),
        ("5. 领导力发展实践", "SimSun", 12, False),
        ("", "SimSun", 12, False),
        ("五、实施流程", "SimHei", 16, True),
        ("1. 问题识别：各小组识别并确定要解决的实际问题", "SimSun", 12, False),
        ("2. 方案设计：小组讨论并制定解决方案", "SimSun", 12, False),
        ("3. 行动实施：在实际行动中验证和完善方案", "SimSun", 12, False),
        ("4. 反思学习：定期回顾和反思学习成果", "SimSun", 12, False),
        ("5. 成果分享：向其他小组和领导层汇报成果", "SimSun", 12, False),
        ("", "SimSun", 12, False),
        ("六、时间安排", "SimHei", 16, True),
        ("第一阶段（第1周）：项目启动，组建学习小组，确定议题", "SimSun", 12, False),
        ("第二阶段（第2-3周）：问题分析和方案设计", "SimSun", 12, False),
        ("第三阶段（第4-6周）：方案实施和持续优化", "SimSun", 12, False),
        ("第四阶段（第7-8周）：反思总结和成果分享", "SimSun", 12, False),
        ("", "SimSun", 12, False),
        ("七、预期成果", "SimHei", 16, True),
        ("1. 完成2-3个关键业务问题的解决方案", "SimSun", 12, False),
        ("2. 提升团队成员的综合能力", "SimSun", 12, False),
        ("3. 建立行动学习的方法论和工具库", "SimSun", 12, False),
        ("4. 形成可持续的学习机制", "SimSun", 12, False),
        ("5. 产出可复用的行动学习案例集", "SimSun", 12, False),
        ("", "SimSun", 12, False),
        ("八、评估机制", "SimHei", 16, True),
        ("1. 过程评估：定期检查各小组的进展情况", "SimSun", 12, False),
        ("2. 结果评估：评估问题解决的实际效果", "SimSun", 12, False),
        ("3. 学习评估：通过问卷和访谈评估学习成果", "SimSun", 12, False),
        ("4. 同伴评估：小组成员之间的互评", "SimSun", 12, False),
        ("5. 专家评估：邀请业务专家对成果进行评审", "SimSun", 12, False),
        ("", "SimSun", 12, False),
        ("九、风险控制", "SimHei", 16, True),
        ("1. 成员参与度不足风险：建立激励机制，定期沟通", "SimSun", 12, False),
        ("2. 议题选择不当风险：教练引导，确保议题的实用性", "SimSun", 12, False),
        ("3. 资源投入不足风险：提前规划和争取必要资源", "SimSun", 12, False),
        ("4. 效果不明显风险：设置阶段性目标，及时调整方向", "SimSun", 12, False),
        ("5. 知识流失风险：建立知识管理体系，做好文档归档", "SimSun", 12, False),
        ("", "SimSun", 12, False),
        ("十、资源配置", "SimHei", 16, True),
        ("1. 人力资源：教练、业务专家、学习小组成员", "SimSun", 12, False),
        ("2. 时间资源：每周至少4小时的学习和讨论时间", "SimSun", 12, False),
        ("3. 场地资源：会议室、学习空间", "SimSun", 12, False),
        ("4. 技术资源：在线协作工具、学习平台", "SimSun", 12, False),
        ("5. 预算资源：培训费用、材料费用、奖励费用", "SimSun", 12, False),
    ]

    y = height - margin_top
    first_page = True

    for text, font_name, font_size, is_bold in content:
        if text == "":
            y -= line_height * 0.5
            continue

        if y < 3 * cm:
            c.showPage()
            y = height - margin_top

        if font_name == "SimHei":
            if is_bold:
                c.setFont("SimHei", font_size)
            else:
                c.setFont("SimSun", font_size)
        else:
            c.setFont(font_name, font_size)

        if is_bold and font_size >= 20:
            c.drawCentredString(width / 2, y, text)
            y -= line_height * 2
        elif is_bold and font_size >= 16:
            c.drawString(margin_left, y, text)
            y -= line_height * 1.5
        else:
            c.drawString(margin_left, y, text)
            y -= line_height

    c.save()
    print("PDF created successfully!")

if __name__ == "__main__":
    create_pdf()
