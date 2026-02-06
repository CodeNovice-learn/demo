$wordApp = New-Object -ComObject Word.Application
$wordApp.Visible = $true
$docPath = [environment]::GetFolderPath('Desktop') + '\新建文档.docx'
$doc = $wordApp.Documents.Open($docPath)

$doc.Content.Delete()

$title = $doc.Paragraphs.Add()
$title.Range.Text = "行动学习策划方案"
$title.Range.Font.Size = 22
$title.Range.Font.Bold = $true
$title.Range.ParagraphFormat.Alignment = 1
$title.Range.InsertParagraphAfter()

$doc.Paragraphs.Add().Range.Text = ""

$section1 = $doc.Paragraphs.Add()
$section1.Range.Text = "一、项目背景"
$section1.Range.Font.Size = 16
$section1.Range.Font.Bold = $true
$section1.Range.InsertParagraphAfter()

$content1 = $doc.Paragraphs.Add()
$content1.Range.Text = "行动学习(Action Learning)是一种通过解决实际业务问题来促进学习与发展的创新方法。本项目旨在通过行动学习模式，提升团队的协作能力、问题解决能力和创新能力，同时推动实际业务问题的有效解决。"
$content1.Range.Font.Size = 12
$content1.Range.ParagraphFormat.LineSpacing = 18
$content1.Range.ParagraphFormat.FirstLineIndent = 24
$content1.Range.InsertParagraphAfter()

$doc.Paragraphs.Add().Range.Text = ""

$section2 = $doc.Paragraphs.Add()
$section2.Range.Text = "二、项目目标"
$section2.Range.Font.Size = 16
$section2.Range.Font.Bold = $true
$section2.Range.InsertParagraphAfter()

$goals = @(
    "1. 提升团队成员的批判性思维和问题解决能力",
    "2. 增强团队协作和沟通效率",
    "3. 解决当前面临的关键业务问题",
    "4. 培养团队的学习型组织文化",
    "5. 建立可持续的行动学习机制"
)

foreach ($goal in $goals) {
    $p = $doc.Paragraphs.Add()
    $p.Range.Text = $goal
    $p.Range.Font.Size = 12
    $p.Range.ParagraphFormat.LineSpacing = 18
    $p.Range.InsertParagraphAfter()
}

$doc.Paragraphs.Add().Range.Text = ""

$section3 = $doc.Paragraphs.Add()
$section3.Range.Text = "三、参与人员"
$section3.Range.Font.Size = 16
$section3.Range.Font.Bold = $true
$section3.Range.InsertParagraphAfter()

$participants = @(
    "1. 项目发起人：负责项目整体指导和资源协调",
    "2. 行动学习教练：负责引导学习过程和反思环节",
    "3. 学习小组：4-6人组成，负责实际问题的研讨和解决",
    "4. 业务专家：提供专业支持和咨询"
)

foreach ($p in $participants) {
    $para = $doc.Paragraphs.Add()
    $para.Range.Text = $p
    $para.Range.Font.Size = 12
    $para.Range.ParagraphFormat.LineSpacing = 18
    $para.Range.InsertParagraphAfter()
}

$doc.Paragraphs.Add().Range.Text = ""

$section4 = $doc.Paragraphs.Add()
$section4.Range.Text = "四、学习议题"
$section4.Range.Font.Size = 16
$section4.Range.Font.Bold = $true
$section4.Range.InsertParagraphAfter()

$topics = @(
    "1. 当前业务瓶颈分析及解决方案",
    "2. 跨部门协作优化",
    "3. 客户满意度提升策略",
    "4. 创新思维培养",
    "5. 领导力发展实践"
)

foreach ($t in $topics) {
    $para = $doc.Paragraphs.Add()
    $para.Range.Text = $t
    $para.Range.Font.Size = 12
    $para.Range.ParagraphFormat.LineSpacing = 18
    $para.Range.InsertParagraphAfter()
}

$doc.Paragraphs.Add().Range.Text = ""

$section5 = $doc.Paragraphs.Add()
$section5.Range.Text = "五、实施流程"
$section5.Range.Font.Size = 16
$section5.Range.Font.Bold = $true
$section5.Range.InsertParagraphAfter()

$steps = @(
    "1. 问题识别：各小组识别并确定要解决的实际问题",
    "2. 方案设计：小组讨论并制定解决方案",
    "3. 行动实施：在实际行动中验证和完善方案",
    "4. 反思学习：定期回顾和反思学习成果",
    "5. 成果分享：向其他小组和领导层汇报成果"
)

foreach ($s in $steps) {
    $para = $doc.Paragraphs.Add()
    $para.Range.Text = $s
    $para.Range.Font.Size = 12
    $para.Range.ParagraphFormat.LineSpacing = 18
    $para.Range.InsertParagraphAfter()
}

$doc.Paragraphs.Add().Range.Text = ""

$section6 = $doc.Paragraphs.Add()
$section6.Range.Text = "六、时间安排"
$section6.Range.Font.Size = 16
$section6.Range.Font.Bold = $true
$section6.Range.InsertParagraphAfter()

$schedule = @(
    "第一阶段（第1周）：项目启动，组建学习小组，确定议题",
    "第二阶段（第2-3周）：问题分析和方案设计",
    "第三阶段（第4-6周）：方案实施和持续优化",
    "第四阶段（第7-8周）：反思总结和成果分享"
)

foreach ($s in $schedule) {
    $para = $doc.Paragraphs.Add()
    $para.Range.Text = $s
    $para.Range.Font.Size = 12
    $para.Range.ParagraphFormat.LineSpacing = 18
    $para.Range.InsertParagraphAfter()
}

$doc.Paragraphs.Add().Range.Text = ""

$section7 = $doc.Paragraphs.Add()
$section7.Range.Text = "七、预期成果"
$section7.Range.Font.Size = 16
$section7.Range.Font.Bold = $true
$section7.Range.InsertParagraphAfter()

$outcomes = @(
    "1. 完成2-3个关键业务问题的解决方案",
    "2. 提升团队成员的综合能力",
    "3. 建立行动学习的方法论和工具库",
    "4. 形成可持续的学习机制",
    "5. 产出可复用的行动学习案例集"
)

foreach ($o in $outcomes) {
    $para = $doc.Paragraphs.Add()
    $para.Range.Text = $o
    $para.Range.Font.Size = 12
    $para.Range.ParagraphFormat.LineSpacing = 18
    $para.Range.InsertParagraphAfter()
}

$doc.Paragraphs.Add().Range.Text = ""

$section8 = $doc.Paragraphs.Add()
$section8.Range.Text = "八、评估机制"
$section8.Range.Font.Size = 16
$section8.Range.Font.Bold = $true
$section8.Range.InsertParagraphAfter()

$evaluation = @(
    "1. 过程评估：定期检查各小组的进展情况",
    "2. 结果评估：评估问题解决的实际效果",
    "3. 学习评估：通过问卷和访谈评估学习成果",
    "4. 同伴评估：小组成员之间的互评",
    "5. 专家评估：邀请业务专家对成果进行评审"
)

foreach ($e in $evaluation) {
    $para = $doc.Paragraphs.Add()
    $para.Range.Text = $e
    $para.Range.Font.Size = 12
    $para.Range.ParagraphFormat.LineSpacing = 18
    $para.Range.InsertParagraphAfter()
}

$doc.Paragraphs.Add().Range.Text = ""

$section9 = $doc.Paragraphs.Add()
$section9.Range.Text = "九、风险控制"
$section9.Range.Font.Size = 16
$section9.Range.Font.Bold = $true
$section9.Range.InsertParagraphAfter()

$risks = @(
    "1. 成员参与度不足风险：建立激励机制，定期沟通",
    "2. 议题选择不当风险：教练引导，确保议题的实用性",
    "3. 资源投入不足风险：提前规划和争取必要资源",
    "4. 效果不明显风险：设置阶段性目标，及时调整方向",
    "5. 知识流失风险：建立知识管理体系，做好文档归档"
)

foreach ($r in $risks) {
    $para = $doc.Paragraphs.Add()
    $para.Range.Text = $r
    $para.Range.Font.Size = 12
    $para.Range.ParagraphFormat.LineSpacing = 18
    $para.Range.InsertParagraphAfter()
}

$doc.Paragraphs.Add().Range.Text = ""

$section10 = $doc.Paragraphs.Add()
$section10.Range.Text = "十、资源配置"
$section10.Range.Font.Size = 16
$section10.Range.Font.Bold = $true
$section10.Range.InsertParagraphAfter()

$resources = @(
    "1. 人力资源：教练、业务专家、学习小组成员",
    "2. 时间资源：每周至少4小时的学习和讨论时间",
    "3. 场地资源：会议室、学习空间",
    "4. 技术资源：在线协作工具、学习平台",
    "5. 预算资源：培训费用、材料费用、奖励费用"
)

foreach ($r in $resources) {
    $para = $doc.Paragraphs.Add()
    $para.Range.Text = $r
    $para.Range.Font.Size = 12
    $para.Range.ParagraphFormat.LineSpacing = 18
    $para.Range.InsertParagraphAfter()
}

$doc.Save()
$doc.Close()
$wordApp.Quit()

Write-Host "Action learning plan has been written to the document!"
