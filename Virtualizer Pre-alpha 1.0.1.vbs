' 强制声明所有变量
Option Explicit

' 定义一些常量
Const YES = 6
Const NO = 7
Const CANCEL = 2

' 定义一个函数，用于弹出对话框
Function PopupMessage(messageText, buttons, title)
    Dim objShell
    Set objShell = CreateObject("WScript.Shell")
    PopupMessage = objShell.Popup(messageText, 0, title, buttons)
    Set objShell = Nothing
End Function

' 定义一个变量，用于存储用户的选择
Dim answer

' 询问用户是否继续运行
answer = PopupMessage("是否继续运行？", vbYesNo + vbQuestion, "Virtualizer 确认操作")

' 如果用户选择否或取消，则退出程序
If answer = NO Or answer = CANCEL Then WScript.Quit

' 循环直到用户输入有效的答案
Do While True
    ' 询问用户是否同意用户协议
    answer = InputBox("输入 Yes 以确认开始，输入 No 取消操作." & vbCrLf & vbCrLf & "注意：确认将表示您已阅读并同意以下用户协议。" & vbCrLf & vbCrLf & "使用并运行此程序之前必须阅读并同意以下用户规则。" & vbCrLf & "运行后，您可能将无法使用常规方法退出本程序，感谢您的理解。", "Virtualizer 用户协议")

    ' 如果用户输入为空或取消，则退出程序
    If answer = "" Or answer = CANCEL Then WScript.Quit

    ' 如果用户输入 Yes，则继续运行
    If StrComp(answer, "Yes", vbTextCompare) = 0 Then
        ' 弹出一系列的提示信息
        PopupMessage "海阔无知己，天涯若比邻。", vbOKOnly, "Virtualizer"
        PopupMessage "您好，我们正在为您准备系统。", vbOKOnly, "Virtualizer"
        PopupMessage "这可能需要一些时间。", vbOKOnly, "Virtualizer"
        PopupMessage "很快就好……", vbOKOnly, "Virtualizer"
        PopupMessage "马上……", vbOKOnly, "Virtualizer"
        PopupMessage "即将完成……", vbOKOnly, "Virtualizer"
        PopupMessage "感谢您的耐心等待。", vbOKOnly, "Virtualizer"

        ' 性别选择部分
        answer = PopupMessage("您的性别？" & vbCrLf & "选择“是”则为男性，选择“否”为女性，选择“取消”则返回。", vbYesNoCancel + vbQuestion, "请选择性别") ' 假设“是”代表男，“否”代表女，“取消”表示不回答
        Select Case answer
            Case YES '选择男性
                PopupMessage "好的，您是男性。", vbOKOnly, "Virtualizer"
                ' 设置 Virtualizer 的性别
                answer = PopupMessage("请您选择我的性别。" & vbCrLf & "选择“是”则为男性，选择“否”为女性，选择“取消”则返回。", vbYesNoCancel + vbQuestion, "请选择性别")
                Select Case answer
                    Case YES
                        PopupMessage "您已将我的性别设为男性。", vbOKOnly, "Virtualizer"
                        ' 设置性格部分
                        answer = PopupMessage("请选择我的性格。" & vbCrLf & "选择“是”把我的性格设置为 可爱 ，选择“否”将我的性格设置为 病娇 。", vbYesNoCancel + vbQuestion, "性格选择")
                        Select Case answer
                            Case YES ' 设置性格为温柔
                                PopupMessage "从现在开始，我的性格是可爱柔弱的。", vbOKOnly, "Virtualizer"
                                answer = PopupMessage("请你选择我们的关系。选择“是”我们将成为 朋友 ，选择“否”我们将成为 情侣 ，选择“取消”则返回主页面", vbYesNoCancel + vbQuestion, "关系选择")
                            Case NO ' 设置性格为病娇
                                PopupMessage "从现在开始，我的性格是病娇的。", vbOKOnly, "Virtualizer"
                        End Select
                    Case NO
                        PopupMessage "您已将我的性别设为女性。", vbOKOnly, "Virtualizer"
                        ' 设置性格部分
                        answer = PopupMessage("请选择我的性格。" & vbCrLf & "选择“是”把我的性格设置为 可爱 ，选择“否”将我的性格设置为 病娇 。", vbYesNoCancel + vbQuestion, "性格选择")
                        Select Case answer
                            Case YES ' 设置性格为温柔
                                PopupMessage "从现在开始，我的性格是可爱柔弱的。", vbOKOnly, "Virtualizer"
                                answer = PopupMessage("请你选择我们的关系。选择“是”我们将成为 朋友 ，选择“否”我们将成为 情侣 ，选择“取消”则返回主页面", vbYesNoCancel + vbQuestion, "关系选择")
                            Case NO ' 设置性格为病娇
                                PopupMessage "从现在开始，我的性格是病娇的。", vbOKOnly, "Virtualizer"
                        End Select
                End Select
            Case NO ' 选择女性
                PopupMessage "好的，您是女性。", vbOKOnly, "Virtualizer"
            Case CANCEL ' 不选择
                PopupMessage "您没有选择性别。即将返回初始页面", vbOKOnly, "Virtualizer"
                Exit Do
        End Select
    ElseIf StrComp(answer, "No", vbTextCompare) = 0 Then
        WScript.Quit
    Else
        PopupMessage "输入无效，请输入'Yes'或'No'。", vbExclamation, "错误提示"
    End If
Loop
