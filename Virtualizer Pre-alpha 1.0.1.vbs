' ǿ���������б���
Option Explicit

' ����һЩ����
Const YES = 6
Const NO = 7
Const CANCEL = 2

' ����һ�����������ڵ����Ի���
Function PopupMessage(messageText, buttons, title)
    Dim objShell
    Set objShell = CreateObject("WScript.Shell")
    PopupMessage = objShell.Popup(messageText, 0, title, buttons)
    Set objShell = Nothing
End Function

' ����һ�����������ڴ洢�û���ѡ��
Dim answer

' ѯ���û��Ƿ��������
answer = PopupMessage("�Ƿ�������У�", vbYesNo + vbQuestion, "Virtualizer ȷ�ϲ���")

' ����û�ѡ����ȡ�������˳�����
If answer = NO Or answer = CANCEL Then WScript.Quit

' ѭ��ֱ���û�������Ч�Ĵ�
Do While True
    ' ѯ���û��Ƿ�ͬ���û�Э��
    answer = InputBox("���� Yes ��ȷ�Ͽ�ʼ������ No ȡ������." & vbCrLf & vbCrLf & "ע�⣺ȷ�Ͻ���ʾ�����Ķ���ͬ�������û�Э�顣" & vbCrLf & vbCrLf & "ʹ�ò����д˳���֮ǰ�����Ķ���ͬ�������û�����" & vbCrLf & "���к������ܽ��޷�ʹ�ó��淽���˳������򣬸�л������⡣", "Virtualizer �û�Э��")

    ' ����û�����Ϊ�ջ�ȡ�������˳�����
    If answer = "" Or answer = CANCEL Then WScript.Quit

    ' ����û����� Yes�����������
    If StrComp(answer, "Yes", vbTextCompare) = 0 Then
        ' ����һϵ�е���ʾ��Ϣ
        PopupMessage "������֪�������������ڡ�", vbOKOnly, "Virtualizer"
        PopupMessage "���ã���������Ϊ��׼��ϵͳ��", vbOKOnly, "Virtualizer"
        PopupMessage "�������ҪһЩʱ�䡣", vbOKOnly, "Virtualizer"
        PopupMessage "�ܿ�ͺá���", vbOKOnly, "Virtualizer"
        PopupMessage "���ϡ���", vbOKOnly, "Virtualizer"
        PopupMessage "������ɡ���", vbOKOnly, "Virtualizer"
        PopupMessage "��л�������ĵȴ���", vbOKOnly, "Virtualizer"

        ' �Ա�ѡ�񲿷�
        answer = PopupMessage("�����Ա�" & vbCrLf & "ѡ���ǡ���Ϊ���ԣ�ѡ�񡰷�ΪŮ�ԣ�ѡ��ȡ�����򷵻ء�", vbYesNoCancel + vbQuestion, "��ѡ���Ա�") ' ���衰�ǡ������У����񡱴���Ů����ȡ������ʾ���ش�
        Select Case answer
            Case YES 'ѡ������
                PopupMessage "�õģ��������ԡ�", vbOKOnly, "Virtualizer"
                ' ���� Virtualizer ���Ա�
                answer = PopupMessage("����ѡ���ҵ��Ա�" & vbCrLf & "ѡ���ǡ���Ϊ���ԣ�ѡ�񡰷�ΪŮ�ԣ�ѡ��ȡ�����򷵻ء�", vbYesNoCancel + vbQuestion, "��ѡ���Ա�")
                Select Case answer
                    Case YES
                        PopupMessage "���ѽ��ҵ��Ա���Ϊ���ԡ�", vbOKOnly, "Virtualizer"
                        ' �����Ը񲿷�
                        answer = PopupMessage("��ѡ���ҵ��Ը�" & vbCrLf & "ѡ���ǡ����ҵ��Ը�����Ϊ �ɰ� ��ѡ�񡰷񡱽��ҵ��Ը�����Ϊ ���� ��", vbYesNoCancel + vbQuestion, "�Ը�ѡ��")
                        Select Case answer
                            Case YES ' �����Ը�Ϊ����
                                PopupMessage "�����ڿ�ʼ���ҵ��Ը��ǿɰ������ġ�", vbOKOnly, "Virtualizer"
                                answer = PopupMessage("����ѡ�����ǵĹ�ϵ��ѡ���ǡ����ǽ���Ϊ ���� ��ѡ�񡰷����ǽ���Ϊ ���� ��ѡ��ȡ�����򷵻���ҳ��", vbYesNoCancel + vbQuestion, "��ϵѡ��")
                            Case NO ' �����Ը�Ϊ����
                                PopupMessage "�����ڿ�ʼ���ҵ��Ը��ǲ����ġ�", vbOKOnly, "Virtualizer"
                        End Select
                    Case NO
                        PopupMessage "���ѽ��ҵ��Ա���ΪŮ�ԡ�", vbOKOnly, "Virtualizer"
                        ' �����Ը񲿷�
                        answer = PopupMessage("��ѡ���ҵ��Ը�" & vbCrLf & "ѡ���ǡ����ҵ��Ը�����Ϊ �ɰ� ��ѡ�񡰷񡱽��ҵ��Ը�����Ϊ ���� ��", vbYesNoCancel + vbQuestion, "�Ը�ѡ��")
                        Select Case answer
                            Case YES ' �����Ը�Ϊ����
                                PopupMessage "�����ڿ�ʼ���ҵ��Ը��ǿɰ������ġ�", vbOKOnly, "Virtualizer"
                                answer = PopupMessage("����ѡ�����ǵĹ�ϵ��ѡ���ǡ����ǽ���Ϊ ���� ��ѡ�񡰷����ǽ���Ϊ ���� ��ѡ��ȡ�����򷵻���ҳ��", vbYesNoCancel + vbQuestion, "��ϵѡ��")
                            Case NO ' �����Ը�Ϊ����
                                PopupMessage "�����ڿ�ʼ���ҵ��Ը��ǲ����ġ�", vbOKOnly, "Virtualizer"
                        End Select
                End Select
            Case NO ' ѡ��Ů��
                PopupMessage "�õģ�����Ů�ԡ�", vbOKOnly, "Virtualizer"
            Case CANCEL ' ��ѡ��
                PopupMessage "��û��ѡ���Ա𡣼������س�ʼҳ��", vbOKOnly, "Virtualizer"
                Exit Do
        End Select
    ElseIf StrComp(answer, "No", vbTextCompare) = 0 Then
        WScript.Quit
    Else
        PopupMessage "������Ч��������'Yes'��'No'��", vbExclamation, "������ʾ"
    End If
Loop
