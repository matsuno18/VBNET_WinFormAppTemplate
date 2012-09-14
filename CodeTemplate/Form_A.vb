Public Class Form_A

#Region "サンプル実装"
    ''' <summary>
    ''' 他スレッドからメッセージボックスを表示させる
    ''' </summary>
    ''' <param name="message_text"></param>
    ''' <param name="dialog_caption"></param>
    ''' <remarks>他スレッドからGUIを操作する方法のサンプル実装です</remarks>
    Public Sub ShowMsgBox(ByVal message_text As String, ByVal dialog_caption As String)
        Invoke(New Action(Of String, String)(AddressOf _ShowMsgBox), message_text, dialog_caption)
    End Sub
    Private Sub _ShowMsgBox(ByVal message_text As String, ByVal dialog_caption As String)
        MsgBox(message_text, MsgBoxStyle.Information, dialog_caption)
    End Sub
#End Region

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        MsgBox("てすと")
    End Sub
End Class