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

    Private Sub Form_A_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ' Exeファイル名をフォームのタイトル文字列の先頭に付加する
        Dim ExeName As String = System.IO.Path.GetFileNameWithoutExtension(Application.ExecutablePath)
        Me.Text = ExeName & " " & Me.Text
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        MsgBox("てすと")
        Throw New Exception()
    End Sub


End Class