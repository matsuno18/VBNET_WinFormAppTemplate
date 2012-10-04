Public Class Form_A

#Region "サンプル実装"
    ''' <summary>
    ''' 他スレッドからメッセージボックスを表示させる
    ''' </summary>
    ''' <param name="message_text"></param>
    ''' <param name="dialog_caption"></param>
    ''' <remarks>他スレッドからGUIを操作する方法のサンプル実装です</remarks>
    Public Sub MsgBoxShow(ByVal message_text As String, ByVal dialog_caption As String)
        Invoke(New Action(Of String, String)(AddressOf _MsgBoxShow), message_text, dialog_caption)
    End Sub
    Private Sub _MsgBoxShow(ByVal message_text As String, ByVal dialog_caption As String)
        MsgBox(message_text, MsgBoxStyle.Information, dialog_caption)
    End Sub
#End Region

    Dim name_id As Integer = 0 ' ワーカーのID(連番)

    ' メッセージテキストを１行追加する
    Public Sub MessageLine(ByVal message_text As String)
        Invoke(New Action(Of String)(AddressOf _MessageLine), message_text)
    End Sub
    Private Sub _MessageLine(ByVal message_text As String)
        TextBox1.AppendText(message_text & vbCrLf)
    End Sub

    Private Sub Form_A_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ' Exeファイル名をフォームのタイトル文字列の先頭に付加する
        Dim ExeName As String = System.IO.Path.GetFileNameWithoutExtension(Application.ExecutablePath)
        Me.Text = ExeName & " " & Me.Text
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Worker.StartWorker(Me, CStr(name_id))
        name_id += 1 ' ワーカーを作るたびに増やしていく
    End Sub


End Class