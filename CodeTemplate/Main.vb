Module Main

    Sub Main(ByVal CmdArgs() As String)

        Dim myApp As New MyAppBase
        myApp.Run(CmdArgs)

    End Sub

End Module

''' <summary>
''' 
''' </summary>
''' <remarks>
''' Visual Basic アプリケーションモデルを利用するために用意されたクラス
''' </remarks>
Class MyAppBase
    Inherits Microsoft.VisualBasic.ApplicationServices.WindowsFormsApplicationBase

    Sub New()
        '' VBアプリケーションモデルの初期化・設定をする
        ''=====================================================================

        ' ビジュアルスタイルを有効にする
        Me.EnableVisualStyles = True

        ' 単一インスタンスのアプリケーションとする
        Me.IsSingleInstance = True

        ' 上記と関連して，２重起動されたときに呼ばれるイベントを設定する
        AddHandler Me.StartupNextInstance, AddressOf MyAppBase_StartupNextInstance

        ' このアプリケーションのメインとなるFormのインスタンスを設定する
        Me.MainForm = New Form_A()

    End Sub

    ' ２重起動されたときに呼ばれる処理
    Private Sub MyAppBase_StartupNextInstance(ByVal sender As Object, ByVal e As Microsoft.VisualBasic.ApplicationServices.StartupNextInstanceEventArgs)
        ''[概要]
        '' ウィンドウをアクティブにして，さらに２重起動されたことを通知する
        ''=====================================================================

        ' ウィンドウのZオーダーを最前面となるように設定する
        Me.MainForm.BringToFront()

        ' アプリケーションをアクティブにする(アクティブにする対象は自分自身)
        AppActivate(Process.GetCurrentProcess.Id)

        ' このアプリケーションのメインフォームをアクティブにする
        Me.MainForm.Activate()

        ' メッセージを表示する
        MsgBox("２重起動されました！", MsgBoxStyle.Information, "通知")

    End Sub

End Class