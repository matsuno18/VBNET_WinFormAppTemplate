Module Base

    Public App As AppBase ' VBアプリケーションモデルのオブジェクト

    ' プログラムのエントリポイント
    <STAThread()> _
    Sub Main(ByVal CmdArgs() As String)

        ' プログラムの初期例外処理の設定を行う
        InitialExceptionHandlingSetup()

        App = New AppBase ' VBアプリケーションモデルの初期化を行う
        App.Run(CmdArgs) ' アプリケーションを開始する

    End Sub

    ' プログラムの初期例外処理の設定を行う
    Private Sub InitialExceptionHandlingSetup()

        ' Application.ThreadException も UnhandledException で処理するようにモードを変更する
        Windows.Forms.Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException)

        ' メインアプリケーションスレッドの例外集約ハンドラを登録する(上のモード変更のため，これは使われないはず)(サンプルとして残しておく)
        AddHandler Windows.Forms.Application.ThreadException, AddressOf Application_ThreadException

        ' アプリケーションドメインの例外集約ハンドラを登録する
        AddHandler System.Threading.Thread.GetDomain.UnhandledException, AddressOf Application_UnhandledException

    End Sub

    ' メインアプリケーションスレッドの未捕捉例外処理
    Private Sub Application_ThreadException(ByVal sender As Object, ByVal e As System.Threading.ThreadExceptionEventArgs)
        MsgBox("メインアプリケーションスレッドで未捕捉の例外が発生しました")
        System.Environment.Exit(-1)
    End Sub

    ' アプリケーションドメインの未捕捉例外処理
    Private Sub Application_UnhandledException(ByVal sender As Object, ByVal e As System.UnhandledExceptionEventArgs)
        MsgBox("未捕捉のエラーが発生しました")
        System.Environment.Exit(-1)
    End Sub

End Module

''' <summary>
''' 
''' </summary>
''' <remarks>
''' Visual Basic アプリケーションモデルを利用するために用意されたクラス
''' </remarks>
Class AppBase
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

        ' 通常起動時に呼ばれるイベントを設定する
        AddHandler Me.Startup, AddressOf MyAppBase_Startup

        ' このアプリケーションのメインとなるFormのインスタンスを設定する
        Me.MainForm = New Form()

    End Sub

    ' 通常起動時に呼ばれる処理
    Private Sub MyAppBase_Startup(ByVal sender As Object, ByVal e As Microsoft.VisualBasic.ApplicationServices.StartupEventArgs)

        ' INIファイルの読み込みなどを記述する
        ' e.Cancel = True で起動をキャンセルすることが可能

    End Sub

    ' ２重起動されたときに呼ばれる処理
    Private Sub MyAppBase_StartupNextInstance(ByVal sender As Object, ByVal e As Microsoft.VisualBasic.ApplicationServices.StartupNextInstanceEventArgs)
        ''[処理内容 概要]
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