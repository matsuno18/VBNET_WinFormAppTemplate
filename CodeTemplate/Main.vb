
''' <summary>
''' プログラムのメイン部分
''' </summary>
''' <remarks>
''' Baseモジュールと結びついている。
''' エントリポイントはBase.Main()
''' </remarks>
Module Main

    ' アプリケーション開始(Base.Main()から呼ばれる)
    Public Sub Application_Startup()

        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(New Form_A())

    End Sub

    ' プログラムがエラー終了するときの処理
    Public Sub Application_Abort(ByVal sender As Object, ByVal e As Object)

        Dim ex As Exception = Nothing

        Dim ex_Thread As System.Threading.ThreadExceptionEventArgs
        Dim ex_Unhandled As System.UnhandledExceptionEventArgs

        ' それぞれの型へのキャストを試みる
        ex_Thread = TryCast(e, System.Threading.ThreadExceptionEventArgs)
        ex_Unhandled = TryCast(e, System.UnhandledExceptionEventArgs)

        If Not ex_Thread Is Nothing Then
            'MsgBox("メインアプリケーションスレッドで未捕捉の例外が発生しました")
            ex = ex_Thread.Exception
        End If

        If Not ex_Unhandled Is Nothing Then
            'MsgBox("未捕捉の例外が発生しました")
            ex = TryCast(ex_Unhandled.ExceptionObject, Exception)
        End If

        ' 未捕捉例外の処理を行う
        UnhandledExceptionProcessing(ex)

        ' 2重起動防止用のMutexを解放する
        Release_2ndRunMutex()

        ' プログラムを終了する
        System.Environment.Exit(-1)

    End Sub

    ' 未捕捉例外の処理を行う
    Private Sub UnhandledExceptionProcessing(ByVal ex As Exception)

        If ex Is Nothing Then
            Return
        End If

        ' スタックトレースを表示する
        MsgBox(ex.ToString())

    End Sub

End Module
