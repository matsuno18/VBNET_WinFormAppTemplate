Option Explicit On
Option Strict On
Option Compare Binary
Option Infer On

''' <summary>
''' プログラムの基礎部分
''' </summary>
''' <remarks>
''' Mainモジュールと結びついている。
''' mutex_2ndRun_uuidはプログラム毎に書き換える必要がある。
''' </remarks>
Module Base

    ' 2重起動防止用Mutexを一意に識別するためのGUID(他のアプリケーションで使い回さないこと！)
    Const mutex_2ndRun_uuid As String = "{0DBF1626-EE07-49EF-B356-3799FF1FF2DF}"

    ' プログラムのエントリポイント
    <STAThread()> _
    Sub Main(ByVal CmdArgs() As String)

        ' 未捕捉例外処理ハンドラの初期設定を行う
        InitialExceptionHandlingSetup()

        ' 2重起動防止用のMutexを初期化して，2重起動チェックを行う
        InitAndCheck_2ndRunMutex() ' 2重起動されていたら，ここで終了する

        ' アプリケーションを開始する
        Application_Startup()

        ' 2重起動用Mutexを解放する
        Release_2ndRunMutex()

    End Sub

#Region "未捕捉例外 関係"

    ' 未捕捉例外処理ハンドラの初期設定を行う
    Public Sub InitialExceptionHandlingSetup()

        ' メインのUIスレッドで発生した全ての例外に対するイベントハンドラを追加する()
        AddHandler Application.ThreadException, AddressOf Application_ThreadException

        ' アプリケーションドメイン内のメインのUIスレッドを除くすべてのスレッドに対してイベントハンドラを追加する
        AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf Application_UnhandledException

    End Sub

    ' UIスレッドで発生する例外を処理する
    Private Sub Application_ThreadException(ByVal sender As Object, ByVal e As System.Threading.ThreadExceptionEventArgs)
        Application_Abort(sender, e)
    End Sub

    ' UIスレッド以外の例外を処理する
    Private Sub Application_UnhandledException(ByVal sender As Object, ByVal e As UnhandledExceptionEventArgs)
        Application_Abort(sender, e)
    End Sub

#End Region

#Region "2重起動防止 関係"

    ' 2重起動防止に使うMutex
    Private mutex_2ndRun As System.Threading.Mutex

    ''' <summary>
    ''' 2重起動防止用のMutexを初期化して，2重起動チェックを行う
    ''' </summary>
    ''' <remarks>2重起動されていたら，このプログラムは終了する</remarks>
    Public Sub InitAndCheck_2ndRunMutex()

        ' 所有権を持たない状態でMutexを作成する
        mutex_2ndRun = New System.Threading.Mutex(False, mutex_2ndRun_uuid)

        ' MutexオブジェクトがGCによって勝手に解放されないようにする
        System.GC.KeepAlive(mutex_2ndRun)

        ' Mutexの所有権を要求する
        If Not mutex_2ndRun.WaitOne(0, False) Then
            ' 所有権を取得できなかった，つまり既に起動されているので終了する
            MsgBox("既に起動されています")
            mutex_2ndRun.Close()
            System.Environment.Exit(1)
        End If

    End Sub

    ' 2重起動防止用のMutexを解放する
    Public Sub Release_2ndRunMutex()
        mutex_2ndRun.ReleaseMutex() ' 所有権を手放す
        mutex_2ndRun.Close() ' Mutexを解放する
    End Sub

#End Region

End Module

