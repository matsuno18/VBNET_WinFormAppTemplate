﻿Public Class Worker

    ' ワーカースレッドを生成して開始する
    Public Shared Function StartWorker(ByVal frm As Form_A, ByVal name As String) As System.Threading.Thread

        Dim worker As System.Threading.Thread
        Dim pgm As New Worker(frm, name)

        worker = New System.Threading.Thread(AddressOf pgm.Startup)
        worker.Start()

        Return worker

    End Function

    Private uiform As Form_A ' ワーカーと紐づいているフォーム
    Private worker_name As String ' ワーカースレッドの識別用名称

    ' コンストラクタ
    Sub New(ByVal uiform As Form_A, ByVal name As String)
        Me.uiform = uiform
        Me.worker_name = name
    End Sub

    ' ワーカーを開始する
    Public Sub Startup()

        Init()
        Main()
        Shutdown()

    End Sub

    ' 初期化
    Private Sub Init()

    End Sub

    ' メイン処理
    Private Sub Main()

        ' 100までカウントアップ。それを画面上で逐次表示する
        For i = 0 To 100
            uiform.MessageLine(worker_name & " >> " & CStr(i))
            System.Threading.Thread.Sleep(50)
        Next

    End Sub

    ' 終了処理
    Private Sub Shutdown()

    End Sub

End Class
