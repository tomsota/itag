'---------------------------------------------
'[Name]     cls99ErrCtrl
'[J.Name]   エラー操作クラス
'[Date]     2010/10/20
'[Func]     エラー操作
'[Memo]
'---------------------------------------------
Public Class cls99ErrCtrl
    '---------------------------------------------
    ' メンバ変数宣言
    '---------------------------------------------
    Public Structure strctErrData
        Public strErrMod As String
        Public strErrNum As Integer
        Public strErrDesc As String
    End Structure

    Dim m_strctErrData As strctErrData
    '---------------------------------------------
    ' プロパティ
    '---------------------------------------------
    Public WriteOnly Property ErrMod As String
        Set(ByVal value As String)
            m_strctErrData.strErrMod = value
        End Set
    End Property

    Public WriteOnly Property ErrDesc As String
        Set(ByVal value As String)
            m_strctErrData.strErrDesc = value
        End Set
    End Property

    Public WriteOnly Property ErrNum As Integer
        Set(ByVal value As Integer)
            m_strctErrData.strErrNum = value
        End Set
    End Property
    '---------------------------------------------
    '[Name]     New
    '[J.Name]   クラス生成時処理
    '[Date]     2010/10/20
    '[Func]     クラス初期化
    '[Input]    なし
    '[Output]   なし
    '[Memo]
    '---------------------------------------------
    Public Sub New()
        '---------------------------------------------
        ' ローカル変数宣言
        '---------------------------------------------
        '---------------------------------------------
        ' 関数メイン
        '---------------------------------------------
        Try
            '---------------------------------------------
            ' 初期処理
            '---------------------------------------------
            '●デバッグログ出力
            'Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name.ToString & "-Start!")
            '---------------------------------------------
            ' メイン処理
            '---------------------------------------------
            '●エラー構造体の初期化
            With m_strctErrData
                .strErrMod = ""
                .strErrDesc = ""
                .strErrNum = 0
            End With

        Catch ex As Exception
            '---------------------------------------------
            ' 例外処理
            '---------------------------------------------
        Finally
            '---------------------------------------------
            ' 終了処理
            '---------------------------------------------
            '●デバッグログ出力
            'Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name.ToString & "-End!")
        End Try
    End Sub

    '---------------------------------------------
    '[Name]     LogWrite
    '[J.Name]   ログ出力処理
    '[Date]     2010/10/20
    '[Func]     エラーをファイルに出力する
    '[Input]    なし
    '[Output]   lngRet as long
    '[Memo]
    '---------------------------------------------
    Public Function LogWrite() As Long
        '---------------------------------------------
        ' ローカル変数宣言
        '---------------------------------------------
        Dim lngRet As Long
        Dim objWriter As IO.StreamWriter
        '---------------------------------------------
        ' 関数メイン
        '---------------------------------------------
        Try
            '---------------------------------------------
            ' 初期処理
            '---------------------------------------------
            '●デバッグログ出力
            Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name & "-Start!")

            '●戻り値初期化
            lngRet = eReturnCode.OK

            '●Writerクラスインスタンス化
            objWriter = New IO.StreamWriter(My.Application.Info.DirectoryPath & "\Err.txt", True)
            '---------------------------------------------
            ' メイン処理
            '---------------------------------------------
            '●エラー情報を格納する
            With m_strctErrData
                objWriter.WriteLine("◆エラー内容：" & _
                                    "[Time]" & Now.ToString & _
                                    "[Mod ]" & .strErrMod.PadRight(40) & _
                                    "[Num ]" & .strErrNum.ToString.PadRight(4) & _
                                    "[Desc]" & .strErrDesc)
            End With

            '●ファイルを閉じる
            objWriter.Close()
        Catch ex As Exception
            '---------------------------------------------
            ' 例外処理
            '---------------------------------------------
            '●戻り値にエラーを格納する
            lngRet = eReturnCode.NG
        Finally
            '---------------------------------------------
            ' 終了処理
            '---------------------------------------------
            '●デバッグログ出力
            Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name & "-End!")

            '●オブジェクトの破棄
            objWriter = Nothing

            '●戻り値を返す
            LogWrite = lngRet

        End Try
    End Function
End Class
