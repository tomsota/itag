Imports iTunesLib
'---------------------------------------------
'[Name]     cls01frmToolMain
'[J.Name]   フォーム操作クラス
'[Date]     2010/10/21
'[Func]     フォーム操作クラス
'[Memo]
'---------------------------------------------
Public Class cls01frmToolMain
    '---------------------------------------------
    ' 構造体宣言
    '---------------------------------------------
    '●カレントトラック
    Private Structure strctCurrentTrack
        Dim blnCurrentOn As Boolean
        Dim strTitle As String
        Dim strArtist As String
        Dim strAlbum As String
        Dim intRate As Integer
        Dim strComment As String
        Dim objArtWork As iTunesLib.IITArtworkCollection
    End Structure
    '---------------------------------------------
    ' メンバ変数宣言
    '---------------------------------------------
    Private m_blnChanged As Boolean                     '変更済みフラグ
    Private WithEvents m_objITunes As iTunesApp         'iTunesApp
    Private m_intIdxSelectPL As Integer                 '選択中のプレイリスト番号
    Private m_intIdxSelectSong As Integer               '選択中のソング番号
    Private m_strctCurrentTrack As strctCurrentTrack    'カレントトラック情報
    Private m_intPlayStatus As Integer                  '現在のプレイヤーステータス
    Private m_blnMute As Boolean                        'ミュート状態
    Private m_intVolume As Integer                      'ボリューム
    Private m_strMakeComment As String                  '編集中のコメント
    Private m_intArtWorkNo As Integer                   '現在選択中のアートワーク番号
    Private m_objLvItm() As ListViewItem
    Private m_objLvItmPL() As Collection
    '---------------------------------------------
    ' プロパティ
    '---------------------------------------------
    '----------------------------------------------
    ' イベントの定義
    '----------------------------------------------
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
            Debug.Print("cls01frmToolMain New -Start!")
            '---------------------------------------------
            ' メイン処理
            '---------------------------------------------
            '●iTunes関連づけ
            m_objITunes = New iTunesApp

            Debug.Print("★★iTunes App Connect★★")

            '●データ部の初期化
            p_objMySetting(0) = My.Settings.strArrayData01
            p_objMySetting(1) = My.Settings.strArrayData02
            p_objMySetting(2) = My.Settings.strArrayData03
            p_objMySetting(3) = My.Settings.strArrayData04
            p_objMySetting(4) = My.Settings.strArrayData05
            p_objMySetting(5) = My.Settings.strArrayData06
            p_objMySetting(6) = My.Settings.strArrayData07
            p_objMySetting(7) = My.Settings.strArrayData08
            p_objMySetting(8) = My.Settings.strArrayData09
            p_objMySetting(9) = My.Settings.strArrayData10

            '●タブ数
            ReDim p_strctData(My.Settings.strArrayTabName.Count - 1)

            '●イニシャライズをコール
            If F00Initialize() = eReturnCode.NG Then
                Throw New ApplicationException("F00Initializeの戻り値が不正です")
            End If

            '●カレントデータ初期化
            With m_strctCurrentTrack
                .intRate = 0
                .strAlbum = ""
                .strArtist = ""
                .strComment = ""
                .strTitle = ""
            End With

            ''●ライブラリデータ読み込み
            'ReDim m_objLvItm(m_objITunes.LibrarySource.Playlists.Item(1).Tracks.Count - 1)
            'Dim objTracks As IITTrackCollection
            'objTracks = m_objITunes.LibrarySource.Playlists.Item(1).Tracks
            'For i As Integer = 1 To m_objLvItm.Length
            '    m_objLvItm(i - 1) = New ListViewItem
            '    With m_objLvItm(i - 1)
            '        .Text = i
            '        .SubItems.Add(objTracks.Item(i).Name)
            '        If i Mod 100 = 0 Then
            '            Debug.Print(i)
            '        End If
            '    End With
            'Next

        Catch ex As Exception
            '---------------------------------------------
            ' 例外処理
            '---------------------------------------------
            If p_objErrCtrl IsNot Nothing Then

                '●エラー情報の取得
                With p_objErrCtrl
                    .ErrMod = "cls01frmToolMain"
                    .ErrNum = Err.Number
                    .ErrDesc = ex.Message.ToString
                End With

                '●エラー情報出力
                If p_objErrCtrl.LogWrite() = eReturnCode.NG Then
                    MessageBox.Show("エラーログ出力に失敗しました。", "ログ出力エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                End If
            Else
                '※エラー制御クラスがいない場合は終了する
                MessageBox.Show("想定外のエラーが発生しました。終了します", "エラー発生", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                frm01ToolMain.Close()
            End If

            Throw New ApplicationException("cls01ToolMainの生成に失敗しました")
        Finally
            '---------------------------------------------
            ' 終了処理
            '---------------------------------------------
            '●デバッグログ出力
            Debug.Print("cls01frmToolMain New -End!")
        End Try
    End Sub
    '---------------------------------------------
    '[Name]     Finalize
    '[J.Name]   クラス終了処理
    '[Date]     2010/10/31
    '[Func]     クラス破棄時処理
    '[Input]    なし
    '[Output]   なし
    '[Memo]
    '---------------------------------------------
    Protected Overrides Sub Finalize()
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
            '●iTunes関連づけ
            m_objITunes = Nothing

            Debug.Print("★★iTunes App DisConnect★★")

        Catch ex As Exception
            '---------------------------------------------
            ' 例外処理
            '---------------------------------------------
            If p_objErrCtrl IsNot Nothing Then
                '●エラー情報の取得
                With p_objErrCtrl
                    .ErrMod = Reflection.MethodBase.GetCurrentMethod.Name
                    .ErrNum = Err.Number
                    .ErrDesc = ex.Message.ToString
                End With

                '●エラー情報出力
                If p_objErrCtrl.LogWrite() = eReturnCode.NG Then
                    MessageBox.Show("エラーログ出力に失敗しました。", "ログ出力エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                End If
            Else
                '※エラー制御クラスがいない場合は終了する
                MessageBox.Show("想定外のエラーが発生しました。終了します", "エラー発生", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                frm01ToolMain.Close()
            End If

        Finally
            '---------------------------------------------
            ' 終了処理
            '---------------------------------------------
            MyBase.Finalize()

            '●デバッグログ出力
            'Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name.ToString & "-End!")
        End Try
    End Sub
    '---------------------------------------------
    '[Name]     objBtnPlay_Click
    '[J.Name]   プレイボタンクリック
    '[Date]     2010/10/31
    '[Func]     プレイボタンクリック
    '[Input]    なし
    '[Output]   lngRet as long          
    '[Memo]
    '---------------------------------------------
    Public Function objBtnPlay_Click() As Long
        '---------------------------------------------
        ' ローカル変数宣言
        '---------------------------------------------
        Dim lngRet As Long
        '---------------------------------------------
        ' 関数メイン
        '---------------------------------------------
        Try
            '---------------------------------------------
            ' 初期処理
            '---------------------------------------------
            '●デバッグログ出力
            Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name.ToString & "-Start!")

            '●戻り値初期化
            lngRet = eReturnCode.OK
            '---------------------------------------------
            ' メイン処理
            '---------------------------------------------
            '●再生or停止
            Select Case m_intPlayStatus
                Case ITPlayerState.ITPlayerStateStopped '停止中
                    m_objITunes.Play()
                Case Else
                    m_objITunes.Stop()
            End Select

            '●画面更新
            Call F10ScrUpDate()

        Catch ex As Exception
            '---------------------------------------------
            ' 例外処理
            '---------------------------------------------
            '●戻り値にエラーを設定
            lngRet = eReturnCode.NG

            If p_objErrCtrl IsNot Nothing Then
                '●エラー情報の取得
                With p_objErrCtrl
                    .ErrMod = Reflection.MethodBase.GetCurrentMethod.Name
                    .ErrNum = Err.Number
                    .ErrDesc = ex.Message.ToString
                End With

                '●エラー情報出力
                If p_objErrCtrl.LogWrite() = eReturnCode.NG Then
                    MessageBox.Show("エラーログ出力に失敗しました。", "ログ出力エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                End If
            Else
                '※エラー制御クラスがいない場合は終了する
                MessageBox.Show("想定外のエラーが発生しました。終了します", "エラー発生", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                frm01ToolMain.Close()
            End If

        Finally
            '---------------------------------------------
            ' 終了処理
            '---------------------------------------------
            '●戻り値リターン
            objBtnPlay_Click = lngRet

            '●デバッグログ出力
            Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name.ToString & "-End!")
        End Try
    End Function
    '---------------------------------------------
    '[Name]     objBtnNextTrack_Click
    '[J.Name]   次のトラッククリック
    '[Date]     2010/10/31
    '[Func]     次のトラッククリック
    '[Input]    なし
    '[Output]   lngRet as long          
    '[Memo]
    '---------------------------------------------
    Public Function objBtnNextTrack_Click() As Long
        '---------------------------------------------
        ' ローカル変数宣言
        '---------------------------------------------
        Dim lngRet As Long
        '---------------------------------------------
        ' 関数メイン
        '---------------------------------------------
        Try
            '---------------------------------------------
            ' 初期処理
            '---------------------------------------------
            '●デバッグログ出力
            Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name.ToString & "-Start!")

            '●戻り値初期化
            lngRet = eReturnCode.OK
            '---------------------------------------------
            ' メイン処理
            '---------------------------------------------
            '●次のトラックへ
            m_objITunes.NextTrack()

            '●最終か確認
            If m_objITunes.PlayerState <> ITPlayerState.ITPlayerStateStopped Then
                m_intIdxSelectSong = m_intIdxSelectSong + 1
                Call F20DataLoad()
            End If

            '●画面更新
            Call F10ScrUpDate()

        Catch ex As Exception
            '---------------------------------------------
            ' 例外処理
            '---------------------------------------------
            '●戻り値にエラーを設定
            lngRet = eReturnCode.NG

            If p_objErrCtrl IsNot Nothing Then
                '●エラー情報の取得
                With p_objErrCtrl
                    .ErrMod = Reflection.MethodBase.GetCurrentMethod.Name
                    .ErrNum = Err.Number
                    .ErrDesc = ex.Message.ToString
                End With

                '●エラー情報出力
                If p_objErrCtrl.LogWrite() = eReturnCode.NG Then
                    MessageBox.Show("エラーログ出力に失敗しました。", "ログ出力エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                End If
            Else
                '※エラー制御クラスがいない場合は終了する
                MessageBox.Show("想定外のエラーが発生しました。終了します", "エラー発生", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                frm01ToolMain.Close()
            End If

        Finally
            '---------------------------------------------
            ' 終了処理
            '---------------------------------------------
            '●戻り値リターン
            objBtnNextTrack_Click = lngRet

            '●デバッグログ出力
            Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name.ToString & "-End!")
        End Try
    End Function
    '---------------------------------------------
    '[Name]     objBtnBackTrack_Click
    '[J.Name]   前のトラッククリック
    '[Date]     2010/10/31
    '[Func]     前のトラッククリック
    '[Input]    なし
    '[Output]   lngRet as long          
    '[Memo]
    '---------------------------------------------
    Public Function objBtnBackTrack_Click() As Long
        '---------------------------------------------
        ' ローカル変数宣言
        '---------------------------------------------
        Dim lngRet As Long
        '---------------------------------------------
        ' 関数メイン
        '---------------------------------------------
        Try
            '---------------------------------------------
            ' 初期処理
            '---------------------------------------------
            '●デバッグログ出力
            Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name.ToString & "-Start!")

            '●戻り値初期化
            lngRet = eReturnCode.OK
            '---------------------------------------------
            ' メイン処理
            '---------------------------------------------
            '●次のトラックへ
            m_objITunes.BackTrack()

            '●最終か確認
            If m_objITunes.PlayerState <> ITPlayerState.ITPlayerStateStopped Then
                m_intIdxSelectSong = m_intIdxSelectSong - 1
                Call F20DataLoad()
            End If

            '●画面更新
            Call F10ScrUpDate()

        Catch ex As Exception
            '---------------------------------------------
            ' 例外処理
            '---------------------------------------------
            '●戻り値にエラーを設定
            lngRet = eReturnCode.NG

            If p_objErrCtrl IsNot Nothing Then
                '●エラー情報の取得
                With p_objErrCtrl
                    .ErrMod = Reflection.MethodBase.GetCurrentMethod.Name
                    .ErrNum = Err.Number
                    .ErrDesc = ex.Message.ToString
                End With

                '●エラー情報出力
                If p_objErrCtrl.LogWrite() = eReturnCode.NG Then
                    MessageBox.Show("エラーログ出力に失敗しました。", "ログ出力エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                End If
            Else
                '※エラー制御クラスがいない場合は終了する
                MessageBox.Show("想定外のエラーが発生しました。終了します", "エラー発生", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                frm01ToolMain.Close()
            End If

        Finally
            '---------------------------------------------
            ' 終了処理
            '---------------------------------------------
            '●戻り値リターン
            objBtnBackTrack_Click = lngRet

            '●デバッグログ出力
            Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name.ToString & "-End!")
        End Try
    End Function
    '---------------------------------------------
    '[Name]     objPlayList_Click
    '[J.Name]   プレイリストクリック
    '[Date]     2010/10/31
    '[Func]     プレイリストクリック
    '[Input]    v_index as integer      押下したインデックス
    '[Output]   lngRet as long          
    '[Memo]
    '---------------------------------------------
    Public Function objPlayList_Click(ByVal v_index As Integer) As Long
        '---------------------------------------------
        ' ローカル変数宣言
        '---------------------------------------------
        Dim lngRet As Long
        '---------------------------------------------
        ' 関数メイン
        '---------------------------------------------
        Try
            'MsgBox(v_index)
            '---------------------------------------------
            ' 初期処理
            '---------------------------------------------
            '●デバッグログ出力
            Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name.ToString & "-Start!")

            '●戻り値初期化
            lngRet = eReturnCode.OK

            '●選択中ソング番号初期化
            m_intIdxSelectSong = 0
            '---------------------------------------------
            ' メイン処理
            '---------------------------------------------
            '●選択中のプレイリスト番号格納
            m_intIdxSelectPL = v_index + 1

            '●リストビュー更新
            Call objSongListUpDate()

        Catch ex As Exception
            '---------------------------------------------
            ' 例外処理
            '---------------------------------------------
            '●戻り値にエラーを設定
            lngRet = eReturnCode.NG

            If p_objErrCtrl IsNot Nothing Then
                '●エラー情報の取得
                With p_objErrCtrl
                    .ErrMod = Reflection.MethodBase.GetCurrentMethod.Name
                    .ErrNum = Err.Number
                    .ErrDesc = ex.Message.ToString
                End With

                '●エラー情報出力
                If p_objErrCtrl.LogWrite() = eReturnCode.NG Then
                    MessageBox.Show("エラーログ出力に失敗しました。", "ログ出力エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                End If
            Else
                '※エラー制御クラスがいない場合は終了する
                MessageBox.Show("想定外のエラーが発生しました。終了します", "エラー発生", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                frm01ToolMain.Close()
            End If

        Finally
            '---------------------------------------------
            ' 終了処理
            '---------------------------------------------
            '●戻り値リターン
            objPlayList_Click = lngRet

            '●デバッグログ出力
            Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name.ToString & "-End!")
        End Try
    End Function
    '---------------------------------------------
    '[Name]     objListView_DblCkick
    '[J.Name]   リストビューWクリック
    '[Date]     2010/10/31
    '[Func]     リストビューWクリック
    '[Input]    v_index as integer      押下したインデックス
    '[Output]   lngRet as long          
    '[Memo]
    '---------------------------------------------
    Public Function objListView_DblCkick(ByVal v_index As Integer) As Long
        '---------------------------------------------
        ' ローカル変数宣言
        '---------------------------------------------
        Dim lngRet As Long
        '---------------------------------------------
        ' 関数メイン
        '---------------------------------------------
        Try
            '---------------------------------------------
            ' 初期処理
            '---------------------------------------------
            '●デバッグログ出力
            Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name.ToString & "-Start!")

            '●戻り値初期化
            lngRet = eReturnCode.OK
            '---------------------------------------------
            ' メイン処理
            '---------------------------------------------
            '●選択中のプレイリスト番号格納
            m_intIdxSelectSong = v_index

            '●未選択状態なら終了
            If m_intIdxSelectSong = 0 Then
                Exit Function
            End If

            '●再生する ※※移動する！！！！！！！！！！！！！※※
            m_objITunes.LibrarySource.Playlists(m_intIdxSelectPL).Tracks(m_intIdxSelectSong).Play()
            'ソートするとバグる・・・


            Call F20DataLoad()

            Call F10ScrUpDate()

        Catch ex As Exception
            '---------------------------------------------
            ' 例外処理
            '---------------------------------------------
            '●戻り値にエラーを設定
            lngRet = eReturnCode.NG

            If p_objErrCtrl IsNot Nothing Then
                '●エラー情報の取得
                With p_objErrCtrl
                    .ErrMod = Reflection.MethodBase.GetCurrentMethod.Name
                    .ErrNum = Err.Number
                    .ErrDesc = ex.Message.ToString
                End With

                '●エラー情報出力
                If p_objErrCtrl.LogWrite() = eReturnCode.NG Then
                    MessageBox.Show("エラーログ出力に失敗しました。", "ログ出力エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                End If
            Else
                '※エラー制御クラスがいない場合は終了する
                MessageBox.Show("想定外のエラーが発生しました。終了します", "エラー発生", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                frm01ToolMain.Close()
            End If

        Finally
            '---------------------------------------------
            ' 終了処理
            '---------------------------------------------
            '●戻り値リターン
            objListView_DblCkick = lngRet

            '●デバッグログ出力
            Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name.ToString & "-End!")
        End Try

    End Function
    '---------------------------------------------
    '[Name]     volume_ctrl
    '[J.Name]   サウンドボリュームコントロール
    '[Date]     2010/11/01
    '[Func]     サウンドボリュームコントロール
    '[Input]    v_intValue as integer           '制御コード
    '           v_intVolume as integer = 0      'ボリューム値
    '[Output]   lngRet as long          
    '[Memo]
    '---------------------------------------------
    Public Function volume_ctrl(ByVal v_intValue As Integer, Optional ByVal v_intVolume As Integer = 0) As Long
        '---------------------------------------------
        ' ローカル変数宣言
        '---------------------------------------------
        Dim lngRet As Long
        '---------------------------------------------
        ' 関数メイン
        '---------------------------------------------
        Try
            '---------------------------------------------
            ' 初期処理
            '---------------------------------------------
            '●デバッグログ出力
            Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name.ToString & "-Start!")

            '●戻り値初期化
            lngRet = eReturnCode.OK
            '---------------------------------------------
            ' メイン処理
            '---------------------------------------------
            Select Case v_intValue
                Case eVolumeCtrl.MUTE
                    If m_blnMute = True Then
                        m_blnMute = False
                        m_objITunes.Mute = False
                    Else
                        m_blnMute = True
                        m_objITunes.Mute = True
                    End If
                Case eVolumeCtrl.Up
                    m_intVolume = m_intVolume + 10
                    m_objITunes.SoundVolume = m_intVolume
                Case eVolumeCtrl.Down
                    m_intVolume = m_intVolume - 10
                    m_objITunes.SoundVolume = m_intVolume
                Case eVolumeCtrl.SETVOLUME
                    m_intVolume = v_intVolume
            End Select

            Call F10ScrUpDate()

        Catch ex As Exception
            '---------------------------------------------
            ' 例外処理
            '---------------------------------------------
            '●戻り値にエラーを設定
            lngRet = eReturnCode.NG

            If p_objErrCtrl IsNot Nothing Then
                '●エラー情報の取得
                With p_objErrCtrl
                    .ErrMod = Reflection.MethodBase.GetCurrentMethod.Name
                    .ErrNum = Err.Number
                    .ErrDesc = ex.Message.ToString
                End With

                '●エラー情報出力
                If p_objErrCtrl.LogWrite() = eReturnCode.NG Then
                    MessageBox.Show("エラーログ出力に失敗しました。", "ログ出力エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                End If
            Else
                '※エラー制御クラスがいない場合は終了する
                MessageBox.Show("想定外のエラーが発生しました。終了します", "エラー発生", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                frm01ToolMain.Close()
            End If

        Finally
            '---------------------------------------------
            ' 終了処理
            '---------------------------------------------
            '●戻り値リターン
            volume_ctrl = lngRet

            '●デバッグログ出力
            Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name.ToString & "-End!")
        End Try
    End Function







    '□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□
    '　　　　　　　　　　　　　　　　　　　　　　開発中
    '□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□
    '---------------------------------------------
    '[Name]     F00Initialize
    '[J.Name]   画面イニシャライズ
    '[Date]     2010/10/21
    '[Func]     画面イニシャライズ
    '[Input]    なし
    '[Output]   lngRet as long
    '[Memo]
    '---------------------------------------------
    Private Function F00Initialize() As Long
        '---------------------------------------------
        ' ローカル変数宣言
        '---------------------------------------------
        Dim lngRet As Long
        '---------------------------------------------
        ' 関数メイン
        '---------------------------------------------
        Try
            '---------------------------------------------
            ' 初期処理
            '---------------------------------------------
            '●デバッグログ出力
            Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name.ToString & "-Start!")

            '●戻り値初期化
            lngRet = eReturnCode.OK
            '---------------------------------------------
            ' メイン処理
            '---------------------------------------------
            With frm01ToolMain
                '●タイトルバー
                .Text = My.Application.Info.Title & _
                        "Version:" & _
                        My.Application.Info.Version.Major & _
                        "." & _
                        My.Application.Info.Version.Minor

                '●プレイリスト
                .objListBoxPlayList.Items.Clear()

                For i = 1 To m_objITunes.LibrarySource.Playlists.Count
                    .objListBoxPlayList.Items.Add(m_objITunes.LibrarySource.Playlists.Item(i).Name)
                    'objListBoxPlayList.Items.AddRange(m_objITunes.LibrarySource.Playlists)
                    'なぜか入れられない・・・
                Next

                '●ボリューム
                m_intVolume = m_objITunes.SoundVolume

                '●ミュート状態
                If m_objITunes.Mute = True Then
                    m_blnMute = True
                Else
                    m_blnMute = False
                End If

                '●タブの初期化
                Call InitTab()

            End With

        Catch ex As Exception
            '---------------------------------------------
            ' 例外処理
            '---------------------------------------------
            '●戻り値にエラーを設定
            lngRet = eReturnCode.NG

            If p_objErrCtrl IsNot Nothing Then
                '●エラー情報の取得
                With p_objErrCtrl
                    .ErrMod = Reflection.MethodBase.GetCurrentMethod.Name
                    .ErrNum = Err.Number
                    .ErrDesc = ex.Message.ToString
                End With

                '●エラー情報出力
                If p_objErrCtrl.LogWrite() = eReturnCode.NG Then
                    MessageBox.Show("エラーログ出力に失敗しました。", "ログ出力エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                End If
            Else
                '※エラー制御クラスがいない場合は終了する
                MessageBox.Show("想定外のエラーが発生しました。終了します", "エラー発生", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                frm01ToolMain.Close()
            End If

        Finally
            '---------------------------------------------
            ' 終了処理
            '---------------------------------------------
            '●戻り値リターン
            F00Initialize = lngRet

            '●デバッグログ出力
            Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name.ToString & "-End!")
        End Try

    End Function
    '---------------------------------------------
    '[Name]     F10ScrUpDate
    '[J.Name]   画面再表示
    '[Date]     2010/10/21
    '[Func]     画面再表示
    '[Input]    なし
    '[Output]   lngRet as long
    '[Memo]
    '---------------------------------------------
    Public Function F10ScrUpDate() As Long
        '---------------------------------------------
        ' ローカル変数宣言
        '---------------------------------------------
        Dim lngRet As Long
        Dim strSplitComment() As String
        '---------------------------------------------
        ' 関数メイン
        '---------------------------------------------
        Try
            '---------------------------------------------
            ' 初期処理
            '---------------------------------------------
            '●デバッグログ出力
            Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name.ToString & "-Start!")

            '●戻り値初期化
            lngRet = eReturnCode.OK

            '---------------------------------------------
            ' メイン処理
            '---------------------------------------------
            With frm01ToolMain

                '●ボリューム情報
                .objBtnSlider.Value = m_intVolume / 10

                '●カレントトラック情報
                .objStatusLabel1.Text = ""

                '●ミュート状態
                If m_blnMute = True Then
                    .VolumeMuteMToolStripMenuItem.Checked = True
                Else
                    .VolumeMuteMToolStripMenuItem.Checked = False
                End If

                '●再生状況
                m_intPlayStatus = m_objITunes.PlayerState

                Select Case m_intPlayStatus
                    Case iTunesLib.ITPlayerState.ITPlayerStateStopped '停止中
                        .objBtnPlay.Text = c_BtnPlay_Play
                        .StopSToolStripMenuItem.Text = c_MnCtrl_Play
                    Case Else
                        .objBtnPlay.Text = c_BtnPlay_Stop
                        .StopSToolStripMenuItem.Text = c_MnCtrl_Stop
                End Select

                If m_strctCurrentTrack.blnCurrentOn = True Then
                    '●曲名他
                    .objStatusLabel1.Text = m_strctCurrentTrack.strTitle & " / " & _
                                            m_strctCurrentTrack.strArtist & " from 「" & _
                                            m_strctCurrentTrack.strAlbum & "」"

                    '●レート
                    .objSliderRating.Value = m_strctCurrentTrack.intRate
                    .objLblRating.Text = m_strctCurrentTrack.intRate.ToString
                    .objPgsBarRate.Value = m_strctCurrentTrack.intRate

                    '●コメント
                    .objTxtBoxComment.Text = m_strctCurrentTrack.strComment
                    .objTextMakeComment.Text = m_strMakeComment

                    '●アートワーク
                    If m_strctCurrentTrack.objArtWork.Count <> 0 Then
                        m_strctCurrentTrack.objArtWork.Item(m_intArtWorkNo).SaveArtworkToFile(My.Application.Info.DirectoryPath & "\temp.bmp")
                        .objPictArtWork.Load(My.Application.Info.DirectoryPath & "\temp.bmp")
                    Else
                        .objPictArtWork.Image = Nothing
                    End If
                End If

                '.objListView.Items(m_intIdxSelectSong - 1).Selected = True
                '.objListView.Focus()





                '●変更有無
                .objCbxChange.Checked = m_blnChanged

                If m_blnChanged = False Then
                    'カレントトラックのコメントを分割
                    strSplitComment = Split(m_strctCurrentTrack.strComment, "]")

                    '全てFalseで初期化する
                    For i = 0 To p_strctData.Length - 1
                        For j = 0 To p_strctData(i).chkbox.Length - 1
                            p_strctData(i).chkbox(j).Checked = False
                        Next
                    Next

                    '分割したコメント数分繰り返し
                    For i = 0 To UBound(strSplitComment)
                        If strSplitComment(i) = "" Then
                            Exit For
                        Else
                            'タブ数分ループ
                            For j = 0 To p_strctData.Length - 1
                                'タブ内のチェックボックス分ループ
                                For k = 0 To p_strctData(j).chkbox.Length - 1
                                    '中身が一致していた場合チェックをONにする
                                    If p_strctData(j).chkbox(k).Text = Replace(strSplitComment(i), "[", "") Then
                                        p_strctData(j).chkbox(k).Checked = True
                                    End If
                                Next
                            Next
                        End If
                    Next
                End If
            End With
        Catch ex As Exception
            '---------------------------------------------
            ' 例外処理
            '---------------------------------------------
            '●戻り値にエラーを設定
            lngRet = eReturnCode.NG

            If p_objErrCtrl IsNot Nothing Then
                '●エラー情報の取得
                With p_objErrCtrl
                    .ErrMod = Reflection.MethodBase.GetCurrentMethod.Name
                    .ErrNum = Err.Number
                    .ErrDesc = ex.Message.ToString
                End With

                '●エラー情報出力
                If p_objErrCtrl.LogWrite() = eReturnCode.NG Then
                    MessageBox.Show("エラーログ出力に失敗しました。", "ログ出力エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                End If
            Else
                '※エラー制御クラスがいない場合は終了する
                MessageBox.Show("想定外のエラーが発生しました。終了します", "エラー発生", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                frm01ToolMain.Close()
            End If

        Finally
            '---------------------------------------------
            ' 終了処理
            '---------------------------------------------
            '●戻り値リターン
            F10ScrUpDate = lngRet

            '●デバッグログ出力
            Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name.ToString & "-End!")
        End Try

    End Function

    '●ソングリスト更新
    Private Sub objSongListUpDate()
        Dim i As Integer
        Dim objLvDat() As ListViewItem
        Dim objTracks As IITTrackCollection
        ReDim objLvDat(m_objITunes.LibrarySource.Playlists.Item(m_intIdxSelectPL).Tracks.Count - 1)

        frm01ToolMain.Cursor = Cursors.WaitCursor

        objTracks = m_objITunes.LibrarySource.Playlists.Item(m_intIdxSelectPL).Tracks

        For i = 1 To objLvDat.Length
            objLvDat(i - 1) = New ListViewItem
            With objTracks(i)
                objLvDat(i - 1).Text = i
                objLvDat(i - 1).SubItems.Add(.Name)
                objLvDat(i - 1).SubItems.Add(.Artist)
                objLvDat(i - 1).SubItems.Add(.Album)
                objLvDat(i - 1).SubItems.Add(.Rating)
                objLvDat(i - 1).SubItems.Add(.PlayedCount)
                objLvDat(i - 1).SubItems.Add(.Comment)
            End With
        Next

        With frm01ToolMain.objListView
            .Items.Clear()
            .Items.AddRange(objLvDat)
            If m_intIdxSelectSong <> 0 Then
                .EnsureVisible(m_intIdxSelectSong - 1)
                .Items(m_intIdxSelectSong - 1).Selected = True
                .Focus()
            End If
        End With

        frm01ToolMain.Cursor = Cursors.Default


    End Sub

    '●仮想リストビューへのアイテム追加
    Public Sub objLV_RetVirtualItem(ByVal sender As Object, ByVal e As System.Windows.Forms.RetrieveVirtualItemEventArgs)

        'e.Item = objLvDat(e.ItemIndex)

    End Sub

    '●チェックボックスクリック
    Private Sub objPreSetCbx_Click(ByVal sender As System.Object, _
                ByVal e As System.EventArgs)
        'イベントの送り側の名前を取得
        Dim senderName As String = DirectCast(sender, CheckBox).Name

        'ボタンのベース名　長さの取得に使用
        Dim strBut As String = "objPreSetCbx"

        'objPreSetCbxのxxを取得して数字に直している
        Dim index As Integer = CInt(senderName.Substring(strBut.Length, _
                                    senderName.Length - strBut.Length))

        '取得したインデックスで各ボタンのクリックイベントを処理
        'MsgBox(p_strctData(frm01ToolMain.objTabCtrl.SelectedIndex).chkbox(index).Text)
        'MsgBox(objPreSetCbx(index).Text)

        If p_strctData(frm01ToolMain.objTabCtrl.SelectedIndex).chkbox(index).Checked = True Then
            m_strMakeComment = m_strMakeComment & "[" & p_strctData(frm01ToolMain.objTabCtrl.SelectedIndex).chkbox(index).Text & "]"
        Else
            m_strMakeComment = m_strMakeComment.Replace("[" & p_strctData(frm01ToolMain.objTabCtrl.SelectedIndex).chkbox(index).Text & "]", "")
        End If

        m_blnChanged = True

        Call F10ScrUpDate()

    End Sub

    '●タグ追加
    Public Sub AddTag()
        Dim intTabIndex As Integer = frm01ToolMain.objTabCtrl.SelectedIndex
        Dim strTmp As String
        strTmp = InputBox("タグ名称を入力して下さい", "タグの追加")
        If strTmp <> "" Then
            p_objMySetting(intTabIndex).Add(strTmp)
            My.Settings.Save()
        End If
        InitTab()
    End Sub

    '●コメント挿入
    Public Sub setComment()
        If m_strMakeComment IsNot Nothing Then
            m_objITunes.CurrentTrack.Comment = m_strMakeComment
        End If
        m_objITunes.CurrentTrack.Rating = m_strctCurrentTrack.intRate
        m_blnChanged = False
        'リストアップデート
        Call objSongListUpDate()
    End Sub

    '●レート設定
    Public Sub setRating()
        With frm01ToolMain
            .objLblRating.Text = .objSliderRating.Value
            .objPgsBarRate.Value = .objSliderRating.Value
            m_strctCurrentTrack.intRate = .objSliderRating.Value
            m_blnChanged = True
        End With
    End Sub

    '●タブの初期化
    Private Sub InitTab()

        Dim Panel As FlowLayoutPanel

        With frm01ToolMain
            '●タブ
            For i = 0 To p_strctData.Length - 1
                With .objTabCtrl.TabPages(i)
                    '●コントロールをクリアする
                    .Controls.Clear()

                    '●タブ名称
                    .Text = My.Settings.strArrayTabName(i)

                    With p_strctData(i)
                        '●コントロールのクリア


                        '●パネル生成
                        Panel = New FlowLayoutPanel
                        Panel.Dock = DockStyle.Fill

                        '●データ部確保
                        ReDim .chkbox(p_objMySetting(i).Count - 1)

                        '●データ格納
                        For j = 0 To .chkbox.Length - 1
                            .chkbox(j) = New CheckBox
                            .chkbox(j).Name = "objPreSetCbx" & j.ToString
                            .chkbox(j).Text = p_objMySetting(i).Item(j).ToString
                            AddHandler .chkbox(j).Click, AddressOf objPreSetCbx_Click
                            Panel.Controls.Add(.chkbox(j))
                        Next
                    End With

                    .Controls.Add(Panel)
                End With
            Next
        End With

    End Sub

    '●編集ロック
    Public Sub SetEditLock()
        m_blnChanged = frm01ToolMain.objCbxChange.Checked
    End Sub

    '●iTunesイベント
    Private Sub m_objITunes_OnPlayerPlayEvent(ByVal iTrack As Object) Handles m_objITunes.OnPlayerPlayEvent
        'MsgBox(m_objITunes.CurrentPlaylist.Index)
        Call F20DataLoad()
    End Sub

    '●曲変更時必ず呼ぶファンクション
    Private Sub F20DataLoad()

        Dim clearCurrent As strctCurrentTrack
        '●ゴミを消去する
        clearCurrent = New strctCurrentTrack
        m_strctCurrentTrack = clearCurrent
        m_intArtWorkNo = 1

        If m_blnChanged = False Then
            With m_strctCurrentTrack

                If m_objITunes.CurrentTrack.Artist IsNot Nothing Then
                    .strArtist = m_objITunes.CurrentTrack.Artist
                End If
                If m_objITunes.CurrentTrack.Album IsNot Nothing Then
                    .strAlbum = m_objITunes.CurrentTrack.Album
                End If
                If m_objITunes.CurrentTrack.Name IsNot Nothing Then
                    .strTitle = m_objITunes.CurrentTrack.Name
                End If

                If m_objITunes.CurrentTrack.Comment IsNot Nothing Then
                    .strComment = m_objITunes.CurrentTrack.Comment
                End If

                .objArtWork = m_objITunes.CurrentTrack.Artwork
                .intRate = m_objITunes.CurrentTrack.Rating
                .blnCurrentOn = True
                m_strMakeComment = .strComment
            End With
        End If

    End Sub

    Public Sub objArtWork_Click()
        If m_intArtWorkNo < m_objITunes.CurrentTrack.Artwork.Count Then
            m_intArtWorkNo = m_intArtWorkNo + 1
        Else
            m_intArtWorkNo = 1
        End If
        F10ScrUpDate()
    End Sub

End Class
