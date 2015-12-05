'---------------------------------------------
'[Name]     frm01ToolMain
'[J.Name]   ツールメインフォーム
'[Date]     2010/10/20
'[Func]     フォーム
'[Memo]
'---------------------------------------------
Public Class frm01ToolMain
    '---------------------------------------------
    ' メンバ変数宣言
    '---------------------------------------------
    Private m_objToolMain As cls01frmToolMain       'フォーム制御クラス
    '---------------------------------------------
    '[Name]     frm01ToolMain_Load
    '[J.Name]   フォームロード
    '[Date]     2010/10/20
    '[Func]     フォーム初期化
    '[Input]    ByVal sender As Object
    '           ByVal e As System.EventArgs
    '[Output]   なし
    '[Memo]
    '---------------------------------------------
    Private Sub frm01ToolMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
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
            Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name & "-Start!")

            '●各種グローバルクラスのインスタンス化
            p_objErrCtrl = New cls99ErrCtrl         'エラー制御クラス

            '●フォーム操作クラスのインスタンス化
            m_objToolMain = New cls01frmToolMain    'フォーム制御クラス
            '---------------------------------------------
            ' メイン処理
            '---------------------------------------------
            '●カレントトラックデータ取得
            'p_objiTunesCtrl.GetCurrentTrack()

            'Call SetCurrentTrackDat()

            '●監視タイマ起動
            objTimer.Start()

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

                '●終了する
                Me.Close()
            Else
                '※エラー制御クラスがいない場合は終了する
                MessageBox.Show("想定外のエラーが発生しました。終了します", "エラー発生", MessageBoxButtons.OK, MessageBoxIcon.Stop)

                '●終了する
                Me.Close()
            End If


        Finally
            '---------------------------------------------
            ' 終了処理
            '---------------------------------------------
            '●デバッグログ出力
            Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name & "-End!")
        End Try
    End Sub
    '---------------------------------------------
    '[Name]     frm01ToolMain_FormClosed
    '[J.Name]   フォームクローズ
    '[Date]     2010/10/20
    '[Func]     フォームクローズ時処理
    '[Input]    ByVal sender As Object
    '           ByVal e As System.Windows.Forms.FormClosedEventArgs
    '[Output]   なし
    '[Memo]
    '---------------------------------------------
    Private Sub frm01ToolMain_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '---------------------------------------------
        ' ローカル変数宣言
        '---------------------------------------------
        Dim i As Integer
        '---------------------------------------------
        ' 関数メイン
        '---------------------------------------------
        Try
            '---------------------------------------------
            ' 初期処理
            '---------------------------------------------
            '●デバッグログ出力
            Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name & "-Start!")
            '---------------------------------------------
            ' メイン処理
            '---------------------------------------------
            For Each objTabPage As TabPage In objTabCtrl.TabPages
                My.Settings.strArrayTabName(i) = objTabPage.Text
                i = i + 1
            Next

            My.Settings.Save()

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
                If p_objErrCtrl.LogWrite = eReturnCode.NG Then
                    MessageBox.Show("エラーログ出力に失敗しました。", "ログ出力エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                End If
            Else
                '※エラー制御クラスがいない場合は終了する
                MessageBox.Show("想定外のエラーが発生しました。終了します", "エラー発生", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                End
            End If
        Finally
            '---------------------------------------------
            ' 終了処理
            '---------------------------------------------
            '●デバッグログ出力
            Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name & "-End!")

            '●オブジェクトの破棄
            p_objErrCtrl = Nothing

            '●プログラム終了
            End
        End Try

    End Sub
    '---------------------------------------------
    '[Name]     objBtnPlay_Click
    '[J.Name]   「PLAY」ボタン押下
    '[Date]     2010/10/21
    '[Func]     「PLAY」ボタン押下
    '[Input]    なし
    '[Output]   なし
    '[Memo]
    '---------------------------------------------
    Private Sub objBtnPlay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles objBtnPlay.Click
        m_objToolMain.objBtnPlay_Click()
    End Sub
    '---------------------------------------------
    '[Name]     objBtnNext_Click
    '[J.Name]   「→|」ボタン押下
    '[Date]     2010/10/30
    '[Func]     NextTrack
    '[Input]    なし
    '[Output]   なし
    '[Memo]
    '---------------------------------------------
    Private Sub objBtnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles objBtnNext.Click
        m_objToolMain.objBtnNextTrack_Click()
    End Sub
    '---------------------------------------------
    '[Name]     objBtnPre_Click
    '[J.Name]   「|←」ボタン押下
    '[Date]     2010/10/30
    '[Func]     BackTrack
    '[Input]    なし
    '[Output]   なし
    '[Memo]
    '---------------------------------------------
    Private Sub objBtnPre_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles objBtnPre.Click
        m_objToolMain.objBtnBackTrack_Click()
    End Sub

    '●Rew
    Private Sub objBtnRew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles objBtnRew.Click
        'p_objiTunesCtrl.Rewind()
    End Sub
    '●FF
    Private Sub objBtnFF_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles objBtnFF.Click
        '    p_objiTunesCtrl.FastForward()
    End Sub

    '●リストボックスを押下
    Private Sub objListBoxPlayList_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles objListBoxPlayList.SelectedIndexChanged
        m_objToolMain.objPlayList_Click(Me.objListBoxPlayList.SelectedIndex)
    End Sub
    '●リストビューダブルクリック
    Private Sub objListView_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles objListView.DoubleClick, PlayPToolStripMenuItem.Click
        m_objToolMain.objListView_DblCkick(objListView.SelectedItems(0).Index + 1)
    End Sub

    '●「タグ追加」ボタン押下
    Private Sub objBtnAddTag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles objBtnAddTag.Click
        m_objToolMain.AddTag()
    End Sub

    '●「挿入」ボタン押下
    Private Sub objBtnInsComment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles objBtnInsComment.Click
        m_objToolMain.setComment()
    End Sub

    '●編集ロックボタン
    Private Sub objCbxChange_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles objCbxChange.CheckedChanged
        m_objToolMain.SetEditLock()
    End Sub


    '↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
    '　　　　　　　　　　　　　　　　　　　開発中
    '---------------------------------------------
    '[Name]     objBtnAddTab_Click
    '[J.Name]   「タブの追加」ボタン押下
    '[Date]     2010/10/20
    '[Func]     タブページの追加
    '[Input]    なし
    '[Output]   lngRet as long
    '[Memo]
    '---------------------------------------------
    Private Sub objBtnAddTab_Click() Handles objBtnAddTab.Click, AddAToolStripMenuItem.Click
        '---------------------------------------------
        ' ローカル変数宣言
        '---------------------------------------------
        Dim strTabName As String
        Dim Panel As FlowLayoutPanel
        '---------------------------------------------
        ' 関数メイン
        '---------------------------------------------
        Try
            '---------------------------------------------
            ' 初期処理
            '---------------------------------------------
            '●デバッグログ出力
            Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name & "-Start!")
            '---------------------------------------------
            ' メイン処理
            '---------------------------------------------
            '●新規タブページを追加する
            With objTabCtrl

                '●タブ名称入力
                strTabName = InputBox("タブ名称を入力して下さい", "タブ名称入力", "NewTab" & .TabPages.Count + 1)

                If strTabName <> "" Then

                    .TabPages.Add(strTabName)
                    .TabPages(.TabPages.Count - 1).BackColor = Color.White
                    Panel = New FlowLayoutPanel


                    'ChkBox = New CheckBox(36) {}
                    'For i = 0 To UBound(ChkBox)
                    '    ChkBox(i) = New CheckBox
                    '    ChkBox(i).Text = "ChkBox" & i
                    '    Panel.Dock = DockStyle.Fill
                    '    Panel.Controls.Add(ChkBox(i))
                    '    Panel.AutoScroll = True
                    'Next

                    .Controls.Add(Panel)

                    '●タブが存在する場合は「削除」押下可能
                    If .TabPages.Count >= 1 Then
                        objBtnTabDelete.Enabled = True
                    End If
                End If
            End With

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
                If p_objErrCtrl.LogWrite = eReturnCode.NG Then
                    MessageBox.Show("エラーログ出力に失敗しました。", "ログ出力エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                End If
            Else
                '※エラー制御クラスがいない場合は終了する
                MessageBox.Show("想定外のエラーが発生しました。終了します", "エラー発生", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                End
            End If
        Finally
            '---------------------------------------------
            ' 終了処理
            '---------------------------------------------
            '●デバッグログ出力
            Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name & "-End!")
        End Try
    End Sub
    '---------------------------------------------
    '[Name]     objBtnDelTab_Click
    '[J.Name]   「タブの削除」ボタン押下
    '[Date]     2010/10/20
    '[Func]     タブページの削除
    '[Input]    なし
    '[Output]   lngRet as long
    '[Memo]
    '---------------------------------------------
    Private Sub objBtnDelTab_Click() Handles DeleteDToolStripMenuItem.Click, objBtnTabDelete.Click
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
            Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name & "-Start!")
            '---------------------------------------------
            ' メイン処理
            '---------------------------------------------
            '●選択中のタブを削除する
            With objTabCtrl
                If .SelectedIndex <> -1 Then
                    '●タブを削除する
                    .TabPages.RemoveAt(.SelectedIndex)

                    '●タブが無くなったら「削除」押下不可
                    If .TabPages.Count = 0 Then
                        objBtnTabDelete.Enabled = False
                    End If
                End If
            End With

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
                If p_objErrCtrl.LogWrite = eReturnCode.NG Then
                    MessageBox.Show("エラーログ出力に失敗しました。", "ログ出力エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                End If
            Else
                '※エラー制御クラスがいない場合は終了する
                MessageBox.Show("想定外のエラーが発生しました。終了します", "エラー発生", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                End
            End If

            Me.Close()
        Finally
            '---------------------------------------------
            ' 終了処理
            '---------------------------------------------
            '●デバッグログ出力
            Debug.Print(Reflection.MethodBase.GetCurrentMethod.Name & "-End!")
        End Try
    End Sub


    Private Sub NameNToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NameNToolStripMenuItem.Click
        Dim strName As String
        strName = InputBox("TabName", "TabName", "タブ名称")
        If strName <> "" Then
            objTabCtrl.TabPages(objTabCtrl.SelectedIndex).Text = strName
        End If

    End Sub

    Private Sub AddAToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddAToolStripMenuItem1.Click
        Dim strName As String
        strName = InputBox("TagName", "TagName", "タグ名称")
        If strName <> "" Then
            Select Case objTabCtrl.SelectedIndex
                Case 0
                    My.Settings.strArrayData01.Insert(My.Settings.strArrayData01.Count, strName)
                Case 1
                Case 2
                Case 3
                Case 4
                Case 5
                Case 6
                Case 7
                Case 8
                Case 9
                Case Else
            End Select
        End If

        My.Settings.Save()

        'm_objToolMain.F00Initialize()

    End Sub






    '●ボリュームセット
    Private Sub objBtnSlider_Scroll(ByVal sender As Object, ByVal e As System.EventArgs) Handles objBtnSlider.Scroll
        'スライダー→関数呼び出しだとスムーズにならない・・・
        'm_objToolMain.volume_ctrl(eVolumeCtrl.SETVOLUME, objBtnSlider.Value)
    End Sub

    Private Sub objSliderRating_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles objSliderRating.Scroll
        m_objToolMain.setRating()
    End Sub

    Private Sub frm01ToolMain_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        'p_objiTunesCtrl.GetCurrentTrack()
        'm_objToolMain.F10ScrUpDate()
    End Sub

    Private Sub objTimer_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles objTimer.Tick
        m_objToolMain.F10ScrUpDate()
    End Sub

    '　　　　　　　　　　　　　　　　　　　開発中
    '↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑

    '---------------------------------------------
    '◆メニューストリップ
    '---------------------------------------------
    '●ファイル-終了
    Private Sub QuitXToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QuitXToolStripMenuItem.Click
        Me.Close()
    End Sub
    '●コントロール-停止
    Private Sub StopSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StopSToolStripMenuItem.Click
        m_objToolMain.objBtnPlay_Click()
    End Sub
    '●コントロール-次のトラック
    Private Sub NextTrackNToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NextTrackNToolStripMenuItem.Click
        m_objToolMain.objBtnNextTrack_Click()
    End Sub
    '●コントロール-前のトラック
    Private Sub BackTrackBToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BackTrackBToolStripMenuItem.Click
        m_objToolMain.objBtnBackTrack_Click()
    End Sub
    '●コントロール-ボリュームアップ
    Private Sub VolumeUpUToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VolumeUpUToolStripMenuItem.Click
        m_objToolMain.volume_ctrl(eVolumeCtrl.Up)
    End Sub
    '●コントロール-ボリュームダウン
    Private Sub VolumeDownDToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VolumeDownDToolStripMenuItem.Click
        m_objToolMain.volume_ctrl(eVolumeCtrl.Down)
    End Sub
    '●コントロール-ボリュームミュート
    Private Sub VolumeMuteMToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VolumeMuteMToolStripMenuItem.Click
        m_objToolMain.volume_ctrl(eVolumeCtrl.MUTE)
    End Sub
    '●表示-アートワーク
    Private Sub View_ArtWork_ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles View_ArtWork_ToolStripMenuItem.Click
        If View_ArtWork_ToolStripMenuItem.Checked = True Then
            View_ArtWork_ToolStripMenuItem.Checked = False
            objPnlArtWork.Visible = False
        Else
            View_ArtWork_ToolStripMenuItem.Checked = True
            objPnlArtWork.Visible = True
        End If
    End Sub

    '●ヘルプ-バージョン
    Private Sub VersionVToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VersionVToolStripMenuItem.Click
        frm99AboutBox.Show()
    End Sub

    '●仮想化ListViewにアイテムを追加
    Private Sub objListView_RetrieveVirtualItem(ByVal sender As Object, ByVal e As System.Windows.Forms.RetrieveVirtualItemEventArgs) Handles objListView.RetrieveVirtualItem
        m_objToolMain.objLV_RetVirtualItem(sender, e)
    End Sub

    '●アートワーク
    Private Sub objPictArtWork_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles objPictArtWork.DoubleClick

    End Sub

    Private Sub objListView_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles objListView.ColumnClick
        'Me.objListView.ListViewItemSorter = New ListViewItemComparer(e.Column)
        'とりあえず封印
    End Sub

    '●アートワーククリック
    Private Sub objPictArtWork_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles objPictArtWork.Click
        m_objToolMain.objArtWork_Click()
    End Sub
End Class

Class ListViewItemComparer
    Implements IComparer

    Private col As Integer

    Public Sub New()
        col = 0
    End Sub

    Public Sub New(ByVal column As Integer)
        col = column
    End Sub

    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare

        If col = 0 Or col = 4 Or col = 5 Then
            If CLng(CType(x, ListViewItem).SubItems(col).Text) > CLng(CType(y, ListViewItem).SubItems(col).Text) Then
                Return 1
            Else
                Return 0
            End If
        Else
            Return [String].Compare(CType(x, ListViewItem).SubItems(col).Text, CType(y, ListViewItem).SubItems(col).Text)
        End If

    End Function
End Class
