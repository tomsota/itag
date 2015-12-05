'---------------------------------------------
'[Name]     modComMm
'[J.Name]   共有メモリ定義
'[Date]     2010/10/20
'[Func]     共有メモリ
'[Memo]
'---------------------------------------------
Module modComMm

    '●グローバルクラス
    Public p_objErrCtrl As cls99ErrCtrl         'エラー制御クラス

    '●データ部
    Public Structure strctData
        Dim strData() As String
        Dim chkbox() As CheckBox
    End Structure

    '●データ部配列
    Public p_strctData() As strctData

    '●MySetting配列
    Public p_objMySetting(9) As System.Collections.Specialized.StringCollection

End Module
