'---------------------------------------------
'[Name]     modToolDef
'[J.Name]   ツール固有定義
'[Date]     2010/10/20
'[Func]     固定値定義
'[Memo]
'---------------------------------------------
Module modToolDef

    '●リターンコード定義
    Public Enum eReturnCode
        OK
        NG
    End Enum

    '●アップダウン定義
    Public Enum eVolumeCtrl
        MUTE
        Up
        Down
        SETVOLUME
    End Enum

    '●PLAYボタンテキスト定義
    Public Const c_BtnPlay_Play As String = "PLAY"
    Public Const c_BtnPlay_Stop As String = "STOP"

    '●メニューストリップテキスト定義
    Public Const c_MnCtrl_Play As String = "Play(&P)"
    Public Const c_MnCtrl_Stop As String = "Stop(&S)"

End Module
