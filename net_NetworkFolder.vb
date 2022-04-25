Imports System.Runtime.InteropServices

Public Module net_NetworkFolder

    Public Declare Unicode Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2W" (
        ByRef lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Integer) As Integer

    Public Declare Unicode Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2W" (
        ByVal lpName As String, ByVal dwFlags As Integer, ByVal fForce As Integer) As Integer

    Public Const RESOURCE_CONNECTED As Integer = &H1
    Public Const RESOURCE_GLOBALNET As Integer = &H2
    Public Const RESOURCE_CONTEXT As Integer = &H5

    Public Const RESOURCETYPE_DISK As Integer = &H1
    Public Const RESOURCETYPE_PRINT As Integer = &H2
    Public Const RESOURCETYPE_ANY As Integer = &H0

    Public Const RESOURCEDISPLAYTYPE_NETWORK As Integer = &H6
    Public Const RESOURCEDISPLAYTYPE_DOMAIN As Integer = &H1
    Public Const RESOURCEDISPLAYTYPE_SERVER As Integer = &H2
    Public Const RESOURCEDISPLAYTYPE_SHARE As Integer = &H3
    Public Const RESOURCEDISPLAYTYPE_DIRECTORY As Integer = &H9
    Public Const RESOURCEDISPLAYTYPE_GENERIC As Integer = &H0

    Public Const RESOURCEUSAGE_CONNECTABLE As Integer = &H1
    Public Const RESOURCEUSAGE_CONTAINER As Integer = &H2

    'ネットワークエラー定数
    Public Const ERROR_SUCCESS As Integer = 0                                  '正常終了
    Public Const ERROR_BAD_NETPATH As Integer = 53                             'ネットワークパスが不正
    Public Const ERROR_ACCESS_DENIED As Integer = 8                            'ネットワーク資源へのアクセスが拒否されました｡
    Public Const ERROR_ALREADY_ASSIGNED As Integer = 85                        'lpLocalName で指定したローカルデバイスは既にネットワーク資源に接続されています。
    Public Const ERROR_BAD_DEV_TYPE As Integer = 66                            'ローカルデバイスの種類とネットワーク資源の種類が一致しません｡
    Public Const ERROR_BAD_DEVICE As Integer = 1200                            'lpLocalName で指定した値が無効です。
    Public Const ERROR_BAD_NET_NAME As Integer = 67                            'lpRemoteName で指定した値を、どのネットワーク資源のプロバイダも受け付けません。資源の名前が無効か、指定した資源が見つかりません。
    Public Const ERROR_BAD_PROFILE As Integer = 1206                           'ユーザープロファイルの形式が正しくありません｡
    Public Const ERROR_BAD_PROVIDER As Integer = 1204                          'lpProvider で指定した値がどのプロバイダとも一致しません。
    Public Const ERROR_BUSY As Integer = 170                                   'ルーターまたはプロバイダがビジー（ おそらく初期化中）です。この関数をもう一度呼び出してください。
    Public Const ERROR_CANCELLED As Integer = 1223                             'ネットワーク資源のプロバイダのいずれかでユーザーがダイアログボックスを使って接続操作を取り消したか､接続先の資源が接続操作を取り消しました｡
    Public Const ERROR_CANNOT_OPEN_PROFILE As Integer = 1205                   '恒久的な接続を処理するためのユーザープロファイルを開くことができません｡
    Public Const ERROR_DEVICE_ALREADY_REMEMBERED As Integer = 1202             'lpLocalName で指定したデバイスのエントリは既にユーザープロファイル内に存在します。
    Public Const ERROR_EXTENDED_ERROR As Integer = 1208                        'ネットワーク固有のエラーが発生しました。エラーの説明を取得するには、WNetGetLastError 関数を使います。
    Public Const ERROR_INVALID_PASSWORD As Integer = 86                        '指定したパスワードが無効です｡
    Public Const ERROR_NO_NET_OR_BAD_PATH As Integer = 1203                    'ネットワークコンポーネントが開始されていないか､指定した名前が利用できないために､操作を行えませんでした｡
    Public Const ERROR_NO_NETWORK As Integer = 1222                            'ネットワークに接続されていません｡
    Public Const ERROR_DEVICE_IN_USE As Integer = 2404                         '指定したデバイスがアクティブなプロセスによって使用中のため､切断できません｡
    Public Const ERROR_NOT_CONNECTED As Integer = 2250                         'lpName パラメータで指定した名前がリダイレクトされているデバイスを表していないか、lpName で指定したデバイスにシステムが接続していません。
    Public Const ERROR_OPEN_FILES As Integer = 2401                            '開いているファイルがあり、fForce が FALSE です。

    <StructLayout(LayoutKind.Sequential, CharSet:=System.Runtime.InteropServices.CharSet.Unicode)>
    Public Structure NETRESOURCE
        Public dwScope As Integer
        Public dwType As Integer
        Public dwDisplayType As Integer
        Public dwUsage As Integer
        Public lpLocalName As String
        Public lpRemoteName As String
        Public lpComment As String
        Public lpProvider As String
    End Structure

    Private Function getErrMsg(err As Integer) As String
        Dim msg As String

        Select Case err
            Case ERROR_SUCCESS
                msg = "正常終了"
            Case ERROR_BAD_NETPATH
                msg = "ネットワークパスが不正です。"
            Case ERROR_ACCESS_DENIED
                msg = "ネットワーク資源へのアクセスが拒否されました。"
            Case ERROR_ALREADY_ASSIGNED
                msg = "lpLocalName で指定したローカルデバイスは既にネットワーク資源に接続されています。"
            Case ERROR_BAD_DEV_TYPE
                msg = "ローカルデバイスの種類とネットワーク資源の種類が一致しません。"
            Case ERROR_BAD_DEVICE
                msg = "lpLocalName で指定した値が無効です。"
            Case ERROR_BAD_NET_NAME
                msg = "lpRemoteName で指定した値を、どのネットワーク資源のプロバイダも受け付けません。資源の名前が無効か、指定した資源が見つかりません。"
            Case ERROR_BAD_PROFILE
                msg = "ユーザープロファイルの形式が正しくありません｡"
            Case ERROR_BAD_PROVIDER
                msg = "lpProvider で指定した値がどのプロバイダとも一致しません。"
            Case ERROR_BUSY
                msg = "ルーターまたはプロバイダがビジー（ おそらく初期化中）です。この関数をもう一度呼び出してください。"
            Case ERROR_CANCELLED
                msg = "ネットワーク資源のプロバイダのいずれかでユーザーがダイアログボックスを使って接続操作を取り消したか､接続先の資源が接続操作を取り消しました｡"
            Case ERROR_CANNOT_OPEN_PROFILE
                msg = "恒久的な接続を処理するためのユーザープロファイルを開くことができません｡"
            Case ERROR_DEVICE_ALREADY_REMEMBERED
                msg = "lpLocalName で指定したデバイスのエントリは既にユーザープロファイル内に存在します。"
            Case ERROR_EXTENDED_ERROR
                msg = "ネットワーク固有のエラーが発生しました。エラーの説明を取得するには、WNetGetLastError 関数を使います。"
            Case ERROR_INVALID_PASSWORD
                msg = "指定したパスワードが無効です｡"
            Case ERROR_NO_NET_OR_BAD_PATH
                msg = "ネットワークコンポーネントが開始されていないか､指定した名前が利用できないために､操作を行えませんでした｡"
            Case ERROR_NO_NETWORK
                msg = "ネットワークに接続されていません｡"
            Case ERROR_DEVICE_IN_USE
                msg = "指定したデバイスがアクティブなプロセスによって使用中のため､切断できません｡"
            Case ERROR_NOT_CONNECTED
                msg = "lpName パラメータで指定した名前がリダイレクトされているデバイスを表していないか、lpName で指定したデバイスにシステムが接続していません。"
            Case ERROR_OPEN_FILES
                msg = "開いているファイルがあり、fForce が FALSE です。"
            Case Else
                msg = "ネットワークに関する、原因不明なエラーです。"
        End Select

        Return msg
    End Function

    Public Sub ConnectNetworkFolder(path As String, user As String, pass As String)
        If System.IO.Directory.Exists(path) Then
            Return
        End If

        Dim nr As NETRESOURCE = New NETRESOURCE()
        With nr
            .dwScope = RESOURCE_GLOBALNET
            .dwType = RESOURCETYPE_DISK
            .dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
            .dwUsage = 0
            .lpLocalName = Nothing
            .lpRemoteName = path
            .lpComment = Nothing
            .lpProvider = Nothing
        End With

        Dim errId As Integer

#If False Then
        errId = WNetCancelConnection2(System.IO.Path.GetPathRoot(path), 0, True)
        If errId <> ERROR_SUCCESS Then
            Throw New ApplicationException(getErrMsg(errId))
        End If
#End If

        errId = WNetAddConnection2(nr, pass, user, 0)
        If errId <> ERROR_SUCCESS Then
            Throw New ApplicationException(getErrMsg(errId))
        End If
    End Sub

End Module
