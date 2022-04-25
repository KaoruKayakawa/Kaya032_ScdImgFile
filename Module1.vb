Imports System.IO
Imports System.Xml

Module Module1

    Private _computer As String = System.Environment.MachineName
    Private _loginUser As String = System.Environment.UserName

    Private _conStr As String
    Private _resTbl As DataTable
    Private _shoTbl As DataTable

    Sub Main()
        Try
            Console.WriteLine("■ ＤＢテーブル [SHOHIN_IMG_FILE] の更新（ ver：" + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString() + "）")

            For Each db As XmlNode In SettingConfig.DbList
                Console.WriteLine("")
                Console.WriteLine("　>> ＤＢ：" + SettingConfig.Db_Name(db))

                _conStr = SettingConfig.Db_ConStr(db)
                _resTbl = db_SHOHIN_IMG_FILE.EmptyTable(_conStr)
                _shoTbl = db_SHOHIN_IMG_FILE.Get_ShohinCd_01(_conStr)

                For Each chk As XmlNode In SettingConfig.Db_PicUnitList(db)
                    Console.WriteLine("　>>>> チェック名称：" + SettingConfig.Db_PicUnit_Name(chk))
                    Console.WriteLine("　>>>>>> 開始（" + DateTime.Now.ToString("HH:mm:ss") + "）")
                    exeCheck(chk)
                    Console.WriteLine("　>>>>>> 終了（" + DateTime.Now.ToString("HH:mm:ss") + "）")
                Next

                Console.WriteLine("　>>>> ＤＢ更新")
                Console.WriteLine("　>>>>>> 開始（" + DateTime.Now.ToString("HH:mm:ss") + "）")
                db_SHOHIN_IMG_FILE.Update_01(_conStr, _resTbl)
                Console.WriteLine("　>>>>>> 終了（" + DateTime.Now.ToString("HH:mm:ss") + "）")
            Next
        Catch ex As Exception
            Console.WriteLine("")
            Console.WriteLine("** エラー **" + vbCrLf + ex.Message)
        Finally
            If Not SettingConfig.WindowAutoClose Then
                Console.WriteLine("")
                Console.WriteLine("** [Enter] 入力で、ウインドウが閉じます。 **")
                Console.ReadLine()
            End If
        End Try
    End Sub

    Private Sub exeCheck(chk As XmlNode)
        For Each dir As XmlNode In SettingConfig.Db_PicUnit_DirList(chk)
            Dim dirPath As String = SettingConfig.Db_PicUnit_Dir_Path(dir)
            Dim user As String = SettingConfig.Db_PicUnit_Dir_User(dir), pass As String = SettingConfig.Db_PicUnit_Dir_Pass(dir)

            Dim di As DirectoryInfo = New DirectoryInfo(dirPath)
            If Not di.Exists Then
                net_NetworkFolder.ConnectNetworkFolder(di.FullName, user, pass)
                di.Refresh()
            End If

            Dim ftLs As String(,) = SettingConfig.Db_PicUnit_Dir_Type(dir)
            Dim fiLs As FileInfo()
            Dim row As DataRow
            Dim udt As DateTime = DateTime.Now

            For Each shoCdFn As DataRow In _shoTbl.Rows
                For a As Integer = 0 To ftLs.GetLength(0) - 1
                    fiLs = di.GetFiles(ftLs(a, 1) + DirectCast(shoCdFn("SSHM_SFILENAME"), String))

                    row = _resTbl.NewRow()
                    row("SIF_PIC_UNIT") = SettingConfig.Db_PicUnit_Name(chk)
                    row("SIF_SCD") = shoCdFn("SSHM_SCD")
                    row("SIF_SFILENAME") = shoCdFn("SSHM_SFILENAME")
                    row("SIF_IMG_TYPE") = ftLs(a, 0)
                    row("SIF_IMG_PATH") = dirPath
                    If fiLs.Length = 0 Then
                        row("SIF_IMG_EXIST") = 0
                        row("SIF_FILE_CREATE") = DBNull.Value
                        row("SIF_FILE_LASTWRITE") = DBNull.Value
                    Else
                        row("SIF_IMG_EXIST") = 1
                        row("SIF_FILE_CREATE") = fiLs(0).CreationTime
                        row("SIF_FILE_LASTWRITE") = fiLs(0).LastWriteTime
                    End If
                    row("SIF_YMD") = udt
                    row("SIF_COMPUTER") = _computer
                    row("SIF_USER") = _loginUser

                    _resTbl.Rows.Add(row)
                Next
            Next
        Next
    End Sub

End Module
