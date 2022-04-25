Imports System.Xml

Public Class SettingConfig
    Protected Shared _root As XmlNode

    Shared Sub New()
        Try
            Dim fn As String
            fn = System.Environment.CurrentDirectory + "\Setting.config"

            Dim doc As XmlDocument = New XmlDocument
            doc.Load(fn)

            _root = doc.SelectSingleNode("/configuration")
        Catch ex As Exception
            Throw New ApplicationException("Iregular Setting.config File.", ex)
        End Try
    End Sub

    Public Shared Function FilePrefix(name As String) As String
        Dim nd As XmlNode = _root.SelectSingleNode(String.Format("./file_prefix/{0}", name))
        If nd Is Nothing Then
            Return Nothing
        End If

        Return nd.InnerText.Trim()
    End Function

    Public Shared ReadOnly Property DbList As XmlNodeList
        Get
            Return _root.SelectNodes("./db")
        End Get
    End Property

    Public Shared Function Db_Name(nd_db As XmlNode) As String
        Return nd_db.Attributes("name").Value
    End Function

    Public Shared Function Db_ConStr(nd_db As XmlNode) As String
        Return nd_db.Attributes("constr").Value
    End Function

    Public Shared Function Db_PicUnitList(nd_db As XmlNode) As XmlNodeList
        Return nd_db.SelectNodes("./pic_unit")
    End Function

    Public Shared Function Db_PicUnit_Name(nd_picUnit As XmlNode) As String
        Return nd_picUnit.Attributes("name").Value
    End Function

    Public Shared Function Db_PicUnit_DirList(nd_check As XmlNode) As XmlNodeList
        Return nd_check.SelectNodes("./dir")
    End Function

    Public Shared Function Db_PicUnit_Dir_Path(nd_dir As XmlNode) As String
        Return nd_dir.Attributes("path").Value
    End Function

    Public Shared Function Db_PicUnit_Dir_User(nd_dir As XmlNode) As String
        Return nd_dir.Attributes("user").Value
    End Function

    Public Shared Function Db_PicUnit_Dir_Pass(nd_dir As XmlNode) As String
        Return nd_dir.Attributes("pass").Value
    End Function

    ''' <summary></summary>
    ''' <param name="nd_dir"></param>
    ''' <returns>
    ''' 2次インデックス
    '''     0：イメージタイプ（IMG, DTL_IMG, KLS_IMG）
    '''     1：ファイル名前置詞（, dtl_, kls_）
    ''' </returns>
    Public Shared Function Db_PicUnit_Dir_Type(nd_dir As XmlNode) As String(,)
        Dim ndLs As XmlNodeList = nd_dir.SelectNodes("./type")
        Dim ary(ndLs.Count - 1, 1) As String

        For a As Integer = 0 To ndLs.Count - 1
            ary(a, 0) = ndLs(a).InnerText.Trim()
            ary(a, 1) = FilePrefix(ary(a, 0))
        Next

        Return ary
    End Function

    Public Shared ReadOnly Property WindowAutoClose As Boolean
        Get
            Return Boolean.Parse(_root.SelectSingleNode("./window_auto_close").InnerText.Trim())
        End Get
    End Property
End Class
