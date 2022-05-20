Imports System.Data.SqlClient
Imports System.Text

Public Class db_SHOHIN_IMG_FILE

    Public Shared Function EmptyTable(conStr As String) As DataTable
        Dim sb As StringBuilder = New StringBuilder(1000)
        sb.AppendLine("SELECT TOP(0) *")
        sb.AppendLine("FROM SHOHIN_IMG_FILE;")

        Dim cn As SqlConnection = New SqlConnection(conStr)
        Dim cmd As SqlCommand = New SqlCommand(sb.ToString(), cn)

        Dim adp As SqlDataAdapter = New SqlDataAdapter(cmd)
        Dim tbl As DataTable = New DataTable()
        adp.FillSchema(tbl, SchemaType.Mapped)

        adp.Dispose()
        cmd.Dispose()
        cn.Dispose()

        Return tbl
    End Function

    Public Shared Sub Update_01(conStr As String, tbl As DataTable)
        Dim cn As SqlConnection = New SqlConnection(conStr)
        Try
            cn.Open()
        Catch
            Throw
        End Try

        Dim trn As SqlTransaction
        Try
            trn = cn.BeginTransaction()
        Catch
            cn.Close()
            cn.Dispose()

            Throw
        End Try

        Dim sb As StringBuilder = New StringBuilder(10000000)
        Dim cmd As SqlCommand = Nothing

        Try
            sb.Clear()
            sb.AppendLine("DELETE FROM SHOHIN_IMG_FILE_HISTORY")
            sb.AppendLine("WHERE SIF_YMD < DATEADD(day, -20, CAST(GETDATE() AS date));")
            sb.AppendLine("INSERT INTO SHOHIN_IMG_FILE_HISTORY")
            sb.AppendLine("SELECT *")
            sb.AppendLine("FROM SHOHIN_IMG_FILE;")
            sb.AppendLine("TRUNCATE TABLE SHOHIN_IMG_FILE;")

            cmd = New SqlCommand(sb.ToString(), cn) With {
                .Transaction = trn,
                .CommandTimeout = 600
            }
            cmd.ExecuteNonQuery()
        Catch
            trn.Rollback()

            trn.Dispose()
            cn.Close()
            cn.Dispose()

            Throw
        Finally
            If cmd IsNot Nothing Then
                cmd.Dispose()
                cmd = Nothing
            End If
        End Try

        Try
            Dim row1 As DataRow
            Dim idx1 As Integer, idx2 As Integer, idx3 As Integer

            For idx1 = 0 To tbl.Rows.Count - 1
                sb.Clear()

                For idx2 = idx1 To If(idx1 + 9999 >= tbl.Rows.Count, tbl.Rows.Count - 1, idx1 + 9999)
                    row1 = tbl.Rows(idx2)

                    sb.Append("INSERT INTO SHOHIN_IMG_FILE VALUES (")
                    For idx3 = 0 To tbl.Columns.Count - 2
                        sb.Append(dbValToText(row1(idx3)) + ", ")
                    Next
                    sb.AppendLine(dbValToText(row1(idx3)) + ");")
                Next

                cmd = New SqlCommand(sb.ToString(), cn) With {
                    .Transaction = trn,
                    .CommandTimeout = 600,
                    .CommandType = CommandType.Text,
                    .CommandText = sb.ToString()
                }
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                cmd = Nothing

                idx1 = idx2 - 1
            Next

            trn.Commit()
        Catch
            trn.Rollback()

            Throw New ApplicationException("テーブル [SHOHIN_IMG_FILE] に、レコードを挿入出来ませんでした。")
        Finally
            If cmd IsNot Nothing Then
                cmd.Dispose()
            End If

            trn.Dispose()
            cn.Close()
            cn.Dispose()
        End Try
    End Sub

    Public Shared Function Get_ShohinCd_01(conStr As String) As DataTable
        Dim sb As StringBuilder = New StringBuilder(1000)
        sb.AppendLine("SELECT DISTINCT SSHM_SCD, ISNULL(SSHM_SFILENAME, '') AS SSHM_SFILENAME")
        sb.AppendLine("FROM ft_SHOHIN_MST(NULL)")
        sb.AppendLine("ORDER BY SSHM_SCD, SSHM_SFILENAME;")

        Dim cn As SqlConnection = New SqlConnection(conStr)
        Dim cmd As SqlCommand = New SqlCommand(sb.ToString(), cn)

        Dim adp As SqlDataAdapter = New SqlDataAdapter(cmd)
        Dim tbl As DataTable = New DataTable()
        adp.Fill(tbl)

        adp.Dispose()
        cmd.Dispose()
        cn.Dispose()

        Return tbl
    End Function

    Public Shared Function dbValToText(dbVal As Object) As String
        Select Case dbVal.GetType()
            Case GetType(String)
                Return "'" + DirectCast(dbVal, String).Replace("'", "''") + "'"
            Case GetType(Char)
                Return DirectCast(dbVal, Char).ToString()
            Case GetType(Boolean)
                Return If(DirectCast(dbVal, Boolean), "1", "0")
            Case GetType(SByte)
                Return DirectCast(dbVal, SByte).ToString("d")
            Case GetType(Short)
                Return DirectCast(dbVal, Short).ToString("d")
            Case GetType(Integer)
                Return DirectCast(dbVal, Integer).ToString("d")
            Case GetType(Long)
                Return DirectCast(dbVal, Long).ToString("d")
            Case GetType(Byte)
                Return DirectCast(dbVal, Byte).ToString("d")
            Case GetType(UShort)
                Return DirectCast(dbVal, UShort).ToString("d")
            Case GetType(UInteger)
                Return DirectCast(dbVal, UInteger).ToString("d")
            Case GetType(ULong)
                Return DirectCast(dbVal, ULong).ToString("d")
            Case GetType(Decimal)
                Return DirectCast(dbVal, Decimal).ToString("f4")
            Case GetType(Single)
                Return DirectCast(dbVal, Single).ToString("f4")
            Case GetType(Double)
                Return DirectCast(dbVal, Double).ToString("f4")
            Case GetType(DateTime)
                Return "'" + DirectCast(dbVal, DateTime).ToString("yyyy-MM-dd HH:mm:ss") + "'"
            Case GetType(TimeSpan)
                Return "'" + DirectCast(dbVal, TimeSpan).ToString("HH:mm:ss.fff") + "'"
            Case GetType(Guid)
                Return "'" + DirectCast(dbVal, Guid).ToString() + "'"
            Case GetType(DBNull)
                Return "NULL"
            Case Else
                Throw New SqlTypes.SqlTypeException("取り扱われていないデータ型です。（" + dbVal.GetType().ToString() + "）")
        End Select
    End Function

End Class
