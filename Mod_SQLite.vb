Imports System.Data.SQLite
Imports System.Text

Module Mod_SQLite

    Friend str_SQLitePath As String = ""
    Friend str_SQLiteConn As String = ""

    Friend Function SQLReadQuery(comText As String, Optional Paramzz As List(Of SQLParamz) = Nothing) As DataTable
        Dim dtx As New DataTable()
        If Paramzz Is Nothing Then Paramzz = New List(Of SQLParamz)()

        Using conX As New SQLiteConnection(str_SQLiteConn), comX As New SQLiteCommand(comText.Trim(), conX)
            conX.Open()

            ' Add parameters
            For Each param In Paramzz
                comX.Parameters.AddWithValue(param.Parame, param.Value)
            Next

            ' Execute query and load data
            Using readerX As SQLiteDataReader = comX.ExecuteReader()
                dtx.Load(readerX)
            End Using
        End Using

        Return dtx
    End Function

    Friend Sub SQLWriteQuery(comText As String, Optional Paramzz As List(Of SQLParamz) = Nothing, Optional useTransaction As Boolean = True)
        If Paramzz Is Nothing Then Paramzz = New List(Of SQLParamz)()

        Using conX As New SQLiteConnection(str_SQLiteConn), comX As New SQLiteCommand(comText.Trim(), conX)
            conX.Open()

            ' Add parameters
            For Each param In Paramzz
                comX.Parameters.AddWithValue(param.Parame, param.Value)
            Next

            If useTransaction Then
                Using transX As SQLiteTransaction = conX.BeginTransaction()
                    comX.Transaction = transX
                    Try
                        comX.ExecuteNonQuery()
                        transX.Commit()
                        Console.WriteLine("SQLWriteQuery: Written!")
                    Catch ex As Exception
                        transX.Rollback()
                        Console.WriteLine("SQLWriteQuery Error: " & ex.Message)
                        Throw
                    End Try
                End Using
            Else
                comX.ExecuteNonQuery()
            End If
        End Using
    End Sub

    Friend Sub SQLDtToDb(DtSource As DataTable, SQLTable As String, SQLColumns() As String)
        Using conX As New SQLiteConnection(str_SQLiteConn), comX As New SQLiteCommand("", conX)
            conX.Open()
            Using transX As SQLiteTransaction = conX.BeginTransaction()
                comX.Transaction = transX
                Try
                    comX.CommandText = $"DELETE FROM {SQLTable};"
                    comX.ExecuteNonQuery()

                    If DtSource.Rows.Count > 0 Then
                        Dim sb As New StringBuilder()
                        Dim insertCount As Integer = 0

                        For Each dr As DataRow In DtSource.Rows
                            insertCount += 1
                            Dim values As List(Of String) = New List(Of String)

                            For Each dc As DataColumn In DtSource.Columns
                                values.Add("'" & dr.Item(dc).ToString().Replace("'", "''") & "'")
                            Next

                            sb.AppendLine($"INSERT INTO [{SQLTable}] ([{String.Join("], [", SQLColumns)}]) VALUES ({String.Join(", ", values)});")

                            ' Execute in batches
                            If insertCount = 777 Then
                                comX.CommandText = sb.ToString()
                                comX.ExecuteNonQuery()
                                sb.Clear()
                                insertCount = 0
                            End If
                        Next

                        ' Execute remaining queries
                        If sb.Length > 0 Then
                            comX.CommandText = sb.ToString()
                            comX.ExecuteNonQuery()
                        End If
                    End If

                    transX.Commit()
                    Console.WriteLine($"{SQLTable}: Written @ {DateTime.UtcNow:HH:mm:ss}")
                Catch ex As Exception
                    transX.Rollback()
                    Console.WriteLine($"{SQLTable} Error: {ex.Message}")
                    Throw
                End Try
            End Using
        End Using
    End Sub

End Module

Friend Class SQLParamz
    Friend Property Parame As String
    Friend Property Value As Object
End Class
