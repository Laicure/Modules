Imports System.Data.SQLite

Module Mod_SQLite

    Friend str_SQLitePath As String = ""
    Friend str_SQLiteConn As String = ""

    Friend Function SQLReadQuery(comText As String, Optional Paramzz As List(Of SQLParamz) = Nothing) As DataTable
        If Paramzz Is Nothing Then Paramzz = New List(Of SQLParamz)()

        Using conX As New SQLiteConnection(str_SQLiteConn),
              comX As New SQLiteCommand(comText.Trim(), conX),
              dtx As New DataTable()

            If conX.State = ConnectionState.Closed Then conX.Open()

            ' Add parameters if present
            For Each param In Paramzz
                If comText.Contains(param.Parame) Then
                    comX.Parameters.AddWithValue(param.Parame, param.Value)
                End If
            Next

            Using readerX As SQLiteDataReader = comX.ExecuteReader()
                dtx.Load(readerX)
            End Using

            Return dtx
        End Using
    End Function

    Friend Sub SQLWriteQuery(comText As String, Optional Paramzz As List(Of SQLParamz) = Nothing, Optional useTransaction As Boolean = True)
        If Paramzz Is Nothing Then Paramzz = New List(Of SQLParamz)()

        Using conX As New SQLiteConnection(str_SQLiteConn),
              comX As New SQLiteCommand(comText.Trim(), conX)
            
            If conX.State = ConnectionState.Closed Then conX.Open()

            ' Add parameters if present
            For Each param In Paramzz
                If comText.Contains(param.Parame) Then
                    comX.Parameters.AddWithValue(param.Parame, param.Value)
                End If
            Next

            If useTransaction Then
                Using transX As SQLiteTransaction = conX.BeginTransaction()
                    Try
                        comX.Transaction = transX
                        comX.ExecuteNonQuery()
                        transX.Commit()
                        Console.WriteLine("SQLWriteQuery: Written!")
                    Catch ex As Exception
                        transX.Rollback()
                        Console.WriteLine($"SQLWriteQuery Error: {ex.Message}")
                        Throw
                    End Try
                End Using
            Else
                comX.ExecuteNonQuery()
            End If
        End Using
    End Sub

    Friend Sub SQLDtToDb(DtSource As DataTable, SQLTable As String, SQLColumns() As String)
        Using conX As New SQLiteConnection(str_SQLiteConn),
              comX As New SQLiteCommand("", conX)
            
            If conX.State = ConnectionState.Closed Then conX.Open()
            Using transX As SQLiteTransaction = conX.BeginTransaction()
                Try
                    comX.Transaction = transX
                    comX.CommandText = $"DELETE FROM [{SQLTable}];"
                    comX.ExecuteNonQuery()

                    If DtSource.Rows.Count > 0 Then
                        Dim insertCount As Integer = 0
                        Dim transacQuery As New Text.StringBuilder()

                        For Each dr As DataRow In DtSource.Rows
                            insertCount += 1
                            Dim valuez = DtSource.Columns.Cast(Of DataColumn)()
                                        .Select(Function(dc) dr(dc).ToString().Replace("'", "''"))
                                        .ToArray()

                            transacQuery.AppendLine($"INSERT INTO [{SQLTable}] ([{String.Join("], [", SQLColumns)}]) VALUES ('{String.Join("', '", valuez)}');")

                            If insertCount = 777 Then
                                comX.CommandText = transacQuery.ToString()
                                comX.ExecuteNonQuery()
                                transacQuery.Clear()
                                insertCount = 0
                            End If
                        Next

                        If transacQuery.Length > 0 Then
                            comX.CommandText = transacQuery.ToString()
                            comX.ExecuteNonQuery()
                        End If
                    End If

                    transX.Commit()
                    Console.WriteLine($"{SQLTable}: Written @ {DateTime.UtcNow.Subtract(dtm_ExecStart).ToString().Substring(0, 8)}")
                Catch ex As Exception
                    transX.Rollback()
                    Console.WriteLine($"{SQLTable}: ({ex.Source}) {ex.Message}")
                    Throw
                End Try
            End Using
        End Using
    End Sub

End Module
