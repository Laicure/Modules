Imports System.Data.SqlClient
Imports System.Text

Module Mod_SQLClient

    Friend str_SQLConn As String = ""

    ' Function to read SQL query and return DataTable
    Friend Function Func_SQLReadQuery(comText As String, Optional Paramzz As List(Of SQLParamz) = Nothing) As DataTable
        If Paramzz Is Nothing Then Paramzz = New List(Of SQLParamz)()

        Using conX As New SqlConnection(str_SQLConn), comX As New SqlCommand(comText.Trim(), conX), dtx As New DataTable()
            conX.Open()

            ' Add parameters
            For Each param In Paramzz
                If comText.Contains(param.Parame) Then
                    comX.Parameters.AddWithValue(param.Parame, param.Value)
                End If
            Next

            ' Execute query
            Using readerX As SqlDataReader = comX.ExecuteReader()
                dtx.Load(readerX)
            End Using

            Return dtx
        End Using
    End Function

    ' Subroutine to execute a write query (INSERT, UPDATE, DELETE)
    Friend Sub Sub_SQLWriteQuery(comText As String, Optional Paramzz As List(Of SQLParamz) = Nothing, Optional useTransaction As Boolean = True)
        If Paramzz Is Nothing Then Paramzz = New List(Of SQLParamz)()

        Using conX As New SqlConnection(str_SQLConn), comX As New SqlCommand(comText.Trim(), conX)
            conX.Open()

            ' Add parameters
            For Each param In Paramzz
                If comText.Contains(param.Parame) Then
                    comX.Parameters.AddWithValue(param.Parame, param.Value)
                End If
            Next

            ' Use transaction if specified
            If useTransaction Then
                Using transX As SqlTransaction = conX.BeginTransaction()
                    comX.Transaction = transX
                    Try
                        comX.ExecuteNonQuery()
                        transX.Commit()
                        Console.WriteLine("SQLWriteQuery: Success")
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

    ' Subroutine to insert a DataTable into a SQL table
    Friend Sub SQLDtToDb(DtSource As DataTable, SQLTable As String, SQLColumns() As String)
        Using conX As New SqlConnection(str_SQLConn), comX As New SqlCommand("", conX)
            conX.Open()
            Using transX As SqlTransaction = conX.BeginTransaction()
                comX.Transaction = transX
                Try
                    ' Clear table before inserting new data
                    comX.CommandText = $"DELETE FROM {SQLTable};"
                    comX.ExecuteNonQuery()

                    If DtSource.Rows.Count > 0 Then
                        Dim insertCount As Integer = 0
                        Dim transacQuery As New StringBuilder()
                        Dim paramList As New List(Of SQLParamz)()

                        For Each dr As DataRow In DtSource.Rows
                            insertCount += 1

                            Dim paramNames As New List(Of String)
                            For colIndex As Integer = 0 To DtSource.Columns.Count - 1
                                Dim paramName As String = $"@p{insertCount}_{colIndex}"
                                paramNames.Add(paramName)
                                paramList.Add(New SQLParamz With {.Parame = paramName, .Value = dr(colIndex)})
                            Next

                            transacQuery.AppendLine($"INSERT INTO [{SQLTable}] ([{String.Join("], [", SQLColumns)}]) VALUES ({String.Join(", ", paramNames)});")

                            ' Execute batch insert every 777 records
                            If insertCount = 777 Then
                                ExecuteBatchInsert(comX, transacQuery, paramList)
                                transacQuery.Clear()
                                paramList.Clear()
                                insertCount = 0
                            End If
                        Next

                        ' Execute any remaining inserts
                        If insertCount > 0 Then
                            ExecuteBatchInsert(comX, transacQuery, paramList)
                        End If
                    End If

                    transX.Commit()
                    Console.WriteLine($"{SQLTable}: Data inserted @ {DateTime.UtcNow.Subtract(dtm_ExecStart).ToString.Substring(0, 8)}")

                Catch ex As Exception
                    transX.Rollback()
                    Console.WriteLine($"{SQLTable}: Error ({ex.Source}) {ex.Message}")
                    Throw
                End Try
            End Using
        End Using
    End Sub

    ' Helper method to execute batch insert
    Private Sub ExecuteBatchInsert(comX As SqlCommand, transacQuery As StringBuilder, paramList As List(Of SQLParamz))
        comX.CommandText = transacQuery.ToString()
        comX.Parameters.Clear()

        For Each param In paramList
            comX.Parameters.AddWithValue(param.Parame, param.Value)
        Next

        comX.ExecuteNonQuery()
    End Sub

    ' Class to handle SQL parameters
    Friend Class SQLParamz
        Friend Property Parame As String
        Friend Property Value As Object
    End Class

End Module
