

Imports System.Data.SqlClient

Module Mod_SQLClient

#Region "SQL"

	Friend str_SQLPath As String = ""
	Friend str_SQLConn As String = ""

	Friend Function SQLReadQuery(comText As String, Optional Paramzz As List(Of SQLParamz) = Nothing) As DataTable
		If Paramzz Is Nothing Then Paramzz = New List(Of SQLParamz)()

		Using conX As New SqlConnection(str_SQLConn), comX As New SqlCommand(comText.Trim(), conX), dtx As New DataTable()
			If conX.State = ConnectionState.Closed Then conX.Open()

			If Paramzz.Count > 0 Then
				If Paramzz.Count > 0 Then
					For i = 0 To Paramzz.Count - 1
						If comText.Contains(Paramzz(i).Parame) Then comX.Parameters.AddWithValue(Paramzz(i).Parame, Paramzz(i).Value)
					Next
				End If
			End If

			Using readerX As SqlDataReader = comX.ExecuteReader()
				dtx.Load(readerX)
			End Using

			If conX.State = ConnectionState.Open Then conX.Close()
			GC.Collect()
			GC.WaitForPendingFinalizers()

			Return dtx
		End Using
	End Function

	Friend Sub SQLWriteQuery(comText As String, Optional Paramzz As List(Of SQLParamz) = Nothing, Optional useTransaction As Boolean = True)
		If Paramzz Is Nothing Then Paramzz = New List(Of SQLParamz)()
		Using conX As New SqlConnection(str_SQLConn), comX As New SqlCommand(comText.Trim(), conX)
			If conX.State = ConnectionState.Closed Then conX.Open()

			If Paramzz.Count > 0 Then
				If Paramzz.Count > 0 Then
					For i = 0 To Paramzz.Count - 1
						If comText.Contains(Paramzz(i).Parame) Then comX.Parameters.AddWithValue(Paramzz(i).Parame, Paramzz(i).Value)
					Next
				End If
			End If

			If useTransaction Then
				Using transX As SqlTransaction = conX.BeginTransaction
					comX.Transaction = transX
					Try
						comX.ExecuteNonQuery()
						transX.Commit()
						Console.WriteLine("SQLWriteQuery: Written!")
					Catch ex As Exception
						transX.Rollback()
						Console.WriteLine("SQLWriteQuery: (" & ex.Source & ")" & ex.Message)
						Throw
					End Try
				End Using
			Else
				comX.ExecuteNonQuery()
			End If

			If conX.State = ConnectionState.Open Then conX.Close()
			GC.Collect()
			GC.WaitForPendingFinalizers()
		End Using
	End Sub

	Friend Sub SQLDtToDb(DtSource As DataTable, SQLTable As String, SQLColumns() As String)
		Using conX As New SqlConnection(str_SQLConn), comX As New SqlCommand("", conX)
			If conX.State = ConnectionState.Closed Then conX.Open()
			Using transX As SqlTransaction = conX.BeginTransaction
				With comX
					.Transaction = transX
					Try
						If DtSource.Rows.Count = 0 Then
							.CommandText = "delete from " & SQLTable & ";"
							.ExecuteNonQuery()
						Else
							.CommandText = "delete from " & SQLTable & ";"
							.ExecuteNonQuery()

							Dim insertCount As Integer = 0
							Dim transacQuery As String = ""
							For Each dr As DataRow In DtSource.Rows
								insertCount += 1
								Dim valuez As New List(Of String)
								For Each dc As DataColumn In DtSource.Columns
									valuez.Add(dr.Item(dc).ToString.Replace("'", "''"))
								Next
								transacQuery &= Trim("insert into [" & SQLTable & "] ([" & String.Join("], [", SQLColumns) & "]) values ('" & String.Join("', '", valuez.ToArray) & "')") & ";" & str_newLine
								If insertCount = 777 Then
									.CommandText = transacQuery
									.ExecuteNonQuery()
									transacQuery = ""
									insertCount = 0
								End If
							Next
							If insertCount < 777 Then
								.CommandText = transacQuery
								.ExecuteNonQuery()
								transacQuery = ""
							End If
						End If

						transX.Commit()
						Console.WriteLine(SQLTable & ": Written @ " & DateTime.UtcNow.Subtract(dtm_ExecStart).ToString.Substring(0, 8))
					Catch ex As Exception
						transX.Rollback()
						Console.WriteLine(SQLTable & ": (" & ex.Source & ") " & ex.Message)
						Throw
					End Try
				End With
			End Using
		End Using
	End Sub

#End Region

End Module
