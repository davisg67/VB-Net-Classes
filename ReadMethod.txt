Imports Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports System.Data
Imports ErrorHandling

Public Class ReadMethod
    Protected _ID As String = String.Empty
    Protected _Method As String = String.Empty
    Protected _MfrID As String = String.Empty

    Public ReadOnly Property ID_Field() As String
        Get
            Return _ID
        End Get
    End Property

    Public Property Method_Field() As String
        Get
            Return _Method
        End Get
        Set(value As String)
            _Method = value
        End Set
    End Property

    Public Property MfrID_Field() As String
        Get
            Return _MfrID
        End Get
        Set(value As String)
            _MfrID = value
        End Set
    End Property


    Public Function Init(ID As String) As String
        If Not IsNumeric(ID) Then
            ErrorHandler.WritelogFile("Global", "ReadMethod(class)", "Init", "The ID value sent was not a number.")
            Return "Error"
        End If

        Try
            Dim Result As String = String.Empty
            Dim conString = ConfigurationManager.ConnectionStrings("MSIUBDAPPCodeCon").ConnectionString
            Dim queryString = "SELECT * FROM parts_read_method WHERE id = " & ID

            Dim connection As New SqlConnection(conString)


            Using (connection)
                Dim command = New SqlCommand(queryString, connection)
                connection.Open()
                Dim SQLreader As SqlDataReader = command.ExecuteReader()
                If SQLreader.HasRows Then
                    While (SQLreader.Read())
                        _ID = SQLreader("id")

                        If IsDBNull(SQLreader("read_method")) Then
                            _Method = String.Empty
                        Else
                            _Method = SQLreader("read_method")
                        End If

                        If IsDBNull(SQLreader("mfr_id")) Then
                            _MfrID = String.Empty
                        Else
                            _MfrID = SQLreader("mfr_id").ToString()
                        End If
                    End While
                Else
                    Throw New Exception("Read Method ID was Not found!")
                End If
                
                connection.Close()
                Return "Success"
            End Using
        Catch ex As SqlException
            ErrorHandler.WritelogFile("PMI", "ReadMethod(Class)", "Init()", ex.Message.ToString())
            Return "Error"
        Catch ex As Exception
            ErrorHandler.WritelogFile("PMI", "ReadMethod(Class)", "Init()", ex.Message.ToString())
            Return "Error"
        End Try
    End Function

	
    Public Shared Function ReadMethodExist(Method As String) As String
        Dim Result As String = "no"

        Try
            Dim conString = ConfigurationManager.ConnectionStrings("MSIUBDAPPCodeCon").ConnectionString
            Dim queryString = "Select * FROM parts_read_method WHERE LOWER(read_method) = '" & Method.ToLower() & "'"

            Using connection As New SqlConnection(conString)
                Dim command = New SqlCommand(queryString, connection)
                connection.Open()

                Using SQLreader As SqlDataReader = command.ExecuteReader()
                    If SQLreader.HasRows Then
                        Result = "yes"
                    End If

                    Return Result
                End Using
            End Using
        Catch ex As SqlException
            ErrorHandler.WritelogFile("PMI", "ReadMethod(class)", "ReadMethodExist()", ex.Message.ToString())
            Return "Error"
        Catch ex As Exception
            ErrorHandler.WritelogFile("PMI", "ReadMethod(class)", "ReadMethodExist()", ex.Message.ToString())
            Return "Error"
        End Try
    End Function



    Public Shared Function GetReadMethod(ID As String) As String
        Dim Result As String = String.Empty

        Try
            Dim conString = ConfigurationManager.ConnectionStrings("MSIUBDAPPCodeCon").ConnectionString
            Dim queryString = "Select id, read_method FROM parts_read_method WHERE id = " & ID

            Using connection As New SqlConnection(conString)
                Dim command = New SqlCommand(queryString, connection)
                connection.Open()

                Using SQLreader As SqlDataReader = command.ExecuteReader()
                    If SQLreader.HasRows Then
                        While (SQLreader.Read())
                            Result = SQLreader("read_method")
                        End While
                    End If

                    Return Result
                End Using
            End Using
        Catch ex As SqlException
            ErrorHandler.WritelogFile("PMI", "ReadMethod(class)", "GetReadMethod()", ex.Message.ToString())
            Return "Error"
        Catch ex As Exception
            ErrorHandler.WritelogFile("PMI", "ReadMethod(class)", "GetReadMethod()", ex.Message.ToString())
            Return "Error"
        End Try
    End Function

	
    Public Shared Function AddReadMethod(Method As String, MfrID As String) As String
        Dim query As String = String.Empty
        query &= "INSERT INTO parts_read_method (read_method, mfr_id) VALUES ('" & Method & "', '" & MfrID & "')"

        Dim constr As String = ConfigurationManager.ConnectionStrings("MSIUBDAPPCodeCon").ConnectionString
        Dim strBuf As String = String.Empty

        Try
            Using con As New SqlConnection(constr)
                Using cmd As New SqlCommand()
                    cmd.Connection = con
                    con.Open()

                    cmd.CommandText = query

                    'Insert record
                    cmd.ExecuteNonQuery()

                    Return "Success"
                End Using
            End Using
        Catch ex As SqlException
            ErrorHandler.WritelogFile("PMI", "ReadMethod(class)", "AddReadMethod", ex.Message.ToString())
            Return ex.Message.ToString()
        Catch ex As Exception
            ErrorHandler.WritelogFile("PMI", "ReadMethod(class)", "AddReadMethod", ex.Message.ToString())
            Return ex.Message.ToString()
        End Try
    End Function

	
    Public Shared Function UpdateReadMethod(ReadMethodID As String, Method As String, MfrID As String) As String
        Dim query As String = "UPDATE parts_read_method SET "
        query &= "read_method = @Method, "
        query &= "mfr_id = @MfrID, "
        query &= "WHERE id = " & ReadMethodID


        Dim constr As String = ConfigurationManager.ConnectionStrings("MSIUBDAPPCodeCon").ConnectionString
        Using con As New SqlConnection(constr)
            Using cmd As New SqlCommand()
                Try
                    cmd.CommandText = query
                    cmd.Parameters.Add("@Method", SqlDbType.VarChar).Value = Method
                    cmd.Parameters.Add("@MfrID", SqlDbType.Int).Value = Integer.Parse(MfrID)

                    cmd.Connection = con
                    con.Open()

                    'Update record
                    cmd.ExecuteNonQuery()

                    Return "Success"
                Catch ex As SqlException
                    ErrorHandler.WritelogFile("PMI", "ReadMethod(Class)", "UpdateReadMethod()", ex.Message.ToString())
                    Return ex.Message.ToString()
                Catch ex As Exception
                    ErrorHandler.WritelogFile("PMI", "ReadMethod(Class)", "UpdateReadMethod()", ex.Message.ToString())
                    Return ex.Message.ToString()
                End Try
            End Using
        End Using
    End Function


End Class
