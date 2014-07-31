Public Class SQL_Executor

    Inherits System.ComponentModel.Component

    ' Declares the variables you will use to hold your thread objects.
    
    Public SQLThread As System.Threading.Thread

    Public SQL As String = ""
    Public Result As String = ""
    Public Database As String = ""

    Public Event SQLComplete(ByVal Result as string)
    

#Region " Component Designer generated code "

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        Container.Add(Me)
    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub error_handler(ByVal message As String)
        Try
            MsgBox("Access SQL Executor 1.0 has trapped the following error: " & vbCrLf & message, MsgBoxStyle.Exclamation, "Access SQL Executor 1.0")
        Catch ex As Exception
            MsgBox("Access SQL Executor 1.0 has trapped the following error: " & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Access SQL Executor 1.0")
        End Try
    End Sub

    Public Sub ChooseThreads(ByVal threadNumber As Integer)
        Try
            ' Determines which thread to start based on the value it receives.
            Select Case threadNumber
                Case 1
                    ' Sets the thread using the AddressOf the subroutine where
                    ' the thread will start.
                    SQLThread = New System.Threading.Thread(AddressOf SQLExecute)
                    ' Starts the thread.
                    SQLThread.Start()
            End Select
        Catch ex As Exception
            error_handler(ex.Message)
        End Try
    End Sub

    Public Sub SQLExecute()
        Dim Conn As Data.OleDb.OleDbConnection
        Try
            Conn = New Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Database & ";")
            Conn.Open()
            Dim recset As Data.OleDb.OleDbCommand = New Data.OleDb.OleDbCommand(SQL, Conn)
            If SQL.Trim().ToLower.StartsWith("select") Then
                Result = (recset.ExecuteNonQuery()).ToString
                Result = "Your query has been executed, though this program is not designed to deal with 'Select from' SQL queries. Its primary purpose is to handle 'Update' and 'Insert into' SQL queries."
            Else
                Result = "Your SQL command has been successfully executed. The number of affected rows is: " & (recset.ExecuteNonQuery()).ToString
            End If
            recset.Dispose()
            Conn.Close()
            Conn.Dispose()
            RaiseEvent SQLComplete(Result)
        Catch dberror As OleDb.OleDbException
            Result = dberror.Message
            Conn.Close()
            RaiseEvent SQLComplete(Result)
        Catch othererror As Exception
            Result = othererror.Message
            RaiseEvent SQLComplete(Result)
        End Try

    End Sub





End Class
