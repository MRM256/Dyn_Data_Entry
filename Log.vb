Imports System.IO
Imports System.Text
Imports System.Windows.Forms

Public Class Log
#Region "ErrorLogger Class"
    Public Class ErrorLogger
        Public Sub New()
            'default constructor
        End Sub

        '-------------------------------------------------------------
        'Name:          WriteToErrorLog
        'Purpose:       Open or create an error log and submit error message
        'Parameters:    msg - message to be written to error file
        '               stkTrace - stack trace from error message
        '               title - title of the error file entry
        'Returns:       Nothing
        '-------------------------------------------------------------
        Public Sub WriteToErrorLog(ByVal msg As String, _
                                   ByVal stkTrace As String, _
                                   ByVal title As String)

            'check and make the directory if necessary; this is set to look in the application
            'folder, you may wish to place the error log in another location depending upon the
            'the user's role and write access to different areas of the file system
            If Not System.IO.Directory.Exists(Application.StartupPath & "\Errors\") Then
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Errors\")
            End If

            'check the file
            Dim fs As FileStream = New FileStream(Application.StartupPath & _
                                                  "\Errors\errlog.txt", _
                                                  FileMode.OpenOrCreate, _
                                                  FileAccess.ReadWrite)
            Dim s As StreamWriter = New StreamWriter(fs)
            s.Close()
            'fs.Close()

            'log it
            Dim fs1 As FileStream = New FileStream(Application.StartupPath & _
                                                   "\Errors\errlog.txt", _
                                                   FileMode.Append, _
                                                   FileAccess.Write)
            Dim s1 As StreamWriter = New StreamWriter(fs1)
            s1.Write("Title: " & title & vbCrLf)
            s1.Write("Message: " & msg & vbCrLf)
            s1.Write("StackTrace: " & stkTrace & vbCrLf)
            s1.Write("Date/Time: " & DateTime.Now.ToString() & vbCrLf)
            s1.Write("==============================================" & _
                     "=============================================" & vbCrLf)
            s1.Close()
            'fs1.Close()
        End Sub

        Public Sub Log_DataTable(ByVal dt As DataTable)
            'Purpose:       Writes the entire DataTable into
            '               a text file
            'Parameters:    dt As DataTable
            'Returns:       Nothing - Creates a text file

            Dim currentRows = dt.[Select](Nothing, Nothing, _
                                          DataViewRowState.CurrentRows)
            If Not System.IO.Directory.Exists(Application.StartupPath & "\DataSet\") Then
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\DataSet\")
            End If

            'check the file
            Dim fs As FileStream = New FileStream(Application.StartupPath & _
                                                  "\DataSet\DataTableLog.txt", _
                                                  FileMode.OpenOrCreate, _
                                                  FileAccess.ReadWrite)
            Dim s As StreamWriter = New StreamWriter(fs)
            s.Close()

            'log it
            Dim fs1 As FileStream = New FileStream(Application.StartupPath & _
                                                   "\DataSet\DataTableLog.txt", _
                                                   FileMode.Append, _
                                                   FileAccess.Write)
            Dim s1 As StreamWriter = New StreamWriter(fs1)
            Dim strTbl As String = ""
            Dim format As String = "{0,10}"

            For Each column As DataColumn In dt.Columns
                format = "{0, " & column.ColumnName.Length & "}"
                s1.Write(format, column.ColumnName & vbTab)
            Next

            s1.Write(vbCrLf)

            For Each row As DataRow In dt.Rows
                For Each column As DataColumn In dt.Columns
                    format = "{0, " & column.ColumnName.Length & "}"
                    s1.Write(format, row(column).ToString & vbTab)
                Next
                s1.Write(vbCrLf)
            Next
            s1.Close()
        End Sub
    End Class
#End Region
    End Class
