Imports System.IO
Imports System.IO.Directory
Imports System.IO.File
Public Class clsFileListDB
    Private mSQLConnect As SqlClient.SqlConnection
    Public Property SQLConnect() As SqlClient.SqlConnection
        Get
            Return mSQLConnect
        End Get
        Set(ByVal Value As SqlClient.SqlConnection)
            mSQLConnect = Value
        End Set
    End Property
    Public Sub DoCommand(ByVal SQLSource As String)
        Dim SQLCommand As New SqlClient.SqlCommand
        Try
            With SQLCommand
                .CommandText = SQLSource
                .CommandType = CommandType.Text
                .Connection = mSQLConnect
                .ExecuteNonQuery()
            End With
        Catch ex As Exception
            Throw ex
        Finally
            SQLCommand = Nothing
        End Try
    End Sub
    Public Sub CountFiles(ByVal BaseDir As DirectoryInfo, ByRef cntFolders As Long, ByRef cntFiles As Long)
        Try
            Dim diList As DirectoryInfo() = BaseDir.GetDirectories()
            For Each di As DirectoryInfo In diList
                cntFolders += 1
                CountFiles(di, cntFiles, cntFolders)
            Next
            Dim fiList As FileInfo() = BaseDir.GetFiles()
            cntFiles += fiList.Length
        Catch ex As Exception
            MessageBox.Show(ex.ToString, "FileList", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
        End Try
    End Sub
    Public Sub ListFiles(ByVal BaseDir As DirectoryInfo)
        Dim SQLSource As String
        Dim ColumnList As String = "[Path],[Size],[Attributes],[CreationTime],[LastAccessTime],[LastWriteTime]"
        Try
            Dim diList As DirectoryInfo() = BaseDir.GetDirectories()
            For Each di As DirectoryInfo In diList
                Dim ValueList() As String = { _
                    Replace(di.FullName, "'", "''"), _
                    di.Attributes.ToString, _
                    di.CreationTime.ToString, _
                    di.LastAccessTime.ToString, _
                    di.LastWriteTime.ToString}
                SQLSource = String.Format("'{0}',0,'{1}','{2}','{3}','{4}'", ValueList)
                SQLSource = String.Format("INSERT INTO FileList ({0}) VALUES ({1})", ColumnList, SQLSource)
                Me.DoCommand(SQLSource)
                Select Case di.Name
                    Case "Temporary Internet Files"
                    Case Else
                        ListFiles(di)
                End Select
            Next
            Dim fiList As FileInfo() = BaseDir.GetFiles()
            For Each fi As FileInfo In fiList
                Dim ValueList() As String = { _
                    Replace(fi.FullName, "'", "''"), _
                    fi.Length.ToString, _
                    fi.Attributes.ToString, _
                    fi.CreationTime.ToString, _
                    fi.LastAccessTime.ToString, _
                    fi.LastWriteTime.ToString}
                SQLSource = String.Format("'{0}',{1},'{2}','{3}','{4}','{5}'", ValueList)
                SQLSource = String.Format("INSERT INTO FileList ({0}) VALUES ({1})", ColumnList, SQLSource)
                Me.DoCommand(SQLSource)
            Next
        Catch ex As Exception
            MessageBox.Show(ex.ToString, "FileListDB.ListFiles", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
        End Try
    End Sub
    'Entry point which delegates to C-style main Private Function
    Public Overloads Shared Sub Main()
        System.Environment.ExitCode = Main(System.Environment.GetCommandLineArgs())
    End Sub
    Private Overloads Shared Function Main(ByVal args() As String) As Integer
        Dim fl As New clsFileListDB
        Dim SQLSource As String
        Try
            If args.Length < 3 Then
                Dim Message As String = _
                    "FileList Options:" & vbCrLf & _
                    vbTab & "args(0) = Full path name of the executable" & vbCrLf & _
                    vbTab & "args(1) = Root directory to list files" & vbCrLf & _
                    vbTab & "args(2) = Database name"
                MessageBox.Show(Message, "FileList", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
                Return 1
            End If

            Dim Root As DirectoryInfo = New DirectoryInfo(args(1))
            If Not Root.Exists Then
                MessageBox.Show(String.Format("{0} does not exist!", args(1)), "FileList", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
                Return 1
            End If

            Dim ConnectionString As String
            ConnectionString = String.Format("{0}={1};", "Application Name", "FileListDB")
            ConnectionString &= String.Format("{0}={1};", "Data Source", SystemInformation.ComputerName)
            ConnectionString &= String.Format("{0}={1};", "Initial Catalog ", args(2))
            ConnectionString &= String.Format("{0}={1};", "Integrated Security ", "SSPI")
            ConnectionString &= String.Format("{0}={1};", "Workstation ID", SystemInformation.ComputerName)
            fl.SQLConnect = New SqlClient.SqlConnection(ConnectionString)
            fl.SQLConnect.Open()

            Try
                SQLSource = "DROP TABLE FileList"
                fl.DoCommand(SQLSource)
            Catch ex As Exception
            End Try

            SQLSource = "CREATE TABLE FileList ("
            SQLSource &= "[ID] int NOT NULL IDENTITY (1, 1),"
            SQLSource &= "[Path] [nvarchar] (1024) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,"
            SQLSource &= "[Size] [bigint] NULL,"
            SQLSource &= "[Attributes] [nvarchar] (256) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,"
            SQLSource &= "[CreationTime] [datetime] NOT NULL,"
            SQLSource &= "[LastAccessTime] [datetime] NOT NULL,"
            SQLSource &= "[LastWriteTime] [datetime] NOT NULL"
            SQLSource &= ") ON [PRIMARY]"
            fl.DoCommand(SQLSource)

            'Dim cntFolders As Long = 0
            'Dim cntFiles As Long = 0
            'fl.CountFiles(Root, cntFolders, cntFiles)
            ''Use these stats to display a window with a progress bar...
            'Dim frm As New frmStats
            'frm.ShowDialog()
            fl.ListFiles(Root)
            Return 0
        Catch ex As Exception
                MessageBox.Show(ex.ToString, "FileListDB", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
            Return 1
        Finally
        End Try
    End Function
End Class
