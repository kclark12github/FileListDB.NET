Imports System.IO
Imports System.IO.Directory
Imports System.IO.File
Public Class clsFileListDB
    Public Event List(ByVal Message As String)
    Private mCancel As Boolean = False
    Private mCount As Long
    Private mSQLConnect As SqlClient.SqlConnection
    Public Property Cancel() As Boolean
        Get
            Return mCancel
        End Get
        Set(ByVal Value As Boolean)
            mCancel = Value
        End Set
    End Property
    Public Property Count() As Long
        Get
            Return mCount
        End Get
        Set(ByVal Value As Long)
            mCount = Value
        End Set
    End Property
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
                Select Case di.Name
                    Case "System Volume Information"
                    Case "Temporary Internet Files"
                    Case Else
                        Dim ValueList() As String = { _
                            Replace(di.FullName, "'", "''"), _
                            di.Attributes.ToString, _
                            IIf(di.CreationTime >= CDate("01/01/1753"), di.CreationTime.ToString, "NULL"), _
                            IIf(di.LastAccessTime >= CDate("01/01/1753"), di.LastAccessTime.ToString, "NULL"), _
                            IIf(di.LastWriteTime >= CDate("01/01/1753"), di.LastWriteTime.ToString, "NULL")}
                        SQLSource = String.Format("'{0}',0,'{1}','{2}','{3}','{4}'", ValueList)
                        SQLSource = String.Format("INSERT INTO FileList ({0}) VALUES ({1})", ColumnList, SQLSource)
                        Me.DoCommand(SQLSource)
                        RaiseEvent List(di.FullName)
                        If mCancel Then Exit Try
                        ListFiles(di)
                End Select
            Next
            Dim fiList As FileInfo() = BaseDir.GetFiles()
            For Each fi As FileInfo In fiList
                Dim ValueList() As String = { _
                    Replace(fi.FullName, "'", "''"), _
                    fi.Length.ToString, _
                    fi.Attributes.ToString, _
                    IIf(fi.CreationTime >= CDate("01/01/1753"), fi.CreationTime.ToString, "NULL"), _
                    IIf(fi.LastAccessTime >= CDate("01/01/1753"), fi.LastAccessTime.ToString, "NULL"), _
                    IIf(fi.LastWriteTime >= CDate("01/01/1753"), fi.LastWriteTime.ToString, "NULL")}
                SQLSource = String.Format("'{0}',{1},'{2}','{3}','{4}','{5}'", ValueList)
                SQLSource = String.Format("INSERT INTO FileList ({0}) VALUES ({1})", ColumnList, SQLSource)
                Me.DoCommand(SQLSource)
                RaiseEvent List(fi.FullName)
                If mCancel Then Exit Try
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
        Dim iCount As Long
        Dim SQLSource As String
        Try
            Application.EnableVisualStyles()
            Application.DoEvents()

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
                Dim SQLCommand As New SqlClient.SqlCommand
                SQLSource = "Select Count(*) From FileList"
                With SQLCommand
                    .CommandText = SQLSource
                    .CommandType = CommandType.Text
                    .Connection = fl.SQLConnect
                    fl.Count = .ExecuteScalar
                End With
            Catch ex As Exception
            End Try

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
            SQLSource &= "[CreationTime] [datetime] NULL,"
            SQLSource &= "[LastAccessTime] [datetime] NULL,"
            SQLSource &= "[LastWriteTime] [datetime] NULL"
            SQLSource &= ") ON [PRIMARY]"
            fl.DoCommand(SQLSource)

            'Use these stats to display a window with a progress bar...
            Dim frm As New frmProgress(fl)
            With frm
                .Text = String.Format("FileListDB - Listing Files From {0} into {1}", args(1), args(2))
                .Show()
                fl.ListFiles(Root)
                .Close()
            End With
            Return 0
        Catch ex As Exception
            MessageBox.Show(ex.ToString, "FileListDB", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
            Return 1
        Finally
        End Try
    End Function
End Class
