Public Class clsFileListDB
    Public Sub New()
        MyBase.New()
        Dim CurrentAssembly As [Assembly] = Assembly.GetEntryAssembly
        mAppName = CurrentAssembly.GetName.Name.ToString()
        mEventLog = New EventLog("Application") : mEventLog.EnableRaisingEvents = True
        If Not EventLog.SourceExists(mAppName) Then EventLog.CreateEventSource(mAppName, "Application")
        mEventLog.Source = mAppName
    End Sub
    Public Event List(ByVal Message As String)
    Private mEventLog As EventLog
    Private mAppName As String
    Private mCancel As Boolean = False
    Private mCount As Long
    Private mSQLConnect As SqlClient.SqlConnection
    Public Property AppName() As String
        Get
            Return mAppName
        End Get
        Set(ByVal Value As String)
            mAppName = Value
        End Set
    End Property
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
    'Public Property SQLConnect() As SqlClient.SqlConnection
    '    Get
    '        Return mSQLConnect
    '    End Get
    '    Set(ByVal Value As SqlClient.SqlConnection)
    '        mSQLConnect = Value
    '    End Set
    'End Property
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
    Public Function DoFileListDB(ByVal RootDir As String, ByVal DatabaseName As String) As Integer
        Dim ConnectionString As String
        Dim frm As frmProgress
        Dim Message As String
        Dim SQLSource As String
        Dim StartTime As Date

        Try
            Dim Root As DirectoryInfo = New DirectoryInfo(RootDir)
            If Not Root.Exists Then
                MessageBox.Show(String.Format("{0} does not exist!", RootDir), mAppName, MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
                Return 1
            End If

            Message = String.Format("{0} - Listing Files From {1} into {2}.FileList", mAppName, RootDir, DatabaseName)
            StartTime = Now
            mEventLog.WriteEntry(Message)

            frm = New frmProgress(Me)
            With frm
                .Text = Message
                .OKtoClose = False
                .prgProgress.Visible = False
                .Show()

                RaiseEvent List(String.Format("Connecting to {0}", DatabaseName))
                Application.DoEvents() : If mCancel Then Exit Try

                ConnectionString = String.Format("{0}={1};", "Application Name", mAppName)
                ConnectionString &= String.Format("{0}={1};", "Data Source", SystemInformation.ComputerName)
                ConnectionString &= String.Format("{0}={1};", "Initial Catalog ", DatabaseName)
                ConnectionString &= String.Format("{0}={1};", "Integrated Security ", "SSPI")
                ConnectionString &= String.Format("{0}={1};", "Workstation ID", SystemInformation.ComputerName)
                mSQLConnect = New SqlClient.SqlConnection(ConnectionString)
                mSQLConnect.Open()

                RaiseEvent List("Determining File Count...")
                Application.DoEvents() : If mCancel Then Exit Try
                Try
                    Dim SQLCommand As New SqlClient.SqlCommand
                    SQLSource = String.Format("Select Count(*) From FileList Where Path Like '{0}%'", RootDir)
                    With SQLCommand
                        .CommandText = SQLSource
                        .CommandType = CommandType.Text
                        .Connection = mSQLConnect
                        mCount = .ExecuteScalar
                    End With
                Catch ex As Exception
                End Try

                'RaiseEvent List(String.Format("Dropping {0} table...", "FileList"))
                'Application.DoEvents() : If mCancel Then Exit Try
                'Try
                '    SQLSource = "DROP TABLE FileList"
                '    DoCommand(SQLSource)
                'Catch ex As Exception
                'End Try

                'RaiseEvent List(String.Format("Recreating {0} table...", "FileList"))
                'Application.DoEvents() : If mCancel Then Exit Try
                'SQLSource = "CREATE TABLE FileList ("
                'SQLSource &= "[ID] int NOT NULL IDENTITY (1, 1),"
                'SQLSource &= "[Path] [nvarchar] (1024) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,"
                'SQLSource &= "[Size] [bigint] NULL,"
                'SQLSource &= "[Attributes] [nvarchar] (256) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,"
                'SQLSource &= "[CreationTime] [datetime] NULL,"
                'SQLSource &= "[LastAccessTime] [datetime] NULL,"
                'SQLSource &= "[LastWriteTime] [datetime] NULL"
                'SQLSource &= ") ON [PRIMARY]"
                'DoCommand(SQLSource)

                RaiseEvent List(String.Format("Deleting {0}% records...", RootDir))
                Application.DoEvents() : If mCancel Then Exit Try
                Try
                    SQLSource = String.Format("Delete From FileList Where Path Like '{0}%'", RootDir)
                    DoCommand(SQLSource)
                Catch ex As Exception
                End Try

                .prgProgress.Visible = True
                mCount = 0
                ListFiles(Root)
                .OKtoClose = True
                .Close()
            End With

            Message = String.Format("{0} Complete - {1:#,##0} entries written to {2}.FileList", mAppName, mCount, DatabaseName) & vbCrLf
            Message &= vbCrLf
            Message &= String.Format("Elapsed Time: {1}", Message, ElapsedTime(StartTime, Now))
            mEventLog.WriteEntry(Message)
            Return 0
        Catch ex As Exception
            MessageBox.Show(ex.ToString, mAppName, MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
            Return 1
        Finally
        End Try
    End Function
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
    Private Function ElapsedTime(ByRef StartTime As Date, ByRef FinishTime As Date, Optional ByRef tFormat As Short = 0) As String
        Dim MM, HH, SS As Integer
        Dim strTime As String

        strTime = vbNullString
        SS = DateDiff(Microsoft.VisualBasic.DateInterval.Second, StartTime, FinishTime)
        HH = SS \ 3600
        SS = SS - (HH * 3600)
        MM = SS \ 60
        SS = SS - (MM * 60)
        If HH > 0 Then strTime = HH & " Hours, "
        If MM > 0 Then strTime = strTime & MM & " Minutes, "
        strTime = strTime & SS & " Seconds"

        Select Case tFormat
            Case 0
                ElapsedTime = VB6.Format(HH, "00") & ":" & VB6.Format(MM, "00") & ":" & VB6.Format(SS, "00")
            Case Else
                ElapsedTime = strTime
        End Select
    End Function
    Public Sub ListFiles(ByVal BaseDir As DirectoryInfo)
        Dim SQLSource As String = vbNullString
        Dim ColumnList As String = "[Path],[Size],[Attributes],[CreationTime],[LastAccessTime],[LastWriteTime]"
        Try
            Dim diList As DirectoryInfo() = BaseDir.GetDirectories()
            For Each di As DirectoryInfo In diList
                Select Case di.Name
                    Case "System Volume Information"
                    Case "Temporary Internet Files"
                    Case Else
                        Try
                            mCount += 1
                            RaiseEvent List(di.FullName)
                            Application.DoEvents() : If mCancel Then Exit Try

                            SQLSource = vbNullString
                            Dim ValueList() As String = { _
                                di.FullName.Replace("'", "''"), _
                                di.Attributes.ToString, _
                                IIf(di.CreationTime >= CDate("01/01/1753"), String.Format("'{0}'", di.CreationTime.ToString), "NULL"), _
                                IIf(di.LastAccessTime >= CDate("01/01/1753"), String.Format("'{0}'", di.LastAccessTime.ToString), "NULL"), _
                                IIf(di.LastWriteTime >= CDate("01/01/1753"), String.Format("'{0}'", di.LastWriteTime.ToString), "NULL")}
                            SQLSource = String.Format("'{0}',0,'{1}',{2},{3},{4}", ValueList)
                            SQLSource = String.Format("INSERT INTO FileList ({0}) VALUES ({1})", ColumnList, SQLSource)
                            Me.DoCommand(SQLSource)
                        Catch ex As Exception
                            Dim Message As String = String.Format("Error processing {0}; ", di.Name) & vbCrLf
                            Message &= vbCrLf
                            If SQLSource <> vbNullString Then
                                Message &= String.Format("SQL: {0}", SQLSource) & vbCrLf
                                Message &= vbCrLf
                            End If
                            Message &= String.Format("Exception: {0}", ex.ToString)
                            mEventLog.WriteEntry(Message)
                        End Try
                        Application.DoEvents() : If mCancel Then Exit Try
                        ListFiles(di)
                End Select
            Next
            Dim fiList As FileInfo() = BaseDir.GetFiles()
            For Each fi As FileInfo In fiList
                Try
                    mCount += 1
                    RaiseEvent List(fi.FullName)
                    Application.DoEvents() : If mCancel Then Exit Try

                    SQLSource = vbNullString
                    Dim ValueList() As String = { _
                        fi.FullName.Replace("'", "''"), _
                        fi.Length.ToString, _
                        fi.Attributes.ToString, _
                        IIf(fi.CreationTime >= CDate("01/01/1753"), String.Format("'{0}'", fi.CreationTime.ToString), "NULL"), _
                        IIf(fi.LastAccessTime >= CDate("01/01/1753"), String.Format("'{0}'", fi.LastAccessTime.ToString), "NULL"), _
                        IIf(fi.LastWriteTime >= CDate("01/01/1753"), String.Format("'{0}'", fi.LastWriteTime.ToString), "NULL")}
                    SQLSource = String.Format("'{0}',{1},'{2}',{3},{4},{5}", ValueList)
                    SQLSource = String.Format("INSERT INTO FileList ({0}) VALUES ({1})", ColumnList, SQLSource)
                    Me.DoCommand(SQLSource)
                Catch ex As Exception
                    Dim Message As String = String.Format("Error processing {0}; ", fi.Name) & vbCrLf
                    Message &= vbCrLf
                    If SQLSource <> vbNullString Then
                        Message &= String.Format("SQL: {0}", SQLSource) & vbCrLf
                        Message &= vbCrLf
                    End If
                    Message &= String.Format("Exception: {0}", ex.ToString)
                    mEventLog.WriteEntry(Message)
                End Try
                Application.DoEvents() : If mCancel Then Exit Try
            Next
        Catch ex As Exception
            MessageBox.Show(ex.ToString, mAppName & ".ListFiles", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
        End Try
    End Sub
    'Entry point which delegates to C-style main Private Function
    Public Overloads Shared Sub Main()
        System.Environment.ExitCode = Main(System.Environment.GetCommandLineArgs())
    End Sub
    Private Overloads Shared Function Main(ByVal args() As String) As Integer
        Dim fl As New clsFileListDB
        Try
            Application.EnableVisualStyles()
            Application.DoEvents()

            If args.Length < 3 Then
                Dim Message As String = _
                    "FileList Options:" & vbCrLf & _
                    vbTab & "args(0) = Full path name of the executable" & vbCrLf & _
                    vbTab & "args(1) = Root directory to list files" & vbCrLf & _
                    vbTab & "args(2) = Database name"
                MessageBox.Show(Message, fl.AppName, MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
                Return 1
            End If
            Dim RootDir As String = args(1)
            Dim DatabaseName As String = args(2)
            Return fl.DoFileListDB(RootDir, DatabaseName)
        Catch ex As Exception
            MessageBox.Show(ex.ToString, fl.AppName, MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
            Return 1
        Finally
        End Try
    End Function
End Class
