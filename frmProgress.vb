Public Class frmProgress
    Inherits System.Windows.Forms.Form
    Public Sub New(ByVal objBase As clsFileListDB)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        mBaseClass = objBase
        mOffset = Me.lblProgress.Top - Me.prgProgress.Top
        prgProgress.Visible = CBool(mBaseClass.Count = 0)
    End Sub

#Region " Windows Form Designer generated code "


    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents prgProgress As System.Windows.Forms.ProgressBar
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.prgProgress = New System.Windows.Forms.ProgressBar
        Me.lblProgress = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'prgProgress
        '
        Me.prgProgress.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prgProgress.Location = New System.Drawing.Point(24, 20)
        Me.prgProgress.Name = "prgProgress"
        Me.prgProgress.Size = New System.Drawing.Size(624, 24)
        Me.prgProgress.Step = 1
        Me.prgProgress.TabIndex = 0
        '
        'lblProgress
        '
        Me.lblProgress.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblProgress.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.lblProgress.Location = New System.Drawing.Point(28, 64)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(620, 64)
        Me.lblProgress.TabIndex = 1
        '
        'frmProgress
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(664, 136)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.prgProgress)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(1024, 176)
        Me.MinimumSize = New System.Drawing.Size(672, 176)
        Me.Name = "frmProgress"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmProgress"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private WithEvents mBaseClass As clsFileListDB
    Private mOffset As Single
    Private mOKtoClose As Boolean
    Public Property BaseClass() As clsFileListDB
        Get
            Return mBaseClass
        End Get
        Set(ByVal Value As clsFileListDB)
            mBaseClass = Value
        End Set
    End Property
    Public Property OKtoClose() As Boolean
        Get
            Return mOKtoClose
        End Get
        Set(ByVal Value As Boolean)
            mOKtoClose = Value
        End Set
    End Property
    Private Sub mBaseClass_List(ByVal Message As String) Handles mBaseClass.List
        lblProgress.Text = Message
        With prgProgress
            If (.Value + 1) <= .Maximum Then .Value += 1
        End With
        Application.DoEvents()
    End Sub
    Private Sub frmProgress_Closing(ByVal EventSender As Object, ByVal EventArgs As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Not mOKtoClose Then
            If MessageBox.Show("OK to Stop FileListDB?", "Stop FileListDB", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = DialogResult.No Then
                EventArgs.Cancel = True
            Else
                mBaseClass.Cancel = True
            End If
        End If
    End Sub
    Private Sub prgProgress_VisibleChanged(ByVal EventSender As Object, ByVal EventArgs As System.EventArgs) Handles prgProgress.VisibleChanged
        If IsNothing(mBaseClass) Then Exit Sub
        With Me
            .prgProgress.Minimum = 0
            .prgProgress.Value = 0
            .prgProgress.Maximum = mBaseClass.Count
            If Not .prgProgress.Visible Then
                .lblProgress.Top = .prgProgress.Top
                .Size = New System.Drawing.Size(.Width, .Height - mOffset)
            Else
                .lblProgress.Top = .prgProgress.Top + mOffset
                .Size = New System.Drawing.Size(.Width, .Height + mOffset)
            End If
        End With
    End Sub
End Class
