<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As ComponentModel.ComponentResourceManager = New ComponentModel.ComponentResourceManager(GetType(Form1))
        BtnOpenFolder = New Button()
        BtnMakeDB = New Button()
        TxtYtKey = New TextBox()
        Label1 = New Label()
        BtnLoadTXT = New Button()
        ListBox1 = New ListBox()
        BtnLoadKey = New Button()
        Label2 = New Label()
        TxtKeyPath = New TextBox()
        GroupBox1 = New GroupBox()
        Label5 = New Label()
        TxtUsername = New TextBox()
        Label6 = New Label()
        TxtPort = New TextBox()
        TxtPassphrase = New TextBox()
        Label4 = New Label()
        TxtHost = New TextBox()
        Label3 = New Label()
        BtnUploadDB = New Button()
        BtnTestSSH = New Button()
        BtnSaveConfig = New Button()
        GroupBox1.SuspendLayout()
        SuspendLayout()
        ' 
        ' BtnOpenFolder
        ' 
        BtnOpenFolder.Location = New Point(246, 99)
        BtnOpenFolder.Name = "BtnOpenFolder"
        BtnOpenFolder.Size = New Size(111, 37)
        BtnOpenFolder.TabIndex = 13
        BtnOpenFolder.Text = "DB 폴더 열기"
        BtnOpenFolder.UseVisualStyleBackColor = True
        ' 
        ' BtnMakeDB
        ' 
        BtnMakeDB.Location = New Point(246, 56)
        BtnMakeDB.Name = "BtnMakeDB"
        BtnMakeDB.Size = New Size(111, 37)
        BtnMakeDB.TabIndex = 12
        BtnMakeDB.Text = "DB 생성하기"
        BtnMakeDB.UseVisualStyleBackColor = True
        ' 
        ' TxtYtKey
        ' 
        TxtYtKey.Location = New Point(6, 81)
        TxtYtKey.Name = "TxtYtKey"
        TxtYtKey.Size = New Size(222, 23)
        TxtYtKey.TabIndex = 11
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(6, 63)
        Label1.Name = "Label1"
        Label1.Size = New Size(81, 15)
        Label1.TabIndex = 10
        Label1.Text = "유튜브 API 키"
        ' 
        ' BtnLoadTXT
        ' 
        BtnLoadTXT.Location = New Point(246, 12)
        BtnLoadTXT.Name = "BtnLoadTXT"
        BtnLoadTXT.Size = New Size(111, 37)
        BtnLoadTXT.TabIndex = 9
        BtnLoadTXT.Text = "대화파일 로드"
        BtnLoadTXT.UseVisualStyleBackColor = True
        ' 
        ' ListBox1
        ' 
        ListBox1.FormattingEnabled = True
        ListBox1.ItemHeight = 15
        ListBox1.Items.AddRange(New Object() {"HonDaeChun WEB 관리자 ", "버전 0.1 / 2023-07-16"})
        ListBox1.Location = New Point(12, 12)
        ListBox1.Name = "ListBox1"
        ListBox1.Size = New Size(228, 124)
        ListBox1.TabIndex = 8
        ' 
        ' BtnLoadKey
        ' 
        BtnLoadKey.Font = New Font("맑은 고딕", 6.75F, FontStyle.Regular, GraphicsUnit.Point)
        BtnLoadKey.ImageAlign = ContentAlignment.TopCenter
        BtnLoadKey.Location = New Point(434, 37)
        BtnLoadKey.Name = "BtnLoadKey"
        BtnLoadKey.Size = New Size(22, 23)
        BtnLoadKey.TabIndex = 16
        BtnLoadKey.Text = "..."
        BtnLoadKey.UseVisualStyleBackColor = True
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(234, 19)
        Label2.Name = "Label2"
        Label2.Size = New Size(74, 15)
        Label2.TabIndex = 17
        Label2.Text = "SSH 키 파일"
        ' 
        ' TxtKeyPath
        ' 
        TxtKeyPath.Location = New Point(234, 37)
        TxtKeyPath.Name = "TxtKeyPath"
        TxtKeyPath.Size = New Size(194, 23)
        TxtKeyPath.TabIndex = 18
        ' 
        ' GroupBox1
        ' 
        GroupBox1.Controls.Add(Label5)
        GroupBox1.Controls.Add(TxtUsername)
        GroupBox1.Controls.Add(Label6)
        GroupBox1.Controls.Add(TxtPort)
        GroupBox1.Controls.Add(TxtPassphrase)
        GroupBox1.Controls.Add(Label1)
        GroupBox1.Controls.Add(TxtYtKey)
        GroupBox1.Controls.Add(Label4)
        GroupBox1.Controls.Add(Label2)
        GroupBox1.Controls.Add(TxtKeyPath)
        GroupBox1.Controls.Add(TxtHost)
        GroupBox1.Controls.Add(BtnLoadKey)
        GroupBox1.Controls.Add(Label3)
        GroupBox1.Location = New Point(12, 142)
        GroupBox1.Name = "GroupBox1"
        GroupBox1.Size = New Size(462, 113)
        GroupBox1.TabIndex = 19
        GroupBox1.TabStop = False
        GroupBox1.Text = "설정"
        ' 
        ' Label5
        ' 
        Label5.AutoSize = True
        Label5.Location = New Point(319, 63)
        Label5.Name = "Label5"
        Label5.Size = New Size(108, 15)
        Label5.TabIndex = 28
        Label5.Text = "SSH 키 passphrase"
        ' 
        ' TxtUsername
        ' 
        TxtUsername.Location = New Point(234, 81)
        TxtUsername.Name = "TxtUsername"
        TxtUsername.Size = New Size(79, 23)
        TxtUsername.TabIndex = 27
        ' 
        ' Label6
        ' 
        Label6.AutoSize = True
        Label6.Location = New Point(234, 63)
        Label6.Name = "Label6"
        Label6.Size = New Size(71, 15)
        Label6.TabIndex = 25
        Label6.Text = "로그인 계정"
        ' 
        ' TxtPort
        ' 
        TxtPort.Location = New Point(163, 37)
        TxtPort.Name = "TxtPort"
        TxtPort.Size = New Size(59, 23)
        TxtPort.TabIndex = 22
        ' 
        ' TxtPassphrase
        ' 
        TxtPassphrase.Location = New Point(319, 81)
        TxtPassphrase.Name = "TxtPassphrase"
        TxtPassphrase.Size = New Size(137, 23)
        TxtPassphrase.TabIndex = 24
        TxtPassphrase.UseSystemPasswordChar = True
        ' 
        ' Label4
        ' 
        Label4.AutoSize = True
        Label4.Location = New Point(161, 19)
        Label4.Name = "Label4"
        Label4.Size = New Size(58, 15)
        Label4.TabIndex = 21
        Label4.Text = "SSH 포트"
        ' 
        ' TxtHost
        ' 
        TxtHost.Location = New Point(6, 37)
        TxtHost.Name = "TxtHost"
        TxtHost.Size = New Size(151, 23)
        TxtHost.TabIndex = 20
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(6, 19)
        Label3.Name = "Label3"
        Label3.Size = New Size(43, 15)
        Label3.TabIndex = 19
        Label3.Text = "호스트"
        ' 
        ' BtnUploadDB
        ' 
        BtnUploadDB.Location = New Point(363, 56)
        BtnUploadDB.Name = "BtnUploadDB"
        BtnUploadDB.Size = New Size(111, 37)
        BtnUploadDB.TabIndex = 20
        BtnUploadDB.Text = "DB 업로드"
        BtnUploadDB.UseVisualStyleBackColor = True
        ' 
        ' BtnTestSSH
        ' 
        BtnTestSSH.Location = New Point(363, 12)
        BtnTestSSH.Name = "BtnTestSSH"
        BtnTestSSH.Size = New Size(111, 37)
        BtnTestSSH.TabIndex = 22
        BtnTestSSH.Text = "SSH 접속 테스트"
        BtnTestSSH.UseVisualStyleBackColor = True
        ' 
        ' BtnSaveConfig
        ' 
        BtnSaveConfig.Location = New Point(363, 99)
        BtnSaveConfig.Name = "BtnSaveConfig"
        BtnSaveConfig.Size = New Size(111, 37)
        BtnSaveConfig.TabIndex = 29
        BtnSaveConfig.Text = "설정 저장"
        BtnSaveConfig.UseVisualStyleBackColor = True
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7.0F, 15.0F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(485, 268)
        Controls.Add(BtnSaveConfig)
        Controls.Add(BtnTestSSH)
        Controls.Add(BtnUploadDB)
        Controls.Add(GroupBox1)
        Controls.Add(BtnOpenFolder)
        Controls.Add(BtnMakeDB)
        Controls.Add(BtnLoadTXT)
        Controls.Add(ListBox1)
        FormBorderStyle = FormBorderStyle.FixedSingle
        Icon = CType(resources.GetObject("$this.Icon"), Icon)
        MaximizeBox = False
        Name = "Form1"
        Text = "HonDaeChun WEB 관리자"
        GroupBox1.ResumeLayout(False)
        GroupBox1.PerformLayout()
        ResumeLayout(False)
    End Sub
    Friend WithEvents BtnOpenFolder As Button
    Friend WithEvents BtnMakeDB As Button
    Friend WithEvents TxtYtKey As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents BtnLoadTXT As Button
    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents BtnLoadKey As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents TxtKeyPath As TextBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents TxtPort As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents TxtHost As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents BtnUploadDB As Button
    Friend WithEvents Label6 As Label
    Friend WithEvents TxtPassphrase As TextBox
    Friend WithEvents TxtUsername As TextBox
    Friend WithEvents BtnTestSSH As Button
    Friend WithEvents Label5 As Label
    Friend WithEvents BtnSaveConfig As Button
End Class
