Imports System.Data.SQLite
Imports System.Globalization
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Text.RegularExpressions
Imports System.Threading
Imports Google.Apis.Services
Imports Google.Apis.YouTube.v3
Imports HtmlAgilityPack
Imports IniParser
Imports IniParser.Model
Imports Renci.SshNet

Public Class Form1

    Private cts As CancellationTokenSource
    Private ytList, ytUpdateList, appleList, appleUpdateList, fileList, fileUpdateList As New List(Of List(Of String))
    Private ConfigLoadFinished As Boolean = False
    Private HasConfigChanged As Boolean = False

    Private Sub ClearLists(Optional mode As String = "")

        ytUpdateList.Clear()
        appleUpdateList.Clear()
        fileUpdateList.Clear()
        If mode = "ALL" Then
            ytList.Clear()
            appleList.Clear()
            fileList.Clear()
        End If

    End Sub

    Private Sub LoadConfig()

        If Not File.Exists("config.ini") Then
            MessageBox.Show("config.ini 파일이 없습니다." & vbCrLf & "파일을 확인한 뒤 다시 실행해주세요.", "경고", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Environment.Exit(0)
        End If

        Dim parser As New FileIniDataParser()
        Dim data As IniData = parser.ReadFile("config.ini")

        TxtKeyPath.Text = data("Settings").GetKeyData("keyfilePath").Value
        TxtHost.Text = data("Settings").GetKeyData("host").Value
        TxtPort.Text = data("Settings").GetKeyData("port").Value
        TxtYtKey.Text = data("Settings").GetKeyData("ytapikey").Value
        TxtUsername.Text = data("Settings").GetKeyData("username").Value
        TxtPassphrase.Text = data("Settings").GetKeyData("keyPassphrase").Value
        ChkBaroUpload.Checked = (data("Settings").GetKeyData("BaroUpload").Value.ToLower() = "true")
        ChkShutdown.Checked = (data("Settings").GetKeyData("ShutdownAfterUpload").Value.ToLower() = "true")

    End Sub

    Private Sub SaveConfig()

        Dim parser As New FileIniDataParser()
        Dim data As IniData = parser.ReadFile("config.ini")

        Dim defaultSection = data("Settings")
        defaultSection("keyfilePath") = TxtKeyPath.Text
        defaultSection("host") = TxtHost.Text
        defaultSection("port") = TxtPort.Text
        defaultSection("ytapikey") = TxtYtKey.Text
        defaultSection("username") = TxtUsername.Text
        defaultSection("keyPassphrase") = TxtPassphrase.Text
        defaultSection("BaroUpload") = If(ChkBaroUpload.Checked, "true", "false")
        defaultSection("ShutdownAfterUpload") = If(ChkShutdown.Checked, "true", "false")

        parser.WriteFile("config.ini", data)
        MessageBox.Show("설정이 저장되었습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information)
        HasConfigChanged = False

    End Sub

    Private Sub ChangeEnables(mode As String)

        Select Case mode
            Case "t"
                BtnLoadTXT.Enabled = True
                BtnMakeDB.Enabled = True
                BtnOpenFolder.Enabled = True
                BtnSaveConfig.Enabled = True
                BtnUploadDB.Enabled = True
                BtnTestSSH.Enabled = True
                BtnRestart.Enabled = True
                BtnUpdateDB.Enabled = True
                GroupBox1.Enabled = True
            Case "f"
                BtnLoadTXT.Enabled = False
                BtnMakeDB.Enabled = False
                BtnOpenFolder.Enabled = False
                BtnSaveConfig.Enabled = False
                BtnUploadDB.Enabled = False
                BtnTestSSH.Enabled = False
                BtnRestart.Enabled = False
                BtnUpdateDB.Enabled = False
                GroupBox1.Enabled = False
        End Select

    End Sub

    Private Async Function PreprocessTextFiles() As Task(Of Boolean)

        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Multiselect = True
        openFileDialog.Title = "대화 파일을 선택해주세요. (여러개 동시에 선택 가능)"
        openFileDialog.Filter = "카카오톡 대화 내보내기 파일 (*.txt)|*.txt"
        If openFileDialog.ShowDialog() <> DialogResult.OK Then Return False
        Dim txtfilePaths As String() = openFileDialog.FileNames

        ClearLists("ALL")
        Dim parser As New FileIniDataParser()
        Dim data As IniData = parser.ReadFile("config.ini")
        Dim allowedExtensions As String() = data("Settings")("allowedExtensions").Split(","c)

        Dim filePattern As String = "파일:\s+(.+)"
        Dim youtubePattern As String = "(?:youtube\.com\/(?:[^\/]+\/.+\/|(?:v|e(?:mbed)?)\/|.*[?&]v=)|youtu\.be\/)([^""&?\/ ]{11})"
        Dim shortsPattern As String = "(?:youtube\.com\/shorts\/)([^""&?\/ ]{11})"
        Dim applePattern As String = "(https:\/\/music\.apple\.com\/kr\/[^\s]+)"
        Dim datePattern As String = "--------------- (\d+년 \d+월 \d+일)"
        Dim namePattern As String = "\[([^\]]+)\]"

        Dim resultList As New List(Of List(Of String))()

        For Each txtfilePath As String In txtfilePaths
            Me.Invoke(Sub()
                          ListBox1.Items.Clear()
                          ListBox1.Items.Add("대화 파일을 불러오는 중...")
                          ChangeEnables("f")
                      End Sub)
            Dim lines As String() = Await File.ReadAllLinesAsync(txtfilePath)
            Dim currentDate As String = ""

            For Each line As String In lines

                Dim dateMatch As Match = Regex.Match(line, datePattern)
                If dateMatch.Success Then currentDate = dateMatch.Groups(1).Value

                Dim fileMatch As Match = Regex.Match(line, filePattern)
                If fileMatch.Success Then
                    Dim senderName As String = Regex.Match(line, namePattern).Groups(1).Value
                    Dim extension As String = Path.GetExtension(fileMatch.Groups(1).Value).ToLower()
                    If allowedExtensions.Contains(extension) AndAlso Not String.IsNullOrEmpty(senderName) Then resultList.Add(New List(Of String)({"파일", fileMatch.Groups(1).Value, currentDate, senderName}))
                End If

                Dim youtubeMatch As Match = Regex.Match(line, youtubePattern)
                If youtubeMatch.Success AndAlso youtubeMatch.Groups(1).Value.Length = 11 Then
                    Dim senderName As String = Regex.Match(line, namePattern).Groups(1).Value
                    If Not String.IsNullOrEmpty(senderName) Then resultList.Add(New List(Of String)({"유튜브", youtubeMatch.Groups(1).Value, currentDate, senderName}))
                End If

                Dim shortsMatch As Match = Regex.Match(line, shortsPattern)
                If shortsMatch.Success AndAlso shortsMatch.Groups(1).Value.Length = 11 Then
                    Dim senderName As String = Regex.Match(line, namePattern).Groups(1).Value
                    If Not String.IsNullOrEmpty(senderName) Then resultList.Add(New List(Of String)({"쇼츠", shortsMatch.Groups(1).Value, currentDate, senderName}))
                End If

                Dim appleMatch As Match = Regex.Match(line, applePattern)
                If appleMatch.Success Then
                    Dim senderName As String = Regex.Match(line, namePattern).Groups(1).Value
                    Dim appleURL As String = appleMatch.Groups(1).Value
                    If appleURL.EndsWith("\") Then
                        appleURL = appleURL.TrimEnd(New Char() {"\"c})
                    ElseIf appleURL.EndsWith("?l=en") Then
                        appleURL = appleURL.Substring(0, appleURL.Length - 5)
                    ElseIf appleURL.EndsWith("&ls") Then
                        appleURL = appleURL.Substring(0, appleURL.Length - 3)
                    End If
                    If Not String.IsNullOrEmpty(senderName) Then resultList.Add(New List(Of String)({"애플", appleURL, currentDate, senderName}))
                End If
            Next
        Next

        Dim dict As New Dictionary(Of String, Tuple(Of String, String, DateTime, String))
        Dim format As String = "yyyy년 M월 d일"
        Dim newFormat As String = "yyyy.MM.dd"

        For Each line As List(Of String) In resultList
            If line.Count >= 4 Then
                Dim dateValue As DateTime
                If DateTime.TryParseExact(line(2), format, CultureInfo.InvariantCulture, DateTimeStyles.None, dateValue) Then
                    If dict.ContainsKey(line(1)) Then
                        Dim existingData = dict(line(1))
                        If dateValue < existingData.Item3 Then dict(line(1)) = New Tuple(Of String, String, DateTime, String)(line(0), line(1), dateValue, line(3))
                    Else
                        dict.Add(line(1), New Tuple(Of String, String, DateTime, String)(line(0), line(1), dateValue, line(3)))
                    End If
                End If
            End If
        Next

        For Each kvp As KeyValuePair(Of String, Tuple(Of String, String, DateTime, String)) In dict
            Dim formattedDate As String = kvp.Value.Item3.ToString(newFormat)
            If kvp.Value.Item1 = "유튜브" Or kvp.Value.Item1 = "쇼츠" Then
                ytList.Add(New List(Of String)({kvp.Value.Item1, kvp.Value.Item2, formattedDate, kvp.Value.Item4}))
            ElseIf kvp.Value.Item1 = "애플" Then
                appleList.Add(New List(Of String)({kvp.Value.Item1, kvp.Value.Item2, formattedDate, kvp.Value.Item4}))
            ElseIf kvp.Value.Item1 = "파일" Then
                fileList.Add(New List(Of String)({kvp.Value.Item1, kvp.Value.Item2, formattedDate, kvp.Value.Item4}))
            End If
        Next

        Return True

    End Function

    Private Async Function PreprocessUpdate() As Task(Of Tuple(Of String, Boolean))

        Dim openFileDialog As New OpenFileDialog
        openFileDialog.InitialDirectory = Path.Combine(Application.StartupPath, "DB")
        openFileDialog.Title = "DB 파일을 선택해주세요."
        openFileDialog.Filter = "DB 파일 (*.db)|*.db"
        If openFileDialog.ShowDialog() <> DialogResult.OK Then Return New Tuple(Of String, Boolean)("", False)

        ClearLists()
        Dim dbPath As String = openFileDialog.FileName
        Using dbConn As New SQLiteConnection($"Data Source={dbPath};Version=3;")
            Await dbConn.OpenAsync()

            Dim ytDB As New List(Of String)
            Using dbCmd As New SQLiteCommand("SELECT 유튜브id FROM music_data", dbConn)
                Using reader As SQLiteDataReader = Await dbCmd.ExecuteReaderAsync()
                    While Await reader.ReadAsync() : ytDB.Add(reader("유튜브id").ToString()) : End While
                End Using
            End Using
            ytUpdateList = ytList.Where(Function(line) Not ytDB.Contains(line(1))).ToList()

            Dim appleDB As New List(Of String)
            Using dbCmd As New SQLiteCommand("SELECT 애플뮤직URL FROM music_data WHERE 유형 = '애플뮤직'", dbConn)
                Using reader As SQLiteDataReader = Await dbCmd.ExecuteReaderAsync()
                    While Await reader.ReadAsync() : appleDB.Add(reader("애플뮤직URL").ToString()) : End While
                End Using
            End Using
            appleUpdateList = appleList.Where(Function(line) Not appleDB.Contains(line(1))).ToList()

            Dim fileDB As New List(Of String)
            Using dbCmd As New SQLiteCommand("SELECT 제목 FROM music_data WHERE 유형='파일'", dbConn)
                Using reader As SQLiteDataReader = Await dbCmd.ExecuteReaderAsync()
                    While Await reader.ReadAsync() : fileDB.Add(reader("제목").ToString()) : End While
                End Using
            End Using
            fileUpdateList = fileList.Where(Function(line) Not fileDB.Contains(line(1))).ToList()

            dbConn.Close()
        End Using

        Return New Tuple(Of String, Boolean)(dbPath, True)

    End Function

    Private Function GetVideoInfo(videoId As String) As Tuple(Of String, String)

        Try
            Dim youtubeService = New YouTubeService(New BaseClientService.Initializer() With
        {
            .ApiKey = TxtYtKey.Text,
            .ApplicationName = Me.GetType().ToString()
        })

            Dim videoRequest = youtubeService.Videos.List("snippet")
            videoRequest.Id = videoId
            Dim videoResponse = videoRequest.Execute()

            If videoResponse.Items.Count > 0 Then
                Dim title As String = videoResponse.Items(0).Snippet.Title
                Dim channelName As String = videoResponse.Items(0).Snippet.ChannelTitle
                Return Tuple.Create(title, channelName)
            Else Return Tuple.Create("", "")
            End If

        Catch ex As Google.GoogleApiException
            If ex.HttpStatusCode = HttpStatusCode.BadRequest AndAlso ex.Message.Contains("API key not valid") Then
                MessageBox.Show("유튜브 API 키가 유효하지 않습니다.", "경고", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Else
                MessageBox.Show("유튜브 API 요청 중에 오류가 발생했습니다.", "경고", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
            cts.Cancel()
            Me.Invoke(Sub()
                          ChangeEnables("t")
                          BtnMakeDB.Enabled = False
                          BtnUpdateDB.Enabled = False
                          ClearLists("ALL")
                      End Sub)
        End Try

    End Function

    Private Function ProcessYT(ct As CancellationToken) As Tuple(Of List(Of List(Of String)), Integer)

        Dim processedCount As Integer = 0
        Dim resultList As New List(Of List(Of String))()

        Me.Invoke(Sub()
                      ListBox1.Items.Clear()
                      ListBox1.Items.Add($"{ytList.Count}개의 유튜브 정보를 받아오는 중...")
                      ListBox1.Items.Add("")
                  End Sub)

        Try
            For Each ytEntry As List(Of String) In ytList
                ct.ThrowIfCancellationRequested()
                processedCount += 1
                Me.Invoke(Sub()
                              ListBox1.Items(1) = $"({processedCount}/{ytList.Count})"
                          End Sub)

                Dim videoInfo As Tuple(Of String, String) = GetVideoInfo(ytEntry(1))
                If videoInfo IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(videoInfo.Item1) AndAlso Not String.IsNullOrWhiteSpace(videoInfo.Item2) Then
                    resultList.Add(New List(Of String)({ytEntry(0), ytEntry(2), ytEntry(3), videoInfo.Item1, videoInfo.Item2, ytEntry(1)}))
                End If
            Next

        Catch ex As OperationCanceledException
            Me.Invoke(Sub()
                          ListBox1.Items.Clear()
                          ListBox1.Items.Add("작업이 중단되었습니다.")
                      End Sub)
            ClearLists("ALL")
            Return Nothing
        End Try

        Return New Tuple(Of List(Of List(Of String)), Integer)(resultList, ytList.Count - resultList.Count)

    End Function

    Public Async Function ProcessApple(ct As CancellationToken) As Task(Of Tuple(Of List(Of List(Of String)), Integer))

        Dim processedcount As Integer = 0
        Dim resultList As New List(Of List(Of String))()
        Dim httpClient As New HttpClient()
        Dim pattern1 As New Regex("^(.+) - ([^의]+)의")
        Dim pattern2 As New Regex("^(.*?)(?: - 아티스트: (.*?))?(?: - 라디오 스테이션| - 플레이리스트| - Apple Music)")

        Me.Invoke(Sub()
                      ListBox1.Items.Clear()
                      ListBox1.Items.Add($"{appleList.Count}개의 애플뮤직 정보를 받아오는 중...")
                      ListBox1.Items.Add("")
                  End Sub)

        Try
            For Each appleEntry As List(Of String) In appleList
                ct.ThrowIfCancellationRequested()
                processedcount += 1
                Me.Invoke(Sub()
                              ListBox1.Items(1) = $"({processedcount}/{appleList.Count})"
                          End Sub)

                Try
                    Dim data As String = Await httpClient.GetStringAsync(appleEntry(1))
                    Dim doc As New HtmlDocument()
                    doc.LoadHtml(data)
                    Dim titleNode As HtmlNode = doc.DocumentNode.SelectSingleNode("//title")

                    If titleNode IsNot Nothing Then
                        Dim match As Match
                        match = If(titleNode.InnerText.Trim().Contains("의 앨범 - Apple Music") OrElse titleNode.InnerText.Trim().Contains("의 뮤직비디오 - Apple Music"), pattern1.Match(titleNode.InnerText.Trim()), pattern2.Match(titleNode.InnerText.Trim()))
                        If match.Success Then
                            Dim artist As String = If(match.Groups(2).Success, match.Groups(2).Value.Trim(), "")
                            resultList.Add(New List(Of String)({appleEntry(2), appleEntry(3), match.Groups(1).Value.Trim(), artist.Replace("&amp;", "&"), appleEntry(1)}))
                        End If
                    End If

                Catch ex As Exception
                End Try

                Await Task.Delay(350)
            Next

        Catch ex As OperationCanceledException
            Me.Invoke(Sub()
                          ListBox1.Items.Clear()
                          ListBox1.Items.Add("작업이 중단되었습니다.")
                      End Sub)
            ClearLists("ALL")
            Return Nothing
        End Try

        Return New Tuple(Of List(Of List(Of String)), Integer)(resultList, appleList.Count - resultList.Count)

    End Function

    Private Function WriteDB(ct As CancellationToken, ByVal mode As String, ByVal ytResultList As List(Of List(Of String)), ByVal appleResultList As List(Of List(Of String)), Optional ByVal exDbFilepath As String = "") As String

        Dim totalCount As Integer = ytResultList.Count + appleResultList.Count + fileList.Count
        Dim processedCount As Integer = 0

        Dim dbFileName As String = "DB_" & DateTime.Now.ToString("yyyyMMdd_HHmm") & ".db"
        Dim dbFilePath As String = Application.StartupPath & "\DB\" & dbFileName
        If mode.ToUpper() = "UPDATE" Then File.Copy(exDbFilepath, dbFilePath)

        Try
            Using conn As New SQLiteConnection($"Data Source={dbFilePath}")
                conn.Open()
                Using cmd As New SQLiteCommand(conn)

                    cmd.CommandText = "CREATE TABLE IF NOT EXISTS music_data (유형 TEXT, 제목 TEXT, 채널명 TEXT, 최초전송일 TEXT, 전송자 TEXT, 유튜브id TEXT, 애플뮤직URL TEXT)"
                    cmd.ExecuteNonQuery()
                    cmd.CommandText = "CREATE TABLE IF NOT EXISTS db_info (id INTEGER PRIMARY KEY, created_at DATETIME)"
                    cmd.ExecuteNonQuery()

                    If mode.ToUpper() = "CREATE" Then
                        cmd.CommandText = "INSERT INTO db_info (id, created_at) VALUES (?, ?)"
                        cmd.Parameters.AddWithValue("id", 1)
                        cmd.Parameters.AddWithValue("created_at", DateTime.Now.ToString("yyyy-MM-dd HH:mm"))
                        cmd.ExecuteNonQuery()
                        cmd.Parameters.Clear()
                    ElseIf mode.ToUpper() = "UPDATE" Then
                        cmd.CommandText = "UPDATE db_info SET created_at = ? WHERE id = 1"
                        cmd.Parameters.AddWithValue("created_at", DateTime.Now.ToString("yyyy-MM-dd HH:mm"))
                        cmd.ExecuteNonQuery()
                        cmd.Parameters.Clear()
                    End If

                    Me.Invoke(Sub()
                                  ListBox1.Items.Clear()
                                  ListBox1.Items.Add("DB 파일을 생성하는 중...")
                                  ListBox1.Items.Add("")
                              End Sub)

                    If ytResultList.Count > 0 Then
                        For Each ytItem In ytResultList
                            ct.ThrowIfCancellationRequested()
                            cmd.CommandText = "INSERT INTO music_data (유형, 제목, 채널명, 최초전송일, 전송자, 유튜브id, 애플뮤직URL) VALUES (?, ?, ?, ?, ?, ?, ?)"
                            cmd.Parameters.AddWithValue("유형", ytItem(0))
                            cmd.Parameters.AddWithValue("제목", ytItem(3))
                            cmd.Parameters.AddWithValue("채널명", ytItem(4))
                            cmd.Parameters.AddWithValue("최초전송일", ytItem(1))
                            cmd.Parameters.AddWithValue("전송자", ytItem(2))
                            cmd.Parameters.AddWithValue("유튜브id", ytItem(5))
                            cmd.Parameters.AddWithValue("애플뮤직URL", Nothing)
                            cmd.ExecuteNonQuery()
                            cmd.Parameters.Clear()
                            processedCount += 1
                            Me.Invoke(Sub()
                                          ListBox1.Items(1) = $"({processedCount}/{totalCount})"
                                      End Sub)
                        Next
                    End If

                    If appleResultList.Count > 0 Then
                        For Each appleItem In appleResultList
                            ct.ThrowIfCancellationRequested()
                            cmd.CommandText = "INSERT INTO music_data (유형, 제목, 채널명, 최초전송일, 전송자, 유튜브id, 애플뮤직URL) VALUES (?, ?, ?, ?, ?, ?, ?)"
                            cmd.Parameters.AddWithValue("유형", "애플뮤직")
                            cmd.Parameters.AddWithValue("제목", appleItem(2))
                            cmd.Parameters.AddWithValue("채널명", appleItem(3))
                            cmd.Parameters.AddWithValue("최초전송일", appleItem(0))
                            cmd.Parameters.AddWithValue("전송자", appleItem(1))
                            cmd.Parameters.AddWithValue("유튜브id", Nothing)
                            cmd.Parameters.AddWithValue("애플뮤직URL", appleItem(4))
                            cmd.ExecuteNonQuery()
                            cmd.Parameters.Clear()
                            processedCount += 1
                            Me.Invoke(Sub()
                                          ListBox1.Items(1) = $"({processedCount}/{totalCount})"
                                      End Sub)
                        Next
                    End If

                    If fileList.Count > 0 Then
                        For Each fileItem As List(Of String) In fileList
                            ct.ThrowIfCancellationRequested()
                            cmd.CommandText = "INSERT INTO music_data (유형, 제목, 채널명, 최초전송일, 전송자, 유튜브id, 애플뮤직URL) VALUES (?, ?, ?, ?, ?, ?, ?)"
                            cmd.Parameters.AddWithValue("유형", fileItem(0))
                            cmd.Parameters.AddWithValue("제목", fileItem(1))
                            cmd.Parameters.AddWithValue("채널명", Nothing)
                            cmd.Parameters.AddWithValue("최초전송일", fileItem(2))
                            cmd.Parameters.AddWithValue("전송자", fileItem(3))
                            cmd.Parameters.AddWithValue("유튜브id", Nothing)
                            cmd.Parameters.AddWithValue("애플뮤직URL", Nothing)
                            cmd.ExecuteNonQuery()
                            cmd.Parameters.Clear()
                            processedCount += 1
                            Me.Invoke(Sub()
                                          ListBox1.Items(1) = $"({processedCount}/{totalCount})"
                                      End Sub)
                        Next
                    End If

                End Using
                conn.Close()
            End Using

        Catch ex As OperationCanceledException
            Me.Invoke(Sub()
                          ListBox1.Items.Clear()
                          ListBox1.Items.Add("작업이 중단되었습니다.")
                      End Sub)
            If File.Exists(dbFilePath) Then File.Delete(dbFilePath)
            ClearLists("ALL")
            Return Nothing
        End Try

        Return dbFileName

    End Function

    Private Function UploadDB(dbFilePath As String) As Boolean

        Me.Invoke(Sub()
                      ChangeEnables("f")
                  End Sub)

        Dim parser As New FileIniDataParser()
        Dim data As IniData = parser.ReadFile("config.ini")

        Dim keyfilePath As String = TxtKeyPath.Text
        Dim username As String = TxtUsername.Text
        Dim host As String = TxtHost.Text
        Dim port As Integer = TxtPort.Text
        Dim appPath As String = data("Settings").GetKeyData("appPath").Value
        Dim startCommand As String = data("Settings").GetKeyData("startCommand").Value
        Dim stopCommand As String = data("Settings").GetKeyData("stopCommand").Value

        Dim keyFile As PrivateKeyFile

        Try
            keyFile = If(String.IsNullOrWhiteSpace(TxtPassphrase.Text), New PrivateKeyFile(keyfilePath), New PrivateKeyFile(keyfilePath, TxtPassphrase.Text))
        Catch ex As InvalidOperationException
            Me.Invoke(Sub()
                          MessageBox.Show("SSH 키의 Passphrase가 올바르지 않습니다." & vbCrLf & "다시 시도해주세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error)
                          ListBox1.Items.Clear()
                          ListBox1.Items.Add("서버 접속 실패")
                          ChangeEnables("t")
                          BtnMakeDB.Enabled = False
                          BtnUpdateDB.Enabled = False
                      End Sub)
            Return False
        End Try

        Dim keyFiles As New List(Of PrivateKeyFile) From {keyFile}
        Dim connInfo As New ConnectionInfo(host, port, username, New PrivateKeyAuthenticationMethod(username, keyFiles.ToArray()))

        Try
            Me.Invoke(Sub()
                          ListBox1.Items.Clear()
                          ListBox1.Items.Add("서버에 접속하는 중...")
                      End Sub)

            Using sshClient As New SshClient(connInfo)
                sshClient.Connect()

                Me.Invoke(Sub()
                              ListBox1.Items.Clear()
                              ListBox1.Items.Add("서버 접속 성공")
                              ListBox1.Items.Add("")
                              ListBox1.Items.Add("서비스를 중지하는 중...")
                          End Sub)

                sshClient.RunCommand(stopCommand)

                Me.Invoke(Sub()
                              ListBox1.Items.Add("DB 파일을 업로드 하는 중...")
                          End Sub)

                sshClient.RunCommand($"rm {appPath}/DB.db")

                Using scpClient As New ScpClient(connInfo)
                    scpClient.Connect()
                    Dim dbFile As New FileInfo(dbFilePath)
                    scpClient.Upload(dbFile, $"{appPath}/DB.db")
                End Using

                sshClient.RunCommand(startCommand)
                sshClient.Disconnect()

                Me.Invoke(Sub()
                              ListBox1.Items.Add("")
                              ListBox1.Items.Add("DB 파일 업로드 / 서비스 재시작 완료")
                              ListBox1.Items.Add("")
                              ListBox1.Items.Add("업로드 된 DB 파일 : " & Path.GetFileName(dbFilePath))
                              ChangeEnables("t")
                              BtnMakeDB.Enabled = False
                              BtnUpdateDB.Enabled = False
                          End Sub)

                If ChkShutdown.Checked = True Then
                    HasConfigChanged = False
                    System.Diagnostics.Process.Start("shutdown", "/s /t 0")
                End If

            End Using
            Return True

        Catch ex As Exception
            Me.Invoke(Sub()
                          ListBox1.Items.Clear()
                          ListBox1.Items.Add("서버 접속 실패")
                          ChangeEnables("t")
                          BtnMakeDB.Enabled = False
                          BtnUpdateDB.Enabled = False
                      End Sub)
            Return False
        End Try

    End Function

    Private Sub TestSSH()

        Dim result As DialogResult = MessageBox.Show("SSH 접속을 테스트하시겠습니까?", "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            Me.Invoke(Sub()
                          ChangeEnables("f")
                          ListBox1.Items.Clear()
                          ListBox1.Items.Add("서버에 접속하는 중...")
                      End Sub)

            Dim parser As New FileIniDataParser()
            Dim data As IniData = parser.ReadFile("config.ini")

            Dim keyfilePath As String = TxtKeyPath.Text
            Dim username As String = TxtUsername.Text
            Dim host As String = TxtHost.Text
            Dim port As Integer = TxtPort.Text

            Dim keyFile As PrivateKeyFile

            Try
                keyFile = If(String.IsNullOrWhiteSpace(TxtPassphrase.Text), New PrivateKeyFile(keyfilePath), New PrivateKeyFile(keyfilePath, TxtPassphrase.Text))

                Dim keyFiles As New List(Of PrivateKeyFile) From {keyFile}
                Dim connInfo As New ConnectionInfo(host, port, username, New PrivateKeyAuthenticationMethod(username, keyFiles.ToArray()))

                Using sshClient As New SshClient(connInfo)
                    sshClient.Connect()

                    Dim command As SshCommand = sshClient.RunCommand("whoami")
                    Dim output As String = command.Result.Trim()
                    Me.Invoke(Sub()
                                  ListBox1.Items.Clear()
                                  ListBox1.Items.Add("서버 접속 성공")
                                  ListBox1.Items.Add("")
                                  ListBox1.Items.Add("접속된 계정: " & output)
                              End Sub)

                    Dim serviceCommand As SshCommand = sshClient.RunCommand("sudo docker ps")
                    Dim serviceOutput As String = serviceCommand.Result.Trim()
                    If serviceOutput.Contains("hondaechun") Then
                        Me.Invoke(Sub()
                                      ListBox1.Items.Add("서비스 상태 : 실행 중")
                                  End Sub)
                    Else
                        Me.Invoke(Sub()
                                      ListBox1.Items.Add("서비스 상태 : 종료됨")
                                  End Sub)
                    End If

                    sshClient.Disconnect()

                End Using
            Catch ex As InvalidOperationException
                Me.Invoke(Sub()
                              MessageBox.Show("SSH 키의 Passphrase가 올바르지 않습니다." & vbCrLf & "다시 시도해주세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error)
                              ListBox1.Items.Clear()
                              ListBox1.Items.Add("서버 접속 실패")
                          End Sub)

                Return
            Catch ex As Exception
                Me.Invoke(Sub()
                              ListBox1.Items.Clear()
                              ListBox1.Items.Add("서버 접속 실패")
                          End Sub)

            Finally
                Me.Invoke(Sub()
                              ChangeEnables("t")
                              BtnMakeDB.Enabled = False
                              BtnUpdateDB.Enabled = False
                          End Sub)

            End Try
        End If

    End Sub

    Private Sub RestartService()

        Dim result As DialogResult = MessageBox.Show("서비스를 재시작 하시겠습니까?", "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            Me.Invoke(Sub()
                          ChangeEnables("f")
                          ListBox1.Items.Clear()
                          ListBox1.Items.Add("서버에 접속하는 중...")
                      End Sub)

            Dim parser As New FileIniDataParser()
            Dim data As IniData = parser.ReadFile("config.ini")

            Dim keyfilePath As String = TxtKeyPath.Text
            Dim username As String = TxtUsername.Text
            Dim host As String = TxtHost.Text
            Dim port As Integer = TxtPort.Text
            Dim startCommand As String = data("Settings").GetKeyData("startCommand").Value
            Dim stopCommand As String = data("Settings").GetKeyData("stopCommand").Value

            Dim keyFile As PrivateKeyFile

            Try
                keyFile = If(String.IsNullOrWhiteSpace(TxtPassphrase.Text), New PrivateKeyFile(keyfilePath), New PrivateKeyFile(keyfilePath, TxtPassphrase.Text))

                Dim keyFiles As New List(Of PrivateKeyFile) From {keyFile}
                Dim connInfo As New ConnectionInfo(host, port, username, New PrivateKeyAuthenticationMethod(username, keyFiles.ToArray()))

                Using sshClient As New SshClient(connInfo)
                    sshClient.Connect()
                    Me.Invoke(Sub()
                                  ListBox1.Items.Clear()
                                  ListBox1.Items.Add("서버 접속 성공")
                                  ListBox1.Items.Add("")
                                  ListBox1.Items.Add("서비스를 중지하는 중...")
                              End Sub)

                    Dim command As SshCommand = sshClient.RunCommand(stopCommand)

                    Me.Invoke(Sub()
                                  ListBox1.Items.Add("서비스를 시작하는 중...")
                              End Sub)

                    command = sshClient.RunCommand(startCommand)
                    sshClient.Disconnect()

                    Me.Invoke(Sub()
                                  ListBox1.Items.Add("")
                                  ListBox1.Items.Add("서비스 재시작이 완료되었습니다.")
                              End Sub)

                End Using

            Catch ex As InvalidOperationException
                Me.Invoke(Sub()
                              MessageBox.Show("SSH 키의 Passphrase가 올바르지 않습니다." & vbCrLf & "다시 시도해주세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error)
                              ListBox1.Items.Clear()
                              ListBox1.Items.Add("서버 접속 실패")
                          End Sub)
                Return
            Catch ex As Exception
                Me.Invoke(Sub()
                              ListBox1.Items.Clear()
                              ListBox1.Items.Add("서버 접속 실패")
                          End Sub)

            Finally
                Me.Invoke(Sub()
                              ChangeEnables("t")
                              BtnMakeDB.Enabled = False
                              BtnUpdateDB.Enabled = False
                          End Sub)
            End Try
        End If

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim dbFolderPath As String = Path.Combine(Application.StartupPath, "DB")
        If Not Directory.Exists(dbFolderPath) Then Directory.CreateDirectory(dbFolderPath)

        LoadConfig()
        ConfigLoadFinished = True
        BtnMakeDB.Enabled = False
        BtnUpdateDB.Enabled = False

    End Sub

    Private Sub Form1_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing

        If e.CloseReason = CloseReason.UserClosing Then
            If BtnMakeDB.Text = "중단하기" Then
                e.Cancel = True
                BtnMakeDB_Click(sender, e)
            ElseIf BtnUpdateDB.Text = "중단하기" Then
                e.Cancel = True
                BtnUpdateDB_Click(sender, e)
            End If
        End If

        If HasConfigChanged Then
            Dim result As DialogResult = MessageBox.Show("설정의 변경사항이 저장되지 않았습니다." & vbCrLf & "저장하시겠습니까?", "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then SaveConfig()
        End If

    End Sub

    Private Async Sub BtnLoadTXT_Click(sender As Object, e As EventArgs) Handles BtnLoadTXT.Click

        Dim isLoaded As Boolean = Await PreprocessTextFiles()
        ChangeEnables("t")
        If isLoaded = True Then
            If ytList.Count = 0 And appleList.Count = 0 And fileList.Count = 0 Then
                BtnUpdateDB.Enabled = False
                BtnMakeDB.Enabled = False
                MessageBox.Show("선택한 파일이 비어 있거나 유효하지 않습니다." & vbCrLf & "올바른 대화 파일을 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                ListBox1.Items.Clear()
                ListBox1.Items.Add("대화 파일 불러오기 실패")
            Else
                ListBox1.Items.Clear()
                ListBox1.Items.Add("대화 파일 불러오기 성공")
                ListBox1.Items.Add("")
                ListBox1.Items.Add($"파일 : {fileList.Count}개")
                ListBox1.Items.Add($"유튜브 : {ytList.Count}개")
                ListBox1.Items.Add($"애플뮤직 : {appleList.Count}개")
                ListBox1.Items.Add("")
                ListBox1.Items.Add("(중복 제외)")
            End If
        End If

    End Sub

    Private Async Sub BtnMakeDB_Click(sender As Object, e As EventArgs) Handles BtnMakeDB.Click

        If String.IsNullOrWhiteSpace(TxtYtKey.Text) Then
            MessageBox.Show("유튜브 API 키를 입력해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            If BtnMakeDB.Text = "DB 생성하기" Then
                Dim message As String = "최대 " & Regex.Match(ListBox1.Items(3).ToString(), "\d+").Value & "개의 유튜브 API 토큰이 사용될 수 있습니다." & vbCrLf & "계속하시겠습니까?"
                Dim result As DialogResult = MessageBox.Show(message, "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If result = DialogResult.Yes Then
                    ChangeEnables("f")
                    BtnMakeDB.Enabled = True

                    cts = New CancellationTokenSource()
                    Try
                        BtnMakeDB.Text = "중단하기"

                        Dim ytResult As Tuple(Of List(Of List(Of String)), Integer) = Await Task.Run(Function() ProcessYT(cts.Token), cts.Token)
                        Dim appleResult As Tuple(Of List(Of List(Of String)), Integer) = Await Task.Run(Function() ProcessApple(cts.Token), cts.Token)
                        Dim dbFileName As String = Await Task.Run(Function() WriteDB(cts.Token, "CREATE", ytResult.Item1, appleResult.Item1))

                        If Not cts.IsCancellationRequested Then
                            Dim uploadSuccess As Boolean = False
                            If ChkBaroUpload.Checked = True Then uploadSuccess = Await Task.Run(Function() UploadDB(Application.StartupPath & "\DB\" & dbFileName))
                            ListBox1.Items.Clear()
                            If ChkBaroUpload.Checked = True AndAlso uploadSuccess = True Then ListBox1.Items.Add("DB 생성 및 업로드를 완료하였습니다.") Else ListBox1.Items.Add("DB 생성을 완료하였습니다.")
                            ListBox1.Items.Add("")
                            ListBox1.Items.Add($"삭제 / 비공개 영상 : {ytResult.Item2}개")
                            ListBox1.Items.Add($"삭제 / 지역제한 애플뮤직 : {appleResult.Item2}개")
                            ListBox1.Items.Add("")
                            If ChkBaroUpload.Checked = True AndAlso uploadSuccess = True Then ListBox1.Items.Add($"업로드 된 DB 파일 : {dbFileName}") Else ListBox1.Items.Add($"생성된 파일명 : {dbFileName}")
                        End If

                        ClearLists("ALL")
                        ChangeEnables("t")
                        BtnUpdateDB.Enabled = False
                        BtnMakeDB.Enabled = False

                    Catch ex As OperationCanceledException
                    End Try
                    BtnMakeDB.Text = "DB 생성하기"
                End If
            Else
                Dim message As String = "DB 생성을 중단하시겠습니까?"
                Dim result As DialogResult = MessageBox.Show(message, "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If result = DialogResult.Yes And cts IsNot Nothing Then
                    cts.Cancel()
                    ChangeEnables("t")
                    BtnMakeDB.Enabled = False
                    BtnUpdateDB.Enabled = False
                End If
            End If
        End If

    End Sub

    Private Async Sub BtnUpdateDB_Click(sender As Object, e As EventArgs) Handles BtnUpdateDB.Click

        If String.IsNullOrWhiteSpace(TxtYtKey.Text) Then
            MessageBox.Show("유튜브 API 키를 입력해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            If BtnUpdateDB.Text = "DB 업데이트" Then
                Dim UpdateInfo As Tuple(Of String, Boolean) = Await PreprocessUpdate()
                If UpdateInfo.Item2 = False Then Exit Sub

                Dim message As String = "최대 " & ytUpdateList.Count & "개의 유튜브 API 토큰이 사용될 수 있습니다." & vbCrLf & "계속하시겠습니까?"
                Dim result As DialogResult = MessageBox.Show(message, "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If result = DialogResult.Yes Then
                    ChangeEnables("f")
                    BtnUpdateDB.Enabled = True

                    cts = New CancellationTokenSource()
                    Try
                        BtnUpdateDB.Text = "중단하기"
                        ytList.Clear()
                        appleList.Clear()
                        fileList.Clear()
                        ytList.AddRange(ytUpdateList)
                        appleList.AddRange(appleUpdateList)
                        fileList.AddRange(fileUpdateList)

                        Dim ytResult As Tuple(Of List(Of List(Of String)), Integer) = Await Task.Run(Function() ProcessYT(cts.Token), cts.Token)
                        Dim appleResult As Tuple(Of List(Of List(Of String)), Integer) = Await Task.Run(Function() ProcessApple(cts.Token), cts.Token)
                        Dim dbFileName As String = Await Task.Run(Function() WriteDB(cts.Token, "UPDATE", ytResult.Item1, appleResult.Item1, UpdateInfo.Item1))

                        If Not cts.IsCancellationRequested Then
                            Dim uploadSuccess As Boolean = False
                            If ChkBaroUpload.Checked = True Then uploadSuccess = Await Task.Run(Function() UploadDB(Application.StartupPath & "\DB\" & dbFileName))
                            ListBox1.Items.Clear()
                            If ChkBaroUpload.Checked = True AndAlso uploadSuccess = True Then ListBox1.Items.Add("DB 업데이트 및 업로드를 완료하였습니다.") Else ListBox1.Items.Add("DB 업데이트를 완료하였습니다.")
                            ListBox1.Items.Add("")
                            ListBox1.Items.Add($"업데이트 한 파일 : {fileList.Count}개")
                            ListBox1.Items.Add($"업데이트 한 유튜브 : {ytList.Count - ytResult.Item2}개")
                            ListBox1.Items.Add($"업데이트 한 애플뮤직 : {appleList.Count - appleResult.Item2}개")
                            ListBox1.Items.Add("")
                            If ChkBaroUpload.Checked = True AndAlso uploadSuccess = True Then ListBox1.Items.Add($"업로드 된 DB 파일 : {dbFileName}") Else ListBox1.Items.Add($"생성된 파일명 : {dbFileName}")
                        End If

                        ClearLists("ALL")
                        ChangeEnables("t")
                        BtnUpdateDB.Enabled = False
                        BtnMakeDB.Enabled = False

                    Catch ex As OperationCanceledException
                    End Try
                    BtnUpdateDB.Text = "DB 업데이트"
                End If
            Else
                Dim message As String = "DB 업데이트를 중단하시겠습니까?"
                Dim result As DialogResult = MessageBox.Show(message, "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If result = DialogResult.Yes And cts IsNot Nothing Then
                    cts.Cancel()
                    ChangeEnables("t")
                    BtnMakeDB.Enabled = False
                    BtnUpdateDB.Enabled = False
                End If
            End If
        End If

    End Sub

    Private Sub BtnOpenFolder_Click(sender As Object, e As EventArgs) Handles BtnOpenFolder.Click

        Try
            Dim myProcess = New Process()
            myProcess.StartInfo.UseShellExecute = True
            myProcess.StartInfo.FileName = Application.StartupPath & "\DB"
            myProcess.Start()
        Catch ex As Exception
            MessageBox.Show("폴더 열기 중 오류가 발생했습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub BtnLoadKey_Click(sender As Object, e As EventArgs) Handles BtnLoadKey.Click

        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Title = "SSH 키 파일을 선택해주세요."
        openFileDialog.Filter = "Key 파일|*.pem;*.key"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            TxtKeyPath.Text = openFileDialog.FileName
        Else
            MessageBox.Show("파일을 선택해주세요.", "경고", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

    End Sub

    Private Async Sub BtnTestSSH_Click(sender As Object, e As EventArgs) Handles BtnTestSSH.Click

        Await Task.Run(Sub() TestSSH())

    End Sub

    Private Async Sub BtnUploadDB_Click(sender As Object, e As EventArgs) Handles BtnUploadDB.Click

        Dim openFileDialog As New OpenFileDialog
        openFileDialog.InitialDirectory = Path.Combine(Application.StartupPath, "DB")
        openFileDialog.Title = "DB 파일을 선택해주세요."
        openFileDialog.Filter = "DB 파일 (*.db)|*.db"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim result As DialogResult = MessageBox.Show("DB 파일을 업로드 하시겠습니까?", "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                Await Task.Run(Sub() UploadDB(openFileDialog.FileName))
            ElseIf result = DialogResult.No Then
                MessageBox.Show("취소되었습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Else
            MessageBox.Show("파일을 선택해주세요.", "경고", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

    End Sub

    Private Async Sub BtnRestart_Click(sender As Object, e As EventArgs) Handles BtnRestart.Click

        Await Task.Run(Sub() RestartService())

    End Sub

    Private Sub BtnSaveConfig_Click(sender As Object, e As EventArgs) Handles BtnSaveConfig.Click

        Dim result As DialogResult = MessageBox.Show("설정을 저장하시겠습니까?", "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then SaveConfig()

    End Sub

    Private Sub TxtHost_TextChanged(sender As Object, e As EventArgs) Handles TxtHost.TextChanged
        If ConfigLoadFinished = True Then HasConfigChanged = True
    End Sub

    Private Sub TxtPort_TextChanged(sender As Object, e As EventArgs) Handles TxtPort.TextChanged
        If ConfigLoadFinished = True Then HasConfigChanged = True
    End Sub

    Private Sub TxtKeyPath_TextChanged(sender As Object, e As EventArgs) Handles TxtKeyPath.TextChanged
        If ConfigLoadFinished = True Then HasConfigChanged = True
    End Sub

    Private Sub TxtYtKey_TextChanged(sender As Object, e As EventArgs) Handles TxtYtKey.TextChanged
        If ConfigLoadFinished = True Then HasConfigChanged = True
    End Sub

    Private Sub TxtUsername_TextChanged(sender As Object, e As EventArgs) Handles TxtUsername.TextChanged
        If ConfigLoadFinished = True Then HasConfigChanged = True
    End Sub

    Private Sub TxtPassphrase_TextChanged(sender As Object, e As EventArgs) Handles TxtPassphrase.TextChanged
        If ConfigLoadFinished = True Then HasConfigChanged = True
    End Sub

    Private Sub ChkBaroUpload_CheckedChanged(sender As Object, e As EventArgs) Handles ChkBaroUpload.CheckedChanged
        If ConfigLoadFinished = True Then HasConfigChanged = True
    End Sub

    Private Sub ChkShutdown_CheckedChanged(sender As Object, e As EventArgs) Handles ChkShutdown.CheckedChanged
        If ConfigLoadFinished = True Then HasConfigChanged = True
    End Sub

End Class