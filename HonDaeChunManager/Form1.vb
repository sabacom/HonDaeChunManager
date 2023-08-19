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
    Private ConfigLoadFinished As Boolean = False
    Private HasConfigChanged As Boolean = False

    Private Sub DelTempfiles()

        Dim tempFolderPath As String = Path.Combine(Application.StartupPath, "DB\temp")
        If Directory.Exists(tempFolderPath) Then Directory.Delete(tempFolderPath, True)

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

    Private Async Function PreprocessTextFiles() As Task(Of Tuple(Of Integer, Integer, Integer))

        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Multiselect = True
        openFileDialog.Title = "대화 파일을 선택해주세요. (여러개 동시에 선택 가능)"
        openFileDialog.Filter = "카카오톡 대화 내보내기 파일 (*.txt)|*.txt"
        If openFileDialog.ShowDialog() <> DialogResult.OK Then Return New Tuple(Of Integer, Integer, Integer)(-1, -1, -1)

        DelTempfiles()
        Dim txtfilePaths As String() = openFileDialog.FileNames
        Dim tempFolderPath As String = Path.Combine(Application.StartupPath, "DB\temp")
        If Not Directory.Exists(tempFolderPath) Then Directory.CreateDirectory(tempFolderPath)

        Dim parser As New FileIniDataParser()
        Dim data As IniData = parser.ReadFile("config.ini")
        Dim allowedExtensions As String() = data("Settings")("allowedExtensions").Split(","c)

        ' 정규식 패턴
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
        Dim youtubeCount As Integer = 0
        Dim appleCount As Integer = 0
        Dim fileCount As Integer = 0

        For Each line As List(Of String) In resultList
            If line.Count >= 4 Then
                Dim contentType As String = line(0)
                Dim content As String = line(1)
                Dim datePart As String = line(2)
                Dim name As String = line(3)

                Dim dateValue As DateTime

                If DateTime.TryParseExact(datePart, format, CultureInfo.InvariantCulture, DateTimeStyles.None, dateValue) Then
                    If dict.ContainsKey(content) Then
                        Dim existingData = dict(content)
                        If dateValue < existingData.Item3 Then dict(content) = New Tuple(Of String, String, DateTime, String)(contentType, content, dateValue, name)
                    Else
                        dict.Add(content, New Tuple(Of String, String, DateTime, String)(contentType, content, dateValue, name))

                        If contentType = "유튜브" Or contentType = "쇼츠" Then
                            youtubeCount += 1
                        ElseIf contentType = "애플" Then
                            appleCount += 1
                        ElseIf contentType = "파일" Then
                            fileCount += 1
                        End If
                    End If
                End If
            End If
        Next

        Dim ytLines As New List(Of String)
        Dim appleLines As New List(Of String)
        Dim fileLines As New List(Of String)

        For Each kvp As KeyValuePair(Of String, Tuple(Of String, String, DateTime, String)) In dict
            Dim formattedDate As String = kvp.Value.Item3.ToString(newFormat)

            If kvp.Value.Item1 = "유튜브" Or kvp.Value.Item1 = "쇼츠" Then
                ytLines.Add(String.Join("/", kvp.Value.Item1, kvp.Value.Item2, formattedDate, kvp.Value.Item4))
            ElseIf kvp.Value.Item1 = "애플" Then
                appleLines.Add(String.Join("/", kvp.Value.Item1, formattedDate, kvp.Value.Item4))
                appleLines.Add(kvp.Value.Item2)
            ElseIf kvp.Value.Item1 = "파일" Then
                fileLines.Add(String.Join("/", kvp.Value.Item1, kvp.Value.Item2, formattedDate, kvp.Value.Item4))
            End If
        Next

        Await File.WriteAllLinesAsync(Application.StartupPath & "\DB\temp\yt.txt", ytLines)
        Await File.WriteAllLinesAsync(Application.StartupPath & "\DB\temp\apple.txt", appleLines)
        Await File.WriteAllLinesAsync(Application.StartupPath & "\DB\temp\files.txt", fileLines)

        Return New Tuple(Of Integer, Integer, Integer)(fileCount, youtubeCount, appleCount)

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
            End If

            Return Tuple.Create("", "")
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
                      End Sub)
        End Try

    End Function

    Private Function ProcessYT(ct As CancellationToken) As Tuple(Of List(Of List(Of String)), Integer)

        Dim filePath As String = Application.StartupPath & "\DB\temp\yt.txt"
        If Not File.Exists(filePath) Then Return New Tuple(Of List(Of List(Of String)), Integer)(New List(Of List(Of String))(), 0)

        Dim processedCount As Integer = 0
        Dim resultList As New List(Of List(Of String))()
        Dim regex As New Regex("^(.+?)/([^/]+)/([^/]+)/(.+)$")
        Dim lines As String() = File.ReadAllLines(filePath)
        Dim OriginLineCount As Integer = lines.Length

        Me.Invoke(Sub()
                      ListBox1.Items.Clear()
                      ListBox1.Items.Add($"{OriginLineCount}개의 유튜브 정보를 받아오는 중...")
                      ListBox1.Items.Add("")
                  End Sub)

        Try
            For Each line As String In lines
                ct.ThrowIfCancellationRequested()

                processedCount += 1
                Me.Invoke(Sub()
                              ListBox1.Items(1) = $"({processedCount}/{OriginLineCount})"
                          End Sub)

                Dim match As Match = regex.Match(line)

                If Not match.Success Then
                    resultList.Add(New List(Of String)({line}))
                    Continue For
                End If

                Dim contentType As String = match.Groups(1).Value
                Dim videoId As String = match.Groups(2).Value
                Dim datePart As String = match.Groups(3).Value
                Dim name As String = match.Groups(4).Value

                If contentType.ToLower() = "유튜브" Or contentType.ToLower() = "쇼츠" Then
                    Dim videoInfo As Tuple(Of String, String) = GetVideoInfo(videoId)
                    If videoInfo IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(videoInfo.Item1) AndAlso Not String.IsNullOrWhiteSpace(videoInfo.Item2) Then
                        resultList.Add(New List(Of String)({contentType, datePart, name, videoInfo.Item1, videoInfo.Item2, videoId}))
                    End If
                Else
                    resultList.Add(New List(Of String)({line}))
                End If
            Next

            Return New Tuple(Of List(Of List(Of String)), Integer)(resultList, OriginLineCount - resultList.Count)

        Catch ex As OperationCanceledException
            Me.Invoke(Sub()
                          ListBox1.Items.Clear()
                          ListBox1.Items.Add("작업이 중단되었습니다.")
                          ListBox1.Items.Add("")
                          DelTempfiles()
                      End Sub)
            Return Nothing
        End Try

    End Function

    Public Async Function ProcessApple(ct As CancellationToken) As Task(Of Tuple(Of List(Of List(Of String)), Integer))

        Dim filePath As String = Application.StartupPath & "\DB\temp\apple.txt"
        If Not File.Exists(filePath) Then Return New Tuple(Of List(Of List(Of String)), Integer)(New List(Of List(Of String))(), 0)

        Dim processedcount As Integer = 0
        Dim resultList As New List(Of List(Of String))()
        Dim httpClient As New HttpClient()
        Dim lines As String() = File.ReadAllLines(filePath)
        Dim originLineCount As Integer = lines.Length / 2

        Me.Invoke(Sub()
                      ListBox1.Items.Clear()
                      ListBox1.Items.Add($"{originLineCount}개의 애플뮤직 정보를 받아오는 중...")
                      ListBox1.Items.Add("")
                  End Sub)

        Try
            For i As Integer = 0 To lines.Length - 1 Step 2
                ct.ThrowIfCancellationRequested()

                processedcount += 1
                Me.Invoke(Sub()
                              ListBox1.Items(1) = $"({processedcount}/{originLineCount})"
                          End Sub)

                Dim metainfo As String = lines(i)
                Dim url As String = lines(i + 1)

                Try
                    Dim data As String = Await httpClient.GetStringAsync(url)
                    Dim doc As New HtmlDocument()
                    doc.LoadHtml(data)
                    Dim titleNode As HtmlNode = doc.DocumentNode.SelectSingleNode("//title")

                    If titleNode IsNot Nothing Then
                        Dim title As String = titleNode.InnerText.Trim()
                        Dim metaInfoParts As Match = Regex.Match(metainfo, "(애플)/([^/]+)/(.+)")
                        resultList.Add(New List(Of String)({"애플", metaInfoParts.Groups(2).Value, metaInfoParts.Groups(3).Value, title, url}))
                    End If
                Catch ex As Exception

                End Try

                Await Task.Delay(350)
            Next

        Catch ex As OperationCanceledException
            Me.Invoke(Sub()
                          ListBox1.Items.Clear()
                          ListBox1.Items.Add("작업이 중단되었습니다.")
                          ListBox1.Items.Add("")
                      End Sub)
            Return Nothing
        End Try

        Return New Tuple(Of List(Of List(Of String)), Integer)(resultList, originLineCount - resultList.Count)

    End Function

    Private Function PreprocessUpdate() As Tuple(Of String, Integer, Integer, Integer)

        Dim openFileDialog As New OpenFileDialog
        openFileDialog.InitialDirectory = Path.Combine(Application.StartupPath, "DB")
        openFileDialog.Title = "DB 파일을 선택해주세요."
        openFileDialog.Filter = "DB 파일 (*.db)|*.db"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim dbPath As String = openFileDialog.FileName
            Dim baseDir As String = Application.StartupPath & "\DB\temp\"
            Dim ytCount As Integer = 0
            Dim appleCount As Integer = 0
            Dim fileCount As Integer = 0

            Using dbConn As New SQLiteConnection($"Data Source={dbPath};Version=3;")
                dbConn.Open()

                If File.Exists(baseDir & "yt.txt") Then
                    Dim ytDB As New List(Of String)
                    Using dbCmd As New SQLiteCommand("SELECT 유튜브id FROM music_data", dbConn)
                        Using reader As SQLiteDataReader = dbCmd.ExecuteReader()
                            While reader.Read() : ytDB.Add(reader("유튜브id").ToString()) : End While
                        End Using
                    End Using

                    File.Copy(baseDir & "yt.txt", baseDir & "yt_update.txt", True)
                    Dim txtYTLines As List(Of String) = File.ReadAllLines(baseDir & "yt_update.txt").ToList()
                    For Each youtubeId In ytDB : txtYTLines.RemoveAll(Function(line) line.Contains($"/{youtubeId}/")) : Next
                    File.WriteAllLines(baseDir & "yt_update.txt", txtYTLines)
                    ytCount = txtYTLines.Count
                End If

                If File.Exists(baseDir & "apple.txt") Then
                    Dim appleDB As New List(Of String)
                    Using dbCmd As New SQLiteCommand("SELECT 애플뮤직URL FROM music_data WHERE 유형 = '애플뮤직'", dbConn)
                        Using reader As SQLiteDataReader = dbCmd.ExecuteReader()
                            While reader.Read() : appleDB.Add(reader("애플뮤직URL").ToString()) : End While
                        End Using
                    End Using

                    File.Copy(baseDir & "apple.txt", baseDir & "apple_update.txt", True)
                    Dim txtAppleLines As List(Of String) = File.ReadAllLines(baseDir & "apple_update.txt").ToList()
                    For Each appleMusicUrl In appleDB
                        For i As Integer = 0 To txtAppleLines.Count - 2 Step 2
                            If txtAppleLines(i + 1) = appleMusicUrl Then
                                txtAppleLines.RemoveAt(i + 1)
                                txtAppleLines.RemoveAt(i)
                                Exit For
                            End If
                        Next
                    Next
                    File.WriteAllLines(baseDir & "apple_update.txt", txtAppleLines)
                    appleCount = txtAppleLines.Count / 2
                End If

                If File.Exists(baseDir & "files.txt") Then
                    Dim filesDB As New List(Of String)
                    Using dbCmd As New SQLiteCommand("SELECT 제목 FROM music_data WHERE 유형='파일'", dbConn)
                        Using reader As SQLiteDataReader = dbCmd.ExecuteReader()
                            While reader.Read() : filesDB.Add(reader("제목").ToString()) : End While
                        End Using
                    End Using

                    File.Copy(baseDir & "files.txt", baseDir & "files_update.txt", True)
                    Dim txtFileLines As List(Of String) = File.ReadAllLines(baseDir & "files_update.txt").ToList()
                    For Each title In filesDB : txtFileLines.RemoveAll(Function(line) line.Contains($"/{title}/")) : Next
                    File.WriteAllLines(baseDir & "files_update.txt", txtFileLines)
                    fileCount = txtFileLines.Count
                End If

                dbConn.Close()

            End Using
            Return New Tuple(Of String, Integer, Integer, Integer)(dbPath, fileCount, ytCount, appleCount)

        End If

    End Function

    Private Function WriteDB(ByVal mode As String, ByVal ytData As List(Of List(Of String)), ByVal appleData As List(Of List(Of String)), Optional ByVal exDbFilepath As String = "") As String

        Dim fileCount As Integer = If(File.Exists(Path.Combine(Application.StartupPath, "DB\temp", "files.txt")), File.ReadLines(Path.Combine(Application.StartupPath, "DB\temp", "files.txt")).Count(), 0)
        Dim totalCount As Integer = ytData.Count + appleData.Count + fileCount
        Dim processedCount As Integer = 0

        Dim dbFileName As String = "DB_" & DateTime.Now.ToString("yyyyMMdd_HHmm") & ".db"
        Dim dbFilePath As String = Application.StartupPath & "\DB\" & dbFileName
        If mode.ToUpper() = "UPDATE" Then File.Copy(exDbFilepath, dbFilePath)

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

                If ytData.Count > 0 Then
                    For Each ytItem In ytData
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

                If appleData.Count > 0 Then
                    For Each appleItem In appleData
                        Dim Title As String = appleItem(3)
                        Title = Regex.Replace(Title, "\u200E", String.Empty)
                        Title = Title.Replace("Apple Music에서 감상하는 ", String.Empty)
                        Title = Title.Replace("Apple Music에서 만나는 ", String.Empty)
                        Title = Title.Replace("Apple Music의 ", String.Empty)
                        Title = Title.Replace("&amp;", "&")

                        cmd.CommandText = "INSERT INTO music_data (유형, 제목, 채널명, 최초전송일, 전송자, 유튜브id, 애플뮤직URL) VALUES (?, ?, ?, ?, ?, ?, ?)"
                        cmd.Parameters.AddWithValue("유형", "애플뮤직")
                        cmd.Parameters.AddWithValue("제목", Title)
                        cmd.Parameters.AddWithValue("채널명", Nothing)
                        cmd.Parameters.AddWithValue("최초전송일", appleItem(1))
                        cmd.Parameters.AddWithValue("전송자", appleItem(2))
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

                If File.Exists(Application.StartupPath & "\DB\temp\files.txt") Then
                    Dim fileLines As List(Of String) = File.ReadLines(Application.StartupPath & "\DB\temp\files.txt").ToList()
                    Dim regex As New Regex("^(.+?)/([^/]+)/([^/]+)/(.+)$")

                    For Each line As String In fileLines
                        Dim match As Match = regex.Match(line)
                        If Not match.Success Then Continue For

                        cmd.CommandText = "INSERT INTO music_data (유형, 제목, 채널명, 최초전송일, 전송자, 유튜브id, 애플뮤직URL) VALUES (?, ?, ?, ?, ?, ?, ?)"
                        cmd.Parameters.AddWithValue("유형", "파일")
                        cmd.Parameters.AddWithValue("제목", match.Groups(2).Value)
                        cmd.Parameters.AddWithValue("채널명", Nothing)
                        cmd.Parameters.AddWithValue("최초전송일", match.Groups(3).Value)
                        cmd.Parameters.AddWithValue("전송자", match.Groups(4).Value)
                        cmd.Parameters.AddWithValue("유튜브id", Nothing)
                        cmd.Parameters.AddWithValue("애플뮤직URL", Nothing)
                        cmd.ExecuteNonQuery()

                        processedCount += 1
                        Me.Invoke(Sub()
                                      ListBox1.Items(1) = $"({processedCount}/{totalCount})"
                                  End Sub)
                    Next

                    cmd.Parameters.Clear()
                End If
            End Using
            conn.Close()
        End Using

        Return dbFileName

    End Function

    Private Sub UploadDB(dbFilePath As String)

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
            Return
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

                If ChkShutdown.Checked = True Then System.Diagnostics.Process.Start("shutdown", "/s /t 0")

            End Using

        Catch ex As Exception
            Me.Invoke(Sub()
                          ListBox1.Items.Clear()
                          ListBox1.Items.Add("서버 접속 실패")
                          ChangeEnables("t")
                          BtnMakeDB.Enabled = False
                          BtnUpdateDB.Enabled = False
                      End Sub)
            Return
        End Try

    End Sub

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

        DelTempfiles()
        LoadConfig()
        BtnMakeDB.Enabled = False
        BtnUpdateDB.Enabled = False

        Dim dbFolderPath As String = Path.Combine(Application.StartupPath, "DB")
        If Not Directory.Exists(dbFolderPath) Then Directory.CreateDirectory(dbFolderPath)

        ConfigLoadFinished = True

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

        DelTempfiles()

    End Sub

    Private Async Sub BtnLoadTXT_Click(sender As Object, e As EventArgs) Handles BtnLoadTXT.Click

        Dim Counts As Tuple(Of Integer, Integer, Integer) = Await PreprocessTextFiles()
        If Counts.Item1 = 0 And Counts.Item2 = 0 And Counts.Item3 = 0 Then
            DelTempfiles()
            ChangeEnables("t")
            BtnUpdateDB.Enabled = False
            BtnMakeDB.Enabled = False

            MessageBox.Show("파일이 비어 있거나 유효하지 않습니다." & vbCrLf & "올바른 대화 파일을 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            ListBox1.Items.Clear()
            ListBox1.Items.Add("대화 파일 불러오기 실패")

        ElseIf Counts.Item1 <> -1 Then
            ChangeEnables("t")
            ListBox1.Items.Clear()
            ListBox1.Items.Add("대화 파일 불러오기 성공")
            ListBox1.Items.Add("")
            ListBox1.Items.Add($"파일 : {Counts.Item1}개")
            ListBox1.Items.Add($"유튜브 : {Counts.Item2}개")
            ListBox1.Items.Add($"애플뮤직 : {Counts.Item3}개")
            ListBox1.Items.Add("")
            ListBox1.Items.Add("(중복 제외)")
        End If

    End Sub

    Private Async Sub BtnMakeDB_Click(sender As Object, e As EventArgs) Handles BtnMakeDB.Click

        If String.IsNullOrWhiteSpace(TxtYtKey.Text) Then
            MessageBox.Show("유튜브 API 키를 입력해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            If BtnMakeDB.Text = "DB 생성하기" Then
                Dim ytnumber As String = Regex.Match(ListBox1.Items(3).ToString(), "\d+").Value
                Dim message As String = "최대 " & ytnumber & "개의 유튜브 API 토큰이 사용될 수 있습니다." & vbCrLf & "계속하시겠습니까?"
                Dim result As DialogResult = MessageBox.Show(message, "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                If result = DialogResult.Yes Then
                    ChangeEnables("f")
                    BtnMakeDB.Enabled = True

                    cts = New CancellationTokenSource()
                    Try
                        BtnMakeDB.Text = "중단하기"

                        Dim ytResult As Tuple(Of List(Of List(Of String)), Integer) = Await Task.Run(Function() ProcessYT(cts.Token), cts.Token)
                        Dim appleResult As Tuple(Of List(Of List(Of String)), Integer) = Await Task.Run(Function() ProcessApple(cts.Token), cts.Token)

                        If Not cts.IsCancellationRequested Then
                            BtnMakeDB.Enabled = False

                            Dim dbFileName As String = Await Task.Run(Function() WriteDB("CREATE", ytResult.Item1, appleResult.Item1))
                            If ChkBaroUpload.Checked = True Then Await Task.Run(Sub() UploadDB(Application.StartupPath & "\DB\" & dbFileName))

                            ListBox1.Items.Clear()
                            If ChkBaroUpload.Checked = True Then ListBox1.Items.Add("DB 생성 및 업로드를 완료하였습니다.") Else ListBox1.Items.Add("DB 생성을 완료하였습니다.")
                            ListBox1.Items.Add("")
                            ListBox1.Items.Add($"삭제 / 비공개 영상 : {ytResult.Item2}개")
                            ListBox1.Items.Add($"삭제 / 지역제한 애플뮤직 : {appleResult.Item2}개")
                            ListBox1.Items.Add("")
                            If ChkBaroUpload.Checked = True Then ListBox1.Items.Add($"업로드 된 DB 파일 : {dbFileName}") Else ListBox1.Items.Add($"생성된 파일명 : {dbFileName}")

                            ChangeEnables("t")
                            BtnUpdateDB.Enabled = False
                            BtnMakeDB.Enabled = False
                            DelTempfiles()
                        End If
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
                Dim dbInfo As Tuple(Of String, Integer, Integer, Integer) = PreprocessUpdate()
                If dbInfo.Item2 = -1 Then Exit Sub
                Dim message As String = "최대 " & dbInfo.Item3 & "개의 유튜브 API 토큰이 사용될 수 있습니다." & vbCrLf & "계속하시겠습니까?"
                Dim result As DialogResult = MessageBox.Show(message, "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                If result = DialogResult.Yes Then
                    ChangeEnables("f")
                    BtnUpdateDB.Enabled = True

                    cts = New CancellationTokenSource()
                    Try
                        BtnUpdateDB.Text = "중단하기"

                        Dim baseDir As String = Application.StartupPath & "\DB\temp\"
                        For Each fname In {"apple", "files", "yt"}
                            If File.Exists(baseDir & fname & ".txt") Then
                                File.Delete(baseDir & fname & ".txt")
                                File.Move(baseDir & fname & "_update.txt", baseDir & fname & ".txt")
                            End If
                        Next

                        Dim ytResult As Tuple(Of List(Of List(Of String)), Integer) = Await Task.Run(Function() ProcessYT(cts.Token), cts.Token)
                        Dim appleResult As Tuple(Of List(Of List(Of String)), Integer) = Await Task.Run(Function() ProcessApple(cts.Token), cts.Token)

                        If Not cts.IsCancellationRequested Then
                            BtnUpdateDB.Enabled = False

                            Dim dbFileName As String = Await Task.Run(Function() WriteDB("UPDATE", ytResult.Item1, appleResult.Item1, dbInfo.Item1))
                            If ChkBaroUpload.Checked = True Then Await Task.Run(Sub() UploadDB(Application.StartupPath & "\DB\" & dbFileName))

                            ListBox1.Items.Clear()
                            If ChkBaroUpload.Checked = True Then ListBox1.Items.Add("DB 업데이트 및 업로드를 완료하였습니다.") Else ListBox1.Items.Add("DB 업데이트를 완료하였습니다.")
                            ListBox1.Items.Add("")
                            ListBox1.Items.Add($"업데이트 한 파일 : {dbInfo.Item2}개")
                            ListBox1.Items.Add($"업데이트 한 유튜브 : {dbInfo.Item3 - ytResult.Item2}개")
                            ListBox1.Items.Add($"업데이트 한 애플뮤직 : {dbInfo.Item4 - appleResult.Item2}개")
                            ListBox1.Items.Add("")
                            If ChkBaroUpload.Checked = True Then ListBox1.Items.Add($"업로드 된 DB 파일 : {dbFileName}") Else ListBox1.Items.Add($"생성된 파일명 : {dbFileName}")

                            ChangeEnables("t")
                            BtnUpdateDB.Enabled = False
                            BtnMakeDB.Enabled = False
                            DelTempfiles()
                        End If
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