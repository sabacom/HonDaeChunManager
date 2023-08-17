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

    Private Function PreprocessTextFiles() As List(Of String)

        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Multiselect = True '
        openFileDialog.Title = "대화 파일을 선택해주세요. (여러개 동시에 선택 가능)"
        openFileDialog.Filter = "카카오톡 대화 내보내기 파일 (*.txt)|*.txt"
        If openFileDialog.ShowDialog() <> DialogResult.OK Then Return New List(Of String)()

        Dim filePaths As String() = openFileDialog.FileNames
        DelTempfiles()

        Dim parser As New FileIniDataParser()
        Dim data As IniData = parser.ReadFile("config.ini")
        Dim allowedExtensions As String() = data("Settings")("allowedExtensions").Split(","c)

        Dim tempFolderPath As String = Path.Combine(Application.StartupPath, "DB\temp")
        If Not Directory.Exists(tempFolderPath) Then Directory.CreateDirectory(tempFolderPath)

        ' 정규식 패턴
        Dim filePattern As String = "파일:\s+(.+)"
        Dim youtubePattern As String = "(?:youtube\.com\/(?:[^\/]+\/.+\/|(?:v|e(?:mbed)?)\/|.*[?&]v=)|youtu\.be\/)([^""&?\/ ]{11})"
        Dim ShortsPattern As String = "(?:youtube\.com\/shorts\/)([^""&?\/ ]{11})"
        Dim applePattern As String = "(https:\/\/music\.apple\.com\/kr\/[^\s]+)"
        Dim datePattern As String = "--------------- (\d+년 \d+월 \d+일)"
        Dim namePattern As String = "\[([^\]]+)\]"

        Dim extractedLines As New List(Of String)()

        For Each filePath As String In filePaths
            Dim lines As String() = File.ReadAllLines(filePath)
            Dim currentDate As String = ""

            For Each line As String In lines

                Dim dateMatch As Match = Regex.Match(line, datePattern)
                If dateMatch.Success Then currentDate = dateMatch.Groups(1).Value

                Dim fileMatch As Match = Regex.Match(line, filePattern)
                If fileMatch.Success Then
                    Dim fileName As String = fileMatch.Groups(1).Value
                    Dim nameMatch As Match = Regex.Match(line, namePattern)
                    Dim senderName As String = ""
                    If nameMatch.Success Then senderName = nameMatch.Groups(1).Value

                    Dim extension As String = Path.GetExtension(fileName).ToLower()
                    If allowedExtensions.Contains(extension) Then
                        Dim extractedLine As String = $"파일/{fileName}/{currentDate}/{senderName}"
                        extractedLines.Add(extractedLine)
                    End If
                End If

                Dim youtubeMatch As Match = Regex.Match(line, youtubePattern)
                If youtubeMatch.Success Then
                    Dim youtubeLink As String = youtubeMatch.Groups(1).Value
                    If youtubeLink.Length <> 11 Then Continue For
                    Dim nameMatch As Match = Regex.Match(line, namePattern)
                    Dim senderName As String = ""
                    If nameMatch.Success Then senderName = nameMatch.Groups(1).Value

                    Dim extractedLine As String = $"유튜브/{youtubeLink}/{currentDate}/{senderName}"
                    extractedLines.Add(extractedLine)
                End If

                Dim ShortsMatch As Match = Regex.Match(line, ShortsPattern)
                If ShortsMatch.Success Then
                    Dim youtubeLink As String = ShortsMatch.Groups(1).Value
                    If youtubeLink.Length <> 11 Then Continue For
                    Dim nameMatch As Match = Regex.Match(line, namePattern)
                    Dim senderName As String = ""
                    If nameMatch.Success Then senderName = nameMatch.Groups(1).Value

                    Dim extractedLine As String = $"쇼츠/{youtubeLink}/{currentDate}/{senderName}"
                    extractedLines.Add(extractedLine)
                End If

                Dim appleMatch As Match = Regex.Match(line, applePattern)
                If appleMatch.Success Then
                    Dim appleLink As String = appleMatch.Groups(1).Value
                    Dim nameMatch As Match = Regex.Match(line, namePattern)
                    Dim senderName As String = ""
                    If nameMatch.Success Then senderName = nameMatch.Groups(1).Value

                    Dim extractedLine As String = $"애플/{currentDate}/{senderName}"
                    extractedLines.Add(extractedLine)
                    File.AppendAllText(Application.StartupPath & "\DB\temp\apple.txt", extractedLine & Environment.NewLine)
                    File.AppendAllText(Application.StartupPath & "\DB\temp\apple.txt", appleLink & Environment.NewLine)
                End If

            Next
        Next

        Return extractedLines

    End Function

    Private Function RemoveAndSplit(lines As List(Of String)) As Tuple(Of Integer, Integer)

        Dim dict As New Dictionary(Of String, String)
        Dim format As String = "yyyy년 M월 d일"
        Dim newFormat As String = "yyyy.MM.dd"
        Dim fileCount As Integer = 0
        Dim youtubeCount As Integer = 0

        For Each line As String In lines
            Dim parts As String() = line.Split("/"c)

            If parts.Length = 4 Then
                Dim contentType As String = parts(0)
                Dim content As String = parts(1)
                Dim datePart As String = parts(2)
                Dim name As String = parts(3)

                Dim dateValue As DateTime

                If DateTime.TryParseExact(datePart, format, CultureInfo.InvariantCulture, DateTimeStyles.None, dateValue) Then
                    datePart = dateValue.ToString(newFormat)

                    If dict.ContainsKey(content) Then
                        Dim existingLine As String = dict(content)
                        Dim existingDatePart As String = existingLine.Split("/"c)(2)
                        Dim existingDateValue As DateTime

                        If DateTime.TryParseExact(existingDatePart, newFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, existingDateValue) Then
                            If dateValue < existingDateValue Then dict(content) = String.Join("/", contentType, content, datePart, name)
                        End If
                    Else
                        dict.Add(content, String.Join("/", contentType, content, datePart, name))

                        If contentType = "파일" Then
                            fileCount += 1
                        ElseIf contentType = "유튜브" Or contentType = "쇼츠" Then
                            youtubeCount += 1
                        End If
                    End If
                End If
            End If
        Next

        Dim ytPath As String = Application.StartupPath & "\DB\temp\yt.txt"
        Dim filePath As String = Application.StartupPath & "\DB\temp\files.txt"

        Dim ytLines As New List(Of String)
        Dim fileLines As New List(Of String)

        For Each kvp As KeyValuePair(Of String, String) In dict
            If kvp.Value.StartsWith("유튜브/") Or kvp.Value.StartsWith("쇼츠/") Then
                ytLines.Add(kvp.Value)
            ElseIf kvp.Value.StartsWith("파일/") Then
                fileLines.Add(kvp.Value)
            End If
        Next

        File.WriteAllLines(ytPath, ytLines)
        File.WriteAllLines(filePath, fileLines)

        Return New Tuple(Of Integer, Integer)(fileCount, youtubeCount)

    End Function

    Private Function RemoveApple() As Integer

        Dim FilePath As String = Application.StartupPath & "\DB\temp\apple.txt"
        If Not File.Exists(FilePath) Then Return 0

        Dim lines As List(Of String) = File.ReadAllLines(FilePath).ToList()

        Dim dict As New Dictionary(Of String, String)
        Dim format As String = "yyyy년 M월 d일"
        Dim newFormat As String = "yyyy.MM.dd"

        For i As Integer = 0 To lines.Count - 1 Step 2
            Dim line1 As String = lines(i)
            Dim line2 As String = lines(i + 1)

            Dim parts1 As String() = line1.Split("/"c)

            If parts1.Length = 3 Then
                Dim contentType As String = parts1(0)
                Dim datePart As String = parts1(1)
                Dim name As String = parts1(2)
                Dim url As String = line2

                If url.EndsWith("\") Then
                    url = url.TrimEnd(New Char() {"\"c})
                ElseIf url.EndsWith("?l=en") Then
                    url = url.Substring(0, url.Length - 5)
                ElseIf url.EndsWith("&ls") Then
                    url = url.Substring(0, url.Length - 3)
                End If

                Dim dateValue As DateTime

                If DateTime.TryParseExact(datePart, format, CultureInfo.InvariantCulture, DateTimeStyles.None, dateValue) Then
                    datePart = dateValue.ToString(newFormat)
                    If Not dict.ContainsKey(url) Then dict.Add(url, String.Join("/", contentType, datePart, name))
                End If
            End If
        Next

        Dim outputFile As String = Application.StartupPath & "\DB\temp\apple.txt"
        Dim outputLines As New List(Of String)()

        For Each kvp As KeyValuePair(Of String, String) In dict
            outputLines.Add(kvp.Value)
            outputLines.Add(kvp.Key)
        Next

        File.WriteAllLines(outputFile, outputLines)

        Return dict.Count

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

    Private Function ProcessYT(ct As CancellationToken) As Integer

        Dim filePath As String = Application.StartupPath & "\DB\temp\yt.txt"
        If Not File.Exists(filePath) Then Return 0

        Dim originalLineCount As Integer = File.ReadAllLines(filePath).Length
        Dim lines As String() = File.ReadAllLines(filePath)

        Dim processedLines As New List(Of String)()

        Me.Invoke(Sub()
                      ListBox1.Items.Clear()
                      ListBox1.Items.Add($"{originalLineCount}개의 유튜브 정보를 받아오는 중...")
                      ListBox1.Items.Add("")
                  End Sub)

        Dim processedCount As Integer = 1
        Me.Invoke(Sub()
                      ListBox1.Items(1) = $"({processedCount}/{originalLineCount})"
                      ListBox1.TopIndex = ListBox1.Items.Count - 1
                  End Sub)
        Try
            For Each line As String In lines
                ct.ThrowIfCancellationRequested()
                Dim parts As String() = line.Split("/")
                If parts.Length < 4 Then
                    processedLines.Add(line)
                    Continue For
                End If

                Dim contentType As String = parts(0)
                Dim videoId As String = parts(1)
                Dim datePart As String = parts(2)
                Dim name As String = parts(3)

                If contentType.ToLower() = "유튜브" Or contentType.ToLower() = "쇼츠" Then
                    Dim videoInfo As Tuple(Of String, String) = GetVideoInfo(videoId)
                    processedCount += 1

                    Me.Invoke(Sub()
                                  ListBox1.Items(1) = $"({processedCount}/{originalLineCount})"
                                  ListBox1.TopIndex = ListBox1.Items.Count - 1
                              End Sub)

                    If videoInfo IsNot Nothing Then
                        Dim title As String = videoInfo.Item1
                        Dim channelName As String = videoInfo.Item2
                        If Not String.IsNullOrWhiteSpace(title) And Not String.IsNullOrWhiteSpace(channelName) Then
                            processedLines.Add($"{contentType}/{datePart}/{name}/{videoId}/{title}")
                            processedLines.Add(channelName)
                        End If
                    End If
                Else
                    processedLines.Add(line)
                End If
            Next

            File.WriteAllLines(filePath, processedLines)
            Dim afterLineCount As Integer = File.ReadAllLines(filePath).Length
            Dim difference As Integer = originalLineCount - (afterLineCount / 2)

            Return difference

        Catch ex As OperationCanceledException
            Me.Invoke(Sub()
                          ListBox1.Items.Clear()
                          ListBox1.Items.Add("작업이 중단되었습니다.")
                          ListBox1.Items.Add("")
                          DelTempfiles()
                      End Sub)
            Return 0
        End Try

    End Function

    Public Async Function ProcessApple(ct As CancellationToken) As Task(Of Integer)

        Dim filePath As String = Application.StartupPath & "\DB\temp\apple.txt"
        If Not File.Exists(filePath) Then Return 0

        Dim httpClient As New HttpClient()
        Dim outputList As New List(Of String)
        Dim originalLineCount As Integer = File.ReadAllLines(filePath).Length / 2
        Dim processedCount As Integer = 0
        Dim lines As String() = System.IO.File.ReadAllLines(filePath)

        Dim processedLines As New List(Of String)()

        Me.Invoke(Sub()
                      ListBox1.Items.Clear()
                      ListBox1.Items.Add($"{originalLineCount}개의 애플뮤직 정보를 받아오는 중...")
                      ListBox1.Items.Add("")
                  End Sub)

        Try
            For i As Integer = 0 To lines.Length - 1 Step 2
                ct.ThrowIfCancellationRequested()

                Dim metaInfo As String = lines(i)

                If i + 1 >= lines.Length Then
                    Console.WriteLine("데이터가 누락되었습니다: " & metaInfo)
                    Continue For
                End If

                Dim url As String = lines(i + 1)

                Try
                    Dim data As String = Await httpClient.GetStringAsync(url)
                    Dim doc As New HtmlDocument()
                    doc.LoadHtml(data)

                    Dim titleNode As HtmlNode = doc.DocumentNode.SelectSingleNode("//title")

                    If titleNode IsNot Nothing Then
                        Dim title As String = titleNode.InnerText
                        outputList.Add(metaInfo + "/" + title)
                        outputList.Add(url)
                    End If
                Catch ex As Exception

                End Try

                Await Task.Delay(350)

                processedCount += 1

                Me.Invoke(Sub()
                              ListBox1.Items(1) = $"({processedCount}/{originalLineCount})"
                              ListBox1.TopIndex = ListBox1.Items.Count - 1
                          End Sub)
            Next

        Catch ex As OperationCanceledException
            Me.Invoke(Sub()
                          ListBox1.Items.Clear()
                          ListBox1.Items.Add("작업이 중단되었습니다.")
                          ListBox1.Items.Add("")
                      End Sub)
            Return 0
        End Try

        System.IO.File.WriteAllLines(Application.StartupPath & "\DB\temp\apple.txt", outputList)
        Dim diff As Integer = originalLineCount - File.ReadAllLines(filePath).Length / 2
        Return diff

    End Function

    Private Function PreprocessUpdate() As Tuple(Of String, Integer, Integer, Integer)

        Dim openFileDialog As New OpenFileDialog
        openFileDialog.InitialDirectory = Path.Combine(Application.StartupPath, "DB")
        openFileDialog.Title = "DB 파일을 선택해주세요."
        openFileDialog.Filter = "DB 파일 (*.db)|*.db"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim dbPath As String = openFileDialog.FileName
            Using dbConn As New SQLiteConnection($"Data Source={dbPath};Version=3;")
                dbConn.Open()

                Dim baseDir As String = Application.StartupPath & "\DB\temp\"
                File.Copy(baseDir & "apple.txt", baseDir & "apple_update.txt", True)
                File.Copy(baseDir & "files.txt", baseDir & "files_update.txt", True)
                File.Copy(baseDir & "yt.txt", baseDir & "yt_update.txt", True)

                Dim appleDB As New List(Of String)
                Using dbCmd As New SQLiteCommand("SELECT 애플뮤직URL FROM music_data WHERE 유형 = '애플뮤직'", dbConn)
                    Using reader As SQLiteDataReader = dbCmd.ExecuteReader()
                        While reader.Read()
                            appleDB.Add(reader("애플뮤직URL").ToString())
                        End While
                    End Using
                End Using

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

                Dim filesDB As New List(Of String)
                Using dbCmd As New SQLiteCommand("SELECT 제목 FROM music_data WHERE 유형='파일'", dbConn)
                    Using reader As SQLiteDataReader = dbCmd.ExecuteReader()
                        While reader.Read()
                            filesDB.Add(reader("제목").ToString())
                        End While
                    End Using
                End Using

                Dim txtFileLines As List(Of String) = File.ReadAllLines(baseDir & "files_update.txt").ToList()
                For Each title In filesDB
                    txtFileLines.RemoveAll(Function(line) line.Contains($"/{title}/"))
                Next
                File.WriteAllLines(baseDir & "files_update.txt", txtFileLines)

                Dim ytDB As New List(Of String)
                Using dbCmd As New SQLiteCommand("SELECT 유튜브id FROM music_data", dbConn)
                    Using reader As SQLiteDataReader = dbCmd.ExecuteReader()
                        While reader.Read()
                            ytDB.Add(reader("유튜브id").ToString())
                        End While
                    End Using
                End Using

                Dim txtYTLines As List(Of String) = File.ReadAllLines(baseDir & "yt_update.txt").ToList()
                For Each youtubeId In ytDB
                    txtYTLines.RemoveAll(Function(line) line.Contains($"/{youtubeId}/"))
                Next
                File.WriteAllLines(baseDir & "yt_update.txt", txtYTLines)

                dbConn.Close()

                Return New Tuple(Of String, Integer, Integer, Integer)(dbPath, txtFileLines.Count, txtYTLines.Count, txtAppleLines.Count / 2)

            End Using

        ElseIf openFileDialog.ShowDialog() = DialogResult.No Then
            MessageBox.Show("취소되었습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return New Tuple(Of String, Integer, Integer, Integer)(String.Empty, -1, -1, -1)
        End If

    End Function

    Private Function WriteDB(ByVal mode As String, Optional ByVal exDbFilepath As String = "") As String

        Dim baseDir As String = AppDomain.CurrentDomain.BaseDirectory
        Dim fileNames As String() = {"apple.txt", "files.txt", "yt.txt"}
        Dim totalLines As Integer = 0
        Dim processedCount As Integer = 0

        For Each fileName In fileNames
            Dim filePath As String = Path.Combine(baseDir, "DB\temp", fileName)

            If File.Exists(filePath) Then
                Using sr As StreamReader = New StreamReader(filePath)
                    Dim lineCount As Integer = File.ReadLines(filePath).Count()
                    If fileName = "apple.txt" Or fileName = "yt.txt" Then lineCount = lineCount / 2
                    totalLines += lineCount
                End Using
            End If
        Next

        Dim dbFileName As String = "DB_" & DateTime.Now.ToString("yyyyMMdd_HHmm") & ".db"
        Dim dbFilePath As String = Application.StartupPath & "\DB\" & dbFileName

        If mode.ToUpper() = "UPDATE" Then File.Copy(exDbFilepath, dbFilePath)

        Using conn As New SQLiteConnection($"Data Source={dbFilePath}")

            conn.Open()

            Using cmd As New SQLiteCommand(conn)

                cmd.CommandText = "CREATE TABLE IF NOT EXISTS music_data (유형 TEXT, 제목 TEXT, 채널명 TEXT, 최초전송일 TEXT, 전송자 TEXT, 유튜브id VARCHAR, 애플뮤직URL TEXT)"
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

                If File.Exists(Application.StartupPath & "\DB\temp\yt.txt") Then

                    Dim lines As List(Of String) = File.ReadLines(Application.StartupPath & "\DB\temp\yt.txt").ToList()
                    Dim idx As Integer = 0
                    While idx < lines.Count - 1
                        Dim line1 As String = lines(idx)
                        Dim line2 As String = lines(idx + 1)
                        idx += 2
                        Dim pattern As String = "(.*?)\/(.*?)\/(.*?)\/(.*?)\/(.*$)"
                        Dim match As Match = Regex.Match(line1, pattern)

                        If match.Success Then
                            cmd.CommandText = "INSERT INTO music_data (유형, 제목, 채널명, 최초전송일, 전송자, 유튜브id, 애플뮤직URL) VALUES (?, ?, ?, ?, ?, ?, ?)"
                            cmd.Parameters.AddWithValue("유형", match.Groups(1).Value.Trim())
                            cmd.Parameters.AddWithValue("제목", match.Groups(5).Value.Trim())
                            cmd.Parameters.AddWithValue("채널명", line2.Trim())
                            cmd.Parameters.AddWithValue("최초전송일", match.Groups(2).Value.Trim())
                            cmd.Parameters.AddWithValue("전송자", match.Groups(3).Value.Trim())
                            cmd.Parameters.AddWithValue("유튜브id", match.Groups(4).Value.Trim())
                            cmd.Parameters.AddWithValue("애플뮤직URL", Nothing)
                            cmd.ExecuteNonQuery()
                        End If

                        processedCount += 1
                        Me.Invoke(Sub()
                                      ListBox1.Items(1) = $"({processedCount}/{totalLines})"
                                      ListBox1.TopIndex = ListBox1.Items.Count - 1
                                  End Sub)

                    End While
                    cmd.Parameters.Clear()

                End If

                If File.Exists(Application.StartupPath & "\DB\temp\files.txt") Then

                    Dim tempLines As List(Of String) = File.ReadLines(Application.StartupPath & "\DB\temp\files.txt").ToList()
                    For Each line As String In tempLines
                        Dim parts() As String = line.Split("/")

                        cmd.CommandText = "INSERT INTO music_data (유형, 제목, 채널명, 최초전송일, 전송자, 유튜브id, 애플뮤직URL) VALUES (?, ?, ?, ?, ?, ?, ?)"
                        cmd.Parameters.AddWithValue("유형", parts(0))
                        cmd.Parameters.AddWithValue("제목", parts(1))
                        cmd.Parameters.AddWithValue("채널명", Nothing)
                        cmd.Parameters.AddWithValue("최초전송일", parts(2))
                        cmd.Parameters.AddWithValue("전송자", parts(3))
                        cmd.Parameters.AddWithValue("유튜브id", Nothing)
                        cmd.Parameters.AddWithValue("애플뮤직URL", Nothing)
                        cmd.ExecuteNonQuery()

                        processedCount += 1
                        Me.Invoke(Sub()
                                      ListBox1.Items(1) = $"({processedCount}/{totalLines})"
                                      ListBox1.TopIndex = ListBox1.Items.Count - 1
                                  End Sub)

                    Next
                    cmd.Parameters.Clear()

                End If

                If File.Exists(Application.StartupPath & "\DB\temp\apple.txt") Then

                    Dim appleLines As List(Of String) = File.ReadLines(Application.StartupPath & "\DB\temp\apple.txt").ToList()
                    Dim idx2 As Integer = 0
                    While idx2 < appleLines.Count - 1
                        Dim line1 As String = appleLines(idx2)
                        Dim line2 As String = appleLines(idx2 + 1)
                        idx2 += 2

                        Dim pattern As String = "애플\/(.*?)\/(.*?)\/(.*$)"
                        Dim match As Match = Regex.Match(line1, pattern)

                        If match.Success Then
                            Dim title As String = match.Groups(3).Value.Trim()
                            title = title.Replace("‎Apple Music에서 감상하는 ", String.Empty)
                            title = title.Replace("‎Apple Music에서 만나는 ", String.Empty)
                            title = title.Replace("‎Apple Music에서 만나는 ", String.Empty)
                            title = title.Replace("‎Apple Music의 ", String.Empty)
                            title = title.Replace("&amp;", "&")

                            cmd.CommandText = "INSERT INTO music_data (유형, 제목, 채널명, 최초전송일, 전송자, 유튜브id, 애플뮤직URL) VALUES (?, ?, ?, ?, ?, ?, ?)"
                            cmd.Parameters.AddWithValue("유형", "애플뮤직")
                            cmd.Parameters.AddWithValue("제목", title)
                            cmd.Parameters.AddWithValue("채널명", Nothing)
                            cmd.Parameters.AddWithValue("최초전송일", match.Groups(1).Value.Trim())
                            cmd.Parameters.AddWithValue("전송자", match.Groups(2).Value.Trim())
                            cmd.Parameters.AddWithValue("유튜브id", Nothing)
                            cmd.Parameters.AddWithValue("애플뮤직URL", line2.Trim())
                            cmd.ExecuteNonQuery()
                        End If

                        processedCount += 1
                        Me.Invoke(Sub()
                                      ListBox1.Items(1) = $"({processedCount}/{totalLines})"
                                      ListBox1.TopIndex = ListBox1.Items.Count - 1
                                  End Sub)

                    End While
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
            If String.IsNullOrWhiteSpace(TxtPassphrase.Text) Then
                keyFile = New PrivateKeyFile(keyfilePath)
            Else
                Dim passphrase As String = TxtPassphrase.Text
                keyFile = New PrivateKeyFile(keyfilePath, passphrase)
            End If
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
                If String.IsNullOrWhiteSpace(TxtPassphrase.Text) Then
                    keyFile = New PrivateKeyFile(keyfilePath)
                Else
                    Dim passphrase As String = TxtPassphrase.Text
                    keyFile = New PrivateKeyFile(keyfilePath, passphrase)
                End If

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
                If String.IsNullOrWhiteSpace(TxtPassphrase.Text) Then
                    keyFile = New PrivateKeyFile(keyfilePath)
                Else
                    Dim passphrase As String = TxtPassphrase.Text
                    keyFile = New PrivateKeyFile(keyfilePath, passphrase)
                End If

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

    Private Sub BtnLoadTXT_Click(sender As Object, e As EventArgs) Handles BtnLoadTXT.Click

        Dim lines As List(Of String) = PreprocessTextFiles()
        If lines.Count > 0 Then
            ListBox1.Items.Clear()
            ListBox1.Items.Add("대화 파일 불러오기 성공")
            Dim countTuple As Tuple(Of Integer, Integer) = RemoveAndSplit(lines)
            Dim fileCount As Integer = countTuple.Item1
            Dim youtubeCount As Integer = countTuple.Item2
            Dim appleCount As Integer = RemoveApple()
            ListBox1.Items.Add("")
            ListBox1.Items.Add($"파일 : {fileCount}개")
            ListBox1.Items.Add($"유튜브 : {youtubeCount}개")
            ListBox1.Items.Add($"애플뮤직 : {appleCount}개")
            ListBox1.Items.Add("")
            ListBox1.Items.Add("(중복 제외)")
            BtnMakeDB.Enabled = True
            BtnUpdateDB.Enabled = True
        Else
            MessageBox.Show("파일을 선택하지 않았거나 잘못된 파일입니다." & vbCrLf & "다시 시도해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            ListBox1.Items.Clear()
            ListBox1.Items.Add("대화 파일 불러오기 실패")
            BtnUpdateDB.Enabled = False
            BtnMakeDB.Enabled = False
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
                        Dim YTDiff As Integer = Await Task.Run(Function() ProcessYT(cts.Token), cts.Token)
                        Dim AppleDiff As Integer = Await Task.Run(Function() ProcessApple(cts.Token), cts.Token)

                        If Not cts.IsCancellationRequested Then
                            BtnMakeDB.Enabled = False

                            Dim dbFileName As String = Await Task.Run(Function() WriteDB("CREATE"))
                            If ChkBaroUpload.Checked = True Then Await Task.Run(Sub() UploadDB(Application.StartupPath & "\DB\" & dbFileName))

                            ListBox1.Items.Clear()
                            If ChkBaroUpload.Checked = True Then ListBox1.Items.Add("DB 생성 및 업로드를 완료하였습니다.") Else ListBox1.Items.Add("DB 생성을 완료하였습니다.")
                            ListBox1.Items.Add("")
                            ListBox1.Items.Add($"삭제 / 비공개 영상 : {YTDiff}개")
                            ListBox1.Items.Add($"삭제 / 지역제한 애플뮤직 : {AppleDiff}개")
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
                            If File.Exists(baseDir & fname & ".txt") Then File.Delete(baseDir & fname & ".txt")
                            File.Move(baseDir & fname & "_update.txt", baseDir & fname & ".txt")
                        Next

                        Dim YTDiff As Integer = Await Task.Run(Function() ProcessYT(cts.Token), cts.Token)
                        Dim AppleDiff As Integer = Await Task.Run(Function() ProcessApple(cts.Token), cts.Token)

                        If Not cts.IsCancellationRequested Then
                            BtnUpdateDB.Enabled = False

                            Dim dbFileName As String = Await Task.Run(Function() WriteDB("UPDATE", dbInfo.Item1))
                            If ChkBaroUpload.Checked = True Then Await Task.Run(Sub() UploadDB(Application.StartupPath & "\DB\" & dbFileName))

                            ListBox1.Items.Clear()
                            If ChkBaroUpload.Checked = True Then ListBox1.Items.Add("DB 업데이트 및 업로드를 완료하였습니다.") Else ListBox1.Items.Add("DB 업데이트를 완료하였습니다.")
                            ListBox1.Items.Add("")
                            ListBox1.Items.Add($"업데이트 한 파일 : {dbInfo.Item2}개")
                            ListBox1.Items.Add($"업데이트 한 유튜브 : {dbInfo.Item3 - YTDiff}개")
                            ListBox1.Items.Add($"업데이트 한 애플뮤직 : {dbInfo.Item4 - AppleDiff}개")
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
            TxtKeyPath.Text = openFileDialog.FileName '
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