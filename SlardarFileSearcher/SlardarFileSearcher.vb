Imports System.Globalization
Imports System.IO
Imports System.Windows.Forms
Module SlardarFileSearcher

    Private LanguageCode As String = "en_AU"
    Private DefaultColor As ConsoleColor = Console.ForegroundColor
    Sub Main()
        Dim LanguageName As String = CultureInfo.CurrentCulture.Name.Replace("-", "").Replace("_", "")
        If LanguageName.ToLower = "zhcn" Then
            LanguageCode = "zh_CN"
        Else
            LanguageCode = "en_AU"
        End If
        Console.Title = "Slardar File Searcher"
        Console.ForegroundColor = ConsoleColor.Green
        Dim FolderBrowserDialog As New FolderBrowserDialog
        FolderBrowserDialog.Description = GetString("FolderBrowserDialogDescription")
        FolderBrowserDialog.RootFolder = Environment.SpecialFolder.MyComputer
        FolderBrowserDialog.SelectedPath = Environment.CurrentDirectory
        FolderBrowserDialog.ShowNewFolderButton = False
        If FolderBrowserDialog.ShowDialog = DialogResult.OK Then
            Dim RootDir As String = FolderBrowserDialog.SelectedPath
            If Not Directory.Exists(RootDir) Then
                Console.ForegroundColor = ConsoleColor.Red
                Console.WriteLine(GetString("FolderNotExits"))
                ExitApp()
            Else
                Dim SearchingText As String = GetSearchText()
                Dim MaxFileSize As Long = GetMaxFileSize()
                Console.WriteLine(GetString("ExtensionSearch"))
                Dim ExtensionSearch As List(Of String) = New List(Of String)
                For Each ExtensionName As String In Console.ReadLine().Trim().ToLower().Split(",").ToList()
                    If ExtensionName <> "*" And ExtensionName.Length > 0 Then
                        ExtensionSearch.Add("." + ExtensionName)
                    End If
                Next
                Console.WriteLine(GetString("ExtensionIgnore"))
                Dim ExtensionIgnore As List(Of String) = New List(Of String)
                For Each ExtensionName As String In Console.ReadLine().Trim().ToLower().Split(",").ToList()
                    If ExtensionName <> "*" And ExtensionName.Length > 0 Then
                        ExtensionIgnore.Add("." + ExtensionName)
                    End If
                Next
                Console.WriteLine(GetString("FolderIgnore"))
                Dim FolderIgnore As List(Of String) = New List(Of String)
                For Each FolderName As String In Console.ReadLine().Trim().ToLower().Split(",").ToList()
                    If FolderName <> "*" And FolderName.Length > 0 Then
                        FolderIgnore.Add(FolderName)
                    End If
                Next
                Dim FoundFileList As List(Of String) = SearchFolder(RootDir, SearchingText, MaxFileSize, ExtensionSearch, ExtensionIgnore, FolderIgnore)
                Console.ForegroundColor = ConsoleColor.Red
                Try
                    Dim OutputIndex As Integer = 0
                    While File.Exists(Environment.CurrentDirectory + "\output" + OutputIndex.ToString() + ".txt")
                        OutputIndex += 1
                    End While
                    Dim OFS As StreamWriter = New StreamWriter(New FileStream(Environment.CurrentDirectory + "\output" + OutputIndex.ToString() + ".txt", FileMode.Create))
                    Console.WriteLine("=====================================================")
                    Console.WriteLine(GetString("FindInFileList") + FoundFileList.Count.ToString())
                    OFS.WriteLine(GetString("SearhingFolder") + RootDir)
                    OFS.WriteLine(GetString("SearchingText") + SearchingText)
                    OFS.WriteLine(GetString("MaxFileSizeText") + MaxFileSize.ToString() + "B")
                    OFS.WriteLine(GetString("ExtensionSearchText"))
                    OFS.WriteLine(GetString("FindInFileList") + FoundFileList.Count.ToString())
                    For Each Extension As String In ExtensionSearch
                        OFS.WriteLine(vbTab + Extension)
                    Next
                    OFS.WriteLine(GetString("ExtensionIgnoreText"))
                    For Each Extension As String In ExtensionIgnore
                        OFS.WriteLine(vbTab + Extension)
                    Next
                    OFS.WriteLine(GetString("FolderIgnoreText"))
                    For Each FolderName As String In FolderIgnore
                        OFS.WriteLine(vbTab + FolderName)
                    Next
                    OFS.WriteLine("=====================================================")
                    OFS.WriteLine()
                    For Each FilePath As String In FoundFileList
                        Console.WriteLine(vbTab + GetString("FindInFile") + FilePath)
                        OFS.WriteLine(vbTab + GetString("FindInFile") + FilePath)
                    Next
                    Console.WriteLine(GetString("Total") + FoundFileList.Count().ToString())
                    OFS.WriteLine(GetString("Total") + FoundFileList.Count().ToString())
                    OFS.Close()
                    Console.WriteLine(GetString("ResultWriteTo") + Environment.CurrentDirectory + "\output" + OutputIndex.ToString() + ".txt")
                Catch ex As Exception
                    Console.ForegroundColor = ConsoleColor.Yellow
                    Console.WriteLine(GetString("OutputError") + vbCrLf + ex.StackTrace)
                End Try
                ExitApp()
            End If
        Else
            ExitApp()
        End If
    End Sub

    Sub ExitApp()
        Console.WriteLine(GetString("ExitText"))
        Console.ReadKey()
        Console.ForegroundColor = DefaultColor
        End
    End Sub
    Function GetMaxFileSize() As Long
        Dim MaxSizeLong As Long = -1
        While MaxSizeLong < 0
            Console.WriteLine(GetString("MaxFileSize"))
            Dim MaxSize As String = Console.ReadLine().ToUpper()
            If MaxSize.Length > 0 Then
                MaxSize = MaxSize.Replace(" ", "").Trim()
                If MaxSize.Last = "B" Then
                    MaxSize = MaxSize.Remove(MaxSize.Length - 1)
                End If
                If Not Long.TryParse(MaxSize, MaxSizeLong) Then
                    Select Case MaxSize.Last
                        Case "K"
                            If Not TextSizeToLong(MaxSize, "K", MaxSizeLong) Then
                                Console.ForegroundColor = ConsoleColor.Red
                                Console.WriteLine(GetString("IllegalSizeText"))
                                Console.ForegroundColor = ConsoleColor.Green
                            End If
                        Case "M"
                            If Not TextSizeToLong(MaxSize, "M", MaxSizeLong) Then
                                Console.ForegroundColor = ConsoleColor.Red
                                Console.WriteLine(GetString("IllegalSizeText"))
                                Console.ForegroundColor = ConsoleColor.Green
                            End If
                        Case "G"
                            If Not TextSizeToLong(MaxSize, "G", MaxSizeLong) Then
                                Console.ForegroundColor = ConsoleColor.Red
                                Console.WriteLine(GetString("IllegalSizeText"))
                                Console.ForegroundColor = ConsoleColor.Green
                            End If
                        Case "T"
                            If Not TextSizeToLong(MaxSize, "T", MaxSizeLong) Then
                                Console.ForegroundColor = ConsoleColor.Red
                                Console.WriteLine(GetString("IllegalSizeText"))
                                Console.ForegroundColor = ConsoleColor.Green
                            End If
                        Case Else
                            Console.ForegroundColor = ConsoleColor.Red
                            Console.WriteLine(GetString("IllegalSizeText"))
                            Console.ForegroundColor = ConsoleColor.Green
                            MaxSizeLong = -1
                    End Select
                End If
            End If
        End While
        Return MaxSizeLong
    End Function
    Function TextSizeToLong(ByVal Text As String, ByVal Unit As Char, ByRef SizeLong As Long) As Boolean
        Text = Text.Remove(Text.Length - 1)
        Dim SizeLongTemp As Long = Long.MinValue
        If Long.TryParse(Text, SizeLongTemp) Then
            Select Case Unit
                Case "K"
                    SizeLong = SizeLongTemp * 1024
                    Return True
                Case "M"
                    SizeLong = SizeLongTemp * 1024 * 1024
                    Return True
                Case "G"
                    SizeLong = SizeLongTemp * 1024 * 1024 * 1024
                    Return True
                Case "T"
                    SizeLong = SizeLongTemp * 1024 * 1024 * 1024 * 1024
                    Return True
                Case Else
                    Return False
            End Select
        Else
            Return False
        End If
    End Function

    Function GetSearchText() As String
        Console.WriteLine(GetString("SearchText"))
        Dim SearchingText As String = ""
        Dim CurrentLine As String = Console.ReadLine()
        While CurrentLine IsNot Nothing
            If CurrentLine.Length > 0 Then
                SearchingText += CurrentLine
            End If
            CurrentLine = Console.ReadLine
            If CurrentLine IsNot Nothing Then
                SearchingText += vbCrLf
            End If
        End While
        If SearchingText.Length = 0 Then
            Return GetSearchText()
        Else
            Return SearchingText
        End If
    End Function

    Function SearchFolder(ByVal SearchingPath As String, ByVal SearchingText As String, ByVal MaxFileSize As Long, ByVal ExtensionSearch As List(Of String), ByVal ExtensionIgnore As List(Of String), ByVal FolderIgnore As List(Of String)) As List(Of String)
        Dim FilePathList As New List(Of String)
        Console.WriteLine(GetString("ReadingFolder") + SearchingPath)
        Try
            For Each FilePath As String In Directory.GetFiles(SearchingPath)
                Try
                    Dim FileInfo As New FileInfo(FilePath)
                    If Not ExtensionIgnore.Contains(FileInfo.Extension) And (ExtensionSearch.Count = 0 Or ExtensionSearch.Contains(FileInfo.Extension.ToLower())) And (MaxFileSize = 0 Or (FileInfo.Length <= MaxFileSize)) Then
                        Console.WriteLine(vbTab + GetString("ReadingFile") + FilePath)
                        If File.ReadAllText(FilePath).Contains(SearchingText) Then
                            Console.ForegroundColor = ConsoleColor.Red
                            Console.WriteLine(vbTab + GetString("FindInFile") + FilePath)
                            Console.ForegroundColor = ConsoleColor.Green
                            FilePathList.Add(FilePath)
                        End If
                    Else
                    End If
                Catch ex As Exception
                    Console.ForegroundColor = ConsoleColor.Yellow
                    Console.WriteLine(vbTab + GetString("ReadFileFailed") + FilePath)
                    Console.ForegroundColor = ConsoleColor.Green
                End Try
            Next

            For Each SubFolder As String In Directory.GetDirectories(SearchingPath)
                If Not FolderIgnore.Contains(Path.GetFileName(SubFolder).ToLower()) Then
                    FilePathList.AddRange(SearchFolder(SubFolder, SearchingText, MaxFileSize, ExtensionSearch, ExtensionIgnore, FolderIgnore))
                End If
            Next
        Catch ex As Exception
            Console.ForegroundColor = ConsoleColor.Yellow
            Console.WriteLine(GetString("ReadFolderFailed") + SearchingPath)
            Console.ForegroundColor = ConsoleColor.Green
        End Try
        Return FilePathList
    End Function


    Function GetString(ByVal Name As String) As String
        If LanguageCode = "zh_CN" Then
            Return My.Resources.Resource_zh_CN.ResourceManager.GetString(Name)
        Else
            Return My.Resources.Resource_en_AU.ResourceManager.GetString(Name)
        End If
    End Function
End Module
