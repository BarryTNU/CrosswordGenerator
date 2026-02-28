
Imports System.IO
Imports System.Drawing.Printing
Imports System.Net

Public Class DictionaryEditor
    Public DefaultPath As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + Path.DirectorySeparatorChar + "Crosswords"
    Public xWordFilePath As String = Path.Combine(DefaultPath, "CrossList.csv")
    Public pWordFilePath As String = Path.Combine(DefaultPath, "PhraseList.csv")
    Public cWordFilePath = Path.Combine(DefaultPath, "CodeList.csv")
    Public DictFilePath As String

#Region "DATA STRUCTURES"
    Public Class Clue
        Public Word As String
        Public Row As Integer
        Public Col As Integer
        Public IsAcross As Boolean
        Public Clue As String
        Public ClueNumber As Integer
        Public PlaceLetter As Boolean = False ' Used in crossword puzzles to indicate whether the letter should be revealed on the grid (for codeword puzzles, all letters are hidden)
    End Class
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub


    Private Word As String
    Private UniqueWord As Boolean = False
    Private PrintPageIndex As Integer = 0
    Private WordList As New List(Of Clue)
    Private currentWord As String = ""
    Private CurrentRow As Integer = 0
    Private CurrentCol As Integer = 0
    Private rnd As New Random()
    Private WithEvents pd As New PrintDocument
    Private WithEvents rb_xword As New RadioButton
    Private WithEvents rb_cword As New RadioButton
    Private WithEvents rb_pword As New RadioButton
    ' Private WithEvents btn_PuzzleType As New Button
    Private WithEvents btn_NewList As New Button
    Private WithEvents btn_Restore As New Button
    Private WordLength As Integer = 12
    ' ===================== UI =====================
    Private lbl_Dictionary As New Label
    Private lbl_New As New Label

    '====================== LISTBOXES =====================
    Private WithEvents lv_Dictionary As New ListView
    Private txt_NewWords As New TextBox
    Private txt_NewClues As New TextBox
    Private Puzzle As String

#End Region

#Region "FORM LOAD"
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If My.Application.CommandLineArgs.Count > 0 Then
            Puzzle = My.Application.CommandLineArgs(0)
        End If

        Me.Text = "VB.NET Dictionary Generator"
        Me.Size = New Size(610, 880)
        Me.Location = New Point(300, 50)

        If Puzzle = "" Then
            Puzzle = "cWord" ' Default to codeword if not set
        End If

        If Puzzle = "pWord" Then
            DictFilePath = pWordFilePath
        ElseIf Puzzle = "xWord" Then
            DictFilePath = xWordFilePath
        ElseIf Puzzle = "cWord" Then
            DictFilePath = cWordFilePath
        End If

        ' Create a folder for the word list if it doesn't exist
        If Not Directory.Exists(DefaultPath) Then
            Directory.CreateDirectory(DefaultPath)
        End If

        '  'Check if there is a dictionary. If not, download a new one from Camsoft.au. If the user declines to download a new dictionary, prompt them to load a dictionary to use the application.
        If Not File.Exists(DictFilePath) Then
            Dim result = MessageBox.Show("No dictionary found. Do you want to download a new dictionary from Camsoft.au?", "Dictionary Not Found", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                DownloadDictionary(Puzzle) ' Get a new Dictionary List from Camsoft.au
            Else
                MessageBox.Show("No dictionary loaded. Please load a dictionary to use the application.")
                Exit Sub
            End If
        End If

        SetupUI()
        LoadFile(Nothing, Nothing) ' Load the dictionary into the list

#Region "MERGE PHRASE FILES"
        ' Only needed if you want to update the phrase list with new phrases, and you have multiple batch files to merge. You can comment this out after running it once, and it will create a merged phrase list that will be used going forward.

        ' Dim outputFile As String = DefaultPath & "\PhraseList.csv"
        ' Dim inputFolder As String = DefaultPath
        ' MergePhraseFiles(inputFolder, outputFile)
#End Region

    End Sub

    Private Sub SetupUI()

        '====================TEXT BOXES FOR NEW ENTRIES===================

        Try
            With txt_NewWords
                .Font = New Font("Segoe UI", 12, FontStyle.Bold)
                .CharacterCasing = CharacterCasing.Upper
                .Location = New Point(50, 70)
                .Width = 130
                .Text = "New words."
                .Select()
                .Focus()
            End With

            With txt_NewClues
                .Font = New Font("Segoe UI", 12, FontStyle.Bold)
                .Location = New Point(180, 70)
                .Width = 230
                .Text = "New clues."
            End With

            '===================== DICTIONARY LABLE=====================

            With lbl_Dictionary
                .Font = New Font("Segoe UI", 14, FontStyle.Bold)
                .Text = "Dictionary"
                .ForeColor = Color.Black
                .BackColor = Color.White
                .BorderStyle = BorderStyle.FixedSingle
                .TextAlign = ContentAlignment.MiddleCenter
                .Size = New Size(200, 40)
                .AutoSize = True
                .Location = New Point(120, 20)
            End With

            '===================== DICTIONARY LIST BOX  =====================
            With lv_Dictionary
                .Font = New Font("Segoe UI", 12, FontStyle.Bold)
                .Location = New Point(50, 100)
                .Size = New Size(500, 650)
                .View = View.Details
                .FullRowSelect = True
                .GridLines = True
                .Columns.Clear()
                .Columns.Add("Word", 200, HorizontalAlignment.Left)
                .Columns.Add("Clue", 275, HorizontalAlignment.Left)
            End With

            '===================== POSITION CONTROLS =====================  
            txt_NewWords.Location = New Point(50, 70)
            txt_NewWords.Width = 200
            txt_NewWords.Text = "New words."
            txt_NewWords.Height = 50
            txt_NewClues.Height = txt_NewWords.Height
            txt_NewClues.Location = New Point(260, 70)
            txt_NewClues.Width = 290
            txt_NewClues.Height = txt_NewWords.Height
            lbl_Dictionary.Location = New Point(160, 20)

            '===================== RADIO BUTTONS =====================

            Dim rb_xword As New RadioButton With {
                 .Font = New Font("Segoe UI", 12, FontStyle.Bold),
            .Text = "Crossword",
                .Location = New Point(280, 760),
                .AutoSize = True
            }
            Dim rb_cword As New RadioButton With {
                .Font = New Font("Segoe UI", 12, FontStyle.Bold),
            .Text = "Codeword",
                .Location = New Point(280, 780),
                .AutoSize = True
            }
            Dim rb_pword As New RadioButton With {
                .Font = New Font("Segoe UI", 12, FontStyle.Bold),
            .Text = "PhraseWord",
                .Location = New Point(280, 800),
                .AutoSize = True
            }
            '===================== BUTTONS =====================

            btn_NewList.Font = New Font("Segoe UI", 12, FontStyle.Bold)
            btn_NewList.Size = New Size(200, 30)
            btn_NewList.Text = "Load New Dictionary"
            btn_NewList.Location = New Point(50, 760)


            btn_Restore.Font = New Font("Segoe UI", 12, FontStyle.Bold)
            btn_Restore.Size = New Size(200, 30)
            btn_Restore.Text = "Restore Dictionary"
            btn_Restore.Location = New Point(50, 790)

            '===================== ADD CONTROLS AND EVENT HANDLERS =====================

            Controls.AddRange({lbl_Dictionary, lv_Dictionary, txt_NewWords, txt_NewClues, btn_NewList, btn_Restore, rb_xword, rb_cword, rb_pword})

            AddHandler txt_NewWords.KeyDown, AddressOf Shared_KeyDown
            AddHandler txt_NewClues.KeyDown, AddressOf Shared_KeyDown
            AddHandler txt_NewWords.GotFocus, AddressOf txtBox_HasFocus
            AddHandler txt_NewClues.GotFocus, AddressOf txtBox_HasFocus
            AddHandler rb_xword.CheckedChanged, AddressOf RadioButton_CheckedChanged
            AddHandler rb_cword.CheckedChanged, AddressOf RadioButton_CheckedChanged
            AddHandler rb_pword.CheckedChanged, AddressOf RadioButton_CheckedChanged
            AddHandler btn_NewList.Click, AddressOf btn_NewList_Click
            AddHandler btn_Restore.Click, AddressOf btn_Restore_Click
            AddHandler lv_Dictionary.KeyDown, AddressOf lv_Dictionary_KeyDown
            AddHandler lv_Dictionary.MouseDoubleClick, AddressOf lv_Dictionary_MouseDoubleClick

        Catch ex As Exception
            MessageBox.Show(ex.Message & "Error setting up the UI.")
        End Try

    End Sub

#End Region
#Region "DOWNLOAD PHRASE LIST FROM CAMSOFT.AU"

    Private Sub DownloadDictionary(pStyle As String)
        Dim fPath As String = DictFilePath
        Dim SourceFile As String = ""
        Dim WordLength As Integer = 0
        Dim ipAddress As String = ""

        If pStyle = "pWord" Then
            SourceFile = "pWord.csv"
            WordLength = 15
        ElseIf pStyle = "xWord" Then
            SourceFile = "xWord.csv"
            WordLength = 12
        ElseIf pStyle = "cWord" Then
            SourceFile = "cWord.csv"
            WordLength = 12
        End If

        ipAddress = "https://camsoft.au/cwg/" & SourceFile

        Try
            Using client As New WebClient()
                client.DownloadFile(ipAddress, fPath)
            End Using

            Dim DictList As New List(Of Clue)

            Using sr As New StreamReader(fPath)
                While Not sr.EndOfStream
                    Dim line As String = sr.ReadLine()
                    Dim parts As String() = line.Split(","c)
                    If parts.Length >= 2 Then
                        Dim entry As New Clue With {
                            .Word = parts(0).Trim(),
                            .Clue = parts(1).Trim()
                        }
                        If Len(entry.Word) <= WordLength AndAlso Len(entry.Word) > 4 Then 'Select only phrases of appropriate length for the puzzle
                            DictList.Add(entry)
                        End If
                    End If
                End While
            End Using

            Dim selected1000 As List(Of Clue) = DictList.Take(1000).ToList() 'If the list is larger than 1000, take a random sample of 1000 phrases to keep the dictionary manageable. You can adjust this number as needed.    
            CopyWithAutoName(fPath) ' Save a copy of the downloaded phrase list with an auto-generated name to avoid overwriting the existing one.
            SaveDictionary(fPath, selected1000)
            DictFilePath = fPath
            MsgBox("New Dictionary installed")

        Catch ex As Exception
            MsgBox("Failed to Load Phrase List from Camsoft.au")
        End Try
    End Sub

#End Region


#Region "Load and Save DICTIONARY"

    Private Async Sub LoadFile(sender As Object, e As EventArgs)
        '
        Dim loadingForm As New FrmLoading()

        Cursor = Cursors.WaitCursor
        lv_Dictionary.Items.Clear()
        WordList.Clear()

        Try
            With loadingForm
                .StartPosition = FormStartPosition.CenterParent
            End With

            loadingForm.Show()

            Dim progress = New Progress(Of Integer)(Sub(p) loadingForm.UpdateProgress(p))
            Dim items = Await Task.Run(Function() PopulateListBox(DictFilePath, progress))
            Dim result = Await Task.Run(Function() PopulateListBox(DictFilePath, progress))
            WordList = result.wlist
            lv_Dictionary.BeginUpdate()
            lv_Dictionary.Items.AddRange(result.items.ToArray())
            lv_Dictionary.EndUpdate()
            loadingForm.Close()
            Dim nbrClues As Integer = lv_Dictionary.Items.Count + 1
            lbl_Dictionary.Text = nbrClues & " Word Dictionary"

            Cursor = Cursors.Default
        Catch ex As Exception
            MessageBox.Show(ex.Message & "Failed to load Dictionary.")
            loadingForm.Close()
            Cursor = Cursors.Default
        End Try

    End Sub

#End Region

#Region "Populate ListBox"
    Private Function PopulateListBox(path As String, progress As IProgress(Of Integer)) As (items As List(Of ListViewItem), wlist As List(Of Clue))
        Dim items As New List(Of ListViewItem)
        Dim wlist As New List(Of Clue)
        Try
            Using fs As New FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Using sr As New StreamReader(fs)
                    Dim totalBytes = sr.BaseStream.Length
                    While Not sr.EndOfStream
                        Dim line = sr.ReadLine()
                        Dim p = line.Split(","c, 2)
                        Dim w = If(p.Length > 1, p(0).Trim(), "")
                        Dim c = If(p.Length > 1, p(1).Trim(), "")
                        If w = "Word" AndAlso c = "Clue" Then Continue While
                        Dim item As New ListViewItem(w)
                        If Not String.IsNullOrEmpty(c) Then
                            c = Char.ToUpper(c(0)) & c.Substring(1)
                        End If
                        item.SubItems.Add(c)
                        items.Add(item)
                        wlist.Add(New Clue With {.Word = w, .Clue = c})
                        If totalBytes > 0 Then
                            Dim percent = CInt((sr.BaseStream.Position / totalBytes) * 100)
                            percent = Math.Max(0, Math.Min(percent, 100))
                            progress.Report(percent)
                        End If
                    End While
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Failed to load Dictionary.")
        End Try

        Return (items, wlist)
    End Function

    Private Sub SaveDictionary(Wordfilepath As String, wlist As List(Of Clue))
        Try
            Using writer As New StreamWriter(Wordfilepath, False)

                Dim Clu As String
                For Each entry In wlist
                    Dim Cr = LCase(entry.Clue)
                    Cr = Char.ToUpper(Cr(0)) & Cr.Substring(1)
                    writer.WriteLine($"{entry.Word},{Cr}")
                Next
                writer.Close()
            End Using

            Dim nbrClues As Integer = lv_Dictionary.Items.Count + 1
            lbl_Dictionary.Text = nbrClues & " Word Dictionary"

        Catch ex As Exception
            MessageBox.Show("Failed to save Dictionary.")
        End Try

    End Sub
    Private Function CopyWithAutoName(sourceFile As String) As String
        Dim folder As String = Path.GetDirectoryName(sourceFile)
        Dim baseName As String = Path.GetFileNameWithoutExtension(sourceFile)
        Dim ext As String = Path.GetExtension(sourceFile)

        Dim newFile As String = Path.Combine(folder, baseName & ext)
        Dim counter As Integer = 1

        While File.Exists(newFile)
            newFile = Path.Combine(folder, $"{baseName} ({counter}){ext}")
            counter += 1
        End While

        File.Copy(sourceFile, newFile)
        Return newFile
    End Function

#End Region


#Region "ADD WORDS TO DICTIONARY"

    Private Sub lv_Dictionary_MouseDoubleClick(sender As Object, e As MouseEventArgs) _
        ' Handles lv_Dictionary.MouseDoubleClick

        Dim hit = lv_Dictionary.HitTest(e.Location)
        Dim Clue As String
        Dim Word As String
        If hit.Item Is Nothing Then Exit Sub
        Try

            '===============If double click in Word column then Edit word and clue================'
            If hit.SubItem Is hit.Item.SubItems(0) Then
                Word = hit.Item.Text
                Clue = hit.Item.SubItems(1).Text
                txt_NewWords.Text = Word
                txt_NewClues.Text = Clue
                lv_Dictionary.Items.Remove(hit.Item)
                WordList.RemoveAll(Function(c) c.Word = Word AndAlso c.Clue = Clue)
                txt_NewWords.Focus()

                '=========== If double-clicked on the clue column, Edit  the clue===========
            ElseIf hit.SubItem Is hit.Item.SubItems(1) Then
                Word = hit.Item.Text
                Clue = hit.SubItem.Text
                txt_NewWords.Text = Word
                txt_NewClues.Text = Clue
                lv_Dictionary.Items.Remove(hit.Item)
                WordList.RemoveAll(Function(c) c.Word = Word AndAlso c.Clue = Clue)
                txt_NewClues.Focus()
            End If

            SaveDictionary(DictFilePath, WordList) 'Save the updated dictionary after removing the old entry .When the edited entry is saved, it will be added back to the dictionary with the updated word or clue.

        Catch ex As Exception
            MessageBox.Show("Error processing the selected item.")
        End Try

    End Sub

    Private Function CheckForDuplicates(Word As String)

        Try
            Word = Word.Trim().ToUpper()

            If Word = "" Then Return False

            If Not WordList.Any(Function(x) x.Word.Equals(Word, StringComparison.OrdinalIgnoreCase)) Then
                Return True
            Else
                MessageBox.Show("Duplicate word.")
                Return False
            End If

        Catch ex As Exception
            MessageBox.Show("Invalid word.")
            Exit Function
        End Try
        Return True

    End Function

    '==========DETECT KEYPRESS EVENTS =====================
    Private Sub lv_Dictionary_KeyDown(sender As Object, e As KeyEventArgs)
        'Delete the selected word and clue when the delete key is pressed

        If e.KeyCode = Keys.Delete Then
            Dim Clu As String
            Dim Question As String = "Delete " & lv_Dictionary.SelectedItems(0).Text & "?"
            If MessageBox.Show("Are you sure?", Question, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                lv_Dictionary.Items.Remove(lv_Dictionary.SelectedItems(0))
                WordList.RemoveAll(Function(c) c.Word = Word AndAlso c.Clue = Clu)
            End If
            e.Handled = True
            SaveDictionary(DictFilePath, WordList)
        End If

    End Sub

    Private Sub Shared_KeyDown(sender As Object, e As KeyEventArgs)

        Dim tb = DirectCast(sender, TextBox)
        Try
            If e.KeyCode = Keys.Enter OrElse e.KeyCode = Keys.Tab Then
                e.SuppressKeyPress = True
                txt_NewWords.Text.Trim().ToUpper()

                If Puzzle = "pWord" Then
                    WordLength = 15
                ElseIf Puzzle = "xWord" Then
                    WordLength = 12
                ElseIf Puzzle = "cWord" Then
                    WordLength = 12
                End If

                If tb Is txt_NewWords Then
                    Word = txt_NewWords.Text.Trim().ToUpper()
                    Dim Clue As String = txt_NewClues.Text.Trim()
                    If Len(Word) < 2 OrElse Len(Word) > 15 Then
                        MessageBox.Show("Word must be 2–15 letters (A–Z only).")
                        txt_NewWords.Clear()
                        txt_NewWords.Focus()
                        Exit Sub
                    End If
                    txt_NewClues.Focus()

                ElseIf tb Is txt_NewClues Then
                    Dim Clue As String = txt_NewClues.Text.Trim()

                    Word = txt_NewWords.Text.Trim().ToUpper()
                    If Len(Clue) < 5 OrElse Len(Clue) > 30 Then
                        MessageBox.Show("Clue must be 5-30 characters.")
                        txt_NewClues.Clear()
                        txt_NewClues.Focus()
                        Exit Sub
                    End If

                    UniqueWord = CheckForDuplicates(Word)
                    If UniqueWord Then
                        Dim item As New ListViewItem(Word)      ' Word Column 
                        item.SubItems.Add(Clue)                 ' Clue Column 
                        lv_Dictionary.Items.Insert(0, item) 'Add the word and clue to the first row in the list
                        WordList.Add(New Clue With {.Word = Word, .Clue = Clue})
                        txt_NewWords.Clear()
                        txt_NewClues.Clear()
                        txt_NewWords.Focus()
                    Else
                        txt_NewWords.Clear()
                        txt_NewClues.Clear()
                        txt_NewWords.Focus()
                    End If
                End If
                SaveDictionary(DictFilePath, WordList)

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message & " Error processing input.")
        End Try

    End Sub

    Private Sub txtBox_HasFocus(sender As Object, e As EventArgs)

        Dim tb = DirectCast(sender, TextBox)
        If tb Is txt_NewWords Then
            ' txt_NewWords.Clear()
        ElseIf tb Is txt_NewClues Then
            'txt_NewClues.Clear()
        End If
    End Sub

#End Region
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
    End Sub

#Region "RADIO BUTTON CHECKED CHANGED"

    Private Sub RadioButton_CheckedChanged(sender As Object, e As EventArgs)
        Dim rb = DirectCast(sender, RadioButton)
        If rb.Checked Then

            lv_Dictionary.Items.Clear()

            Select Case rb.Text
                Case "Crossword"
                    Puzzle = "xWord"
                    DictFilePath = xWordFilePath
                    LoadFile(Nothing, Nothing)
                    lv_Dictionary.Show()
                Case "Codeword"
                    Puzzle = "cWord"
                    DictFilePath = cWordFilePath
                    LoadFile(Nothing, Nothing)
                    lv_Dictionary.Show()
                Case "PhraseWord"
                    Puzzle = "pWord"
                    DictFilePath = pWordFilePath
                    LoadFile(Nothing, Nothing)
                    lv_Dictionary.Show()
            End Select
        End If
    End Sub

    Private Function btn_NewList_Click(sender As Object, e As EventArgs)

        Dim Message As String = "If you choose to download a new dictionary, a copy of your current dictionary will be saved with an auto-generated name to avoid overwriting it. Do you want to proceed with downloading a new dictionary?"
        Dim response = MsgBox(Message, MessageBoxButtons.YesNo Or vbQuestion, "Download New Dictionary")
        If response = DialogResult.Yes Then
            Dim newfilepath = CopyWithAutoName(pWordFilePath) ' Save a copy of the current phrase list with an auto-generated name to avoid overwriting the existing one.
            DownloadDictionary(Puzzle) ' Get a new Dictionary List from Camsoft.au
        Else
            MessageBox.Show("Download cancelled. Your current dictionary is safe.")
        End If
    End Function

    Private Sub btn_Restore_Click(sender As Object, e As EventArgs)
        Dim filter As String = ""
        Dim FileName As String = "
"
        If Puzzle = "pWord" Then
            filter = "Phrase List files (*.csv)|PhraseList*.csv|All files (*.*)|*.*"
            FileName = "PhraseList.csv"
        ElseIf Puzzle = "xWord" Then
            filter = "Crossword files (*.csv)|CrossList*.csv|All files (*.*)|*.*"
            FileName = "CrossList.csv"
        ElseIf Puzzle = "cWord" Then
            filter = "Codeword files (*.csv)|CodeList*.csv|All files (*.*)|*.*"
            FileName = "CodeList.csv"
        End If

        Dim openFileDialog As New OpenFileDialog With {
           .InitialDirectory = DefaultPath,
           .Filter = filter,
           .Title = "Select a dictionary file to restore"
                }

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            DictFilePath = openFileDialog.FileName

            Dim dest = Path.Combine(DefaultPath, FileName)
            '  copy to temp then atomic replace (requires dest exists)
            Dim temp = dest & ".tmp"
            System.IO.File.Copy(DictFilePath, temp, True)
            System.IO.File.Replace(temp, dest, Nothing)

            LoadFile(Nothing, Nothing) ' Load the selected dictionary into the list view
        End If
    End Sub


#End Region

#Region "MERGE PHRASE FILES"
    '====================== This is only needed if you want to update the phrase list with new phrases, and you have multiple batch files to merge. You can comment this out after running it once, and it will create a merged phrase list that will be used going forward. =====================

    Private Sub MergePhraseFiles(inputFolder As String, outputFile As String)
        '===================== MERGE PHRASE FILES ====================
        '===================== This is only needed if you want to update the phrase list with new phrases, and you have multiple batch files to merge. You can comment this out after running it once, and it will create a merged phrase list that will be used going forward. =====================

        Dim files As New List(Of String) From {
   DefaultPath & "\batchA.csv",
   DefaultPath & "\batchB.csv",
   DefaultPath & "\batchC.csv",
   DefaultPath & "\batchD.csv",
   DefaultPath & "\batchE.csv",
   DefaultPath & "\batchF.csv",
   DefaultPath & "\batchG.csv",
   DefaultPath & "\batchH.csv",
   DefaultPath & "\batchI.csv",
   DefaultPath & "\batchJ.csv",
   DefaultPath & "\batchK.csv",
   DefaultPath & "\batchL.csv",
   DefaultPath & "\batchM.csv",
   DefaultPath & "\batchN.csv"
  }

        ' Dictionary ensures uniqueness by phrase
        Dim merged As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        For Each filePath In Directory.EnumerateFiles(inputFolder, "batch*.csv")
            For Each line In File.ReadLines(filePath)
                If String.IsNullOrWhiteSpace(line) Then Continue For

                Dim parts = line.Split(","c, 2)
                If parts.Length < 2 Then Continue For

                Dim phrase = parts(0).Trim().ToUpper()
                Dim clue = parts(1).Trim()

                ' Only add if not already present
                If Not merged.ContainsKey(phrase) Then
                    merged.Add(phrase, clue)
                End If
            Next
        Next
        ' Write merged CSV
        Using sw As New StreamWriter(outputFile, False)

            For Each kv In merged
                '    sw.WriteLine($"{kv.Key},{kv.Value}")
            Next
        End Using
        MessageBox.Show($"Merged {merged.Count} unique phrases.")
    End Sub
#End Region
End Class