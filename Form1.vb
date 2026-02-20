Imports System.IO
Imports System.Drawing.Printing
Imports System.Configuration
Imports System.Diagnostics
Imports System.Net.Http
Imports System.Net

Public Class Form1

    Public DefaultPath As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + Path.DirectorySeparatorChar + "Crosswords"
    Public WordFilePath As String = Path.Combine(DefaultPath, "WordList.csv")
    Public PhraseFilePath As String = Path.Combine(DefaultPath, "PhraseList.csv")
    Public DictFilePath As String = Path.Combine(DefaultPath, "Wordlist.csv")
    Public ipAddress As String = "https://camsoft.au/cwg/PhraseList.csv"

#Region "DATA STRUCTURES"

    Friend WithEvents DgvGrid As DataGridView
    Const GridSize As Integer = 15
    Dim Grid(GridSize - 1, GridSize - 1) As Char
    Dim lstClues As New List(Of Clue)
    Dim Dictionary As New List(Of Clue)
    Dim WordLookup As New HashSet(Of String) ' For fast word existence checks
    Dim PuzzleNumber As Integer = 0
    Dim ClueNumbers(GridSize - 1, GridSize - 1) As Integer

    Public Class Clue
        Public Word As String
        Public Row As Integer
        Public Col As Integer
        Public IsAcross As Boolean
        Public Clue As String
        Public ClueNumber As Integer
        Public PlaceLetter As Boolean = False ' Used in crossword puzzles to indicate whether the letter should be revealed on the grid (for codeword puzzles, all letters are hidden)
    End Class

    Public Enum DirectionType
        Across
        Down
    End Enum


    Public Const Down As Boolean = False
    Public Puzzle As String = ""
    Public PrintPageIndex As Integer = 0
    Public WordList As List(Of String)
    Public PlacedWords As New List(Of String)
    Public Condensed As Boolean = False
    Dim rnd As New Random()

        Public Class LetterNumber
        Public Letter As String
        Public Number As Integer
    End Class

    Public Class CellData
        Public Letter As String
        Public AcrossNumber As Integer
        Public DownNumber As Integer
    End Class

    Public cD(,) As CellData
    Public WithEvents pd As New Printing.PrintDocument

    ' ===================== UI =====================
    Private pnlGrid As Panel
    Private pnlBottom As Panel
    Private pnlClues As Panel
    Public letterdg As New DataGridView
    Private btnEditDict As New Button
    Private WithEvents RbCrossword As New RadioButton
    Private WithEvents RbPhrase As New RadioButton
    Private WithEvents RbCodeword As New RadioButton
    Private lblNrClues As New Label
    Private cmbNrClues As New ComboBox
    Private NrOfClues As Integer = 3
    Private btnNew As New Button
    Private btnPrint As New Button
    Private PrintWords As Boolean = False
    '====================== LISTBOXES =====================
    Public lstAcross As New ListBox
    Public lstDown As New ListBox
    Public CluesAcross As New List(Of Clue)
    Public CluesDown As New List(Of Clue)
    Public txt_Alphabet As New TextBox
    Private LstRnr As List(Of Integer)
    Const EMPTY As Char = "."c


#End Region

    ' ===================== UI SETUP =====================

#Region "FORM LOAD"

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try
            Me.Text = "VB.NET Crossword"
            Me.Width = 1050
            Me.Height = 870

            Me.Location = New Point(100, 50)

            If My.Application.CommandLineArgs.Count > 0 Then
                Puzzle = My.Application.CommandLineArgs(0)
            End If

            ' Create directory for word list if it doesn't exist
            If Not Directory.Exists(DefaultPath) Then
                Directory.CreateDirectory(DefaultPath)
            End If

            'Check if there is a dictionary. If not, Create one from the included word list
            If Not File.Exists(DictFilePath) Then
                createDictionary("xWord", DictFilePath) ' Ensure dictionary exists
            End If

            If Not File.Exists(PhraseFilePath) Then
                createDictionary("pWord", PhraseFilePath)
            End If

            If Puzzle = "" Then Puzzle = "cWord" ' Default to codeword if not set

            ' Puzzle = "cWord"
            'Puzzle = "xWord"
            ' Puzzle = "pWord"

            If Puzzle = "pWord" Then
                DictFilePath = PhraseFilePath
            ElseIf Puzzle = "xWord" Then
                DictFilePath = WordFilePath
            ElseIf Puzzle = "cWord" Then
                DictFilePath = WordFilePath
            End If

            LoadDictionary(DictFilePath)

            SetupGrid()
            InitGrid()
            SetupUI()
            If Puzzle = "cWord" Then
                SetupLetterGrid()
                SetupTextBox()
            End If
            GeneratePuzzle()

        Catch ex As Exception
            MessageBox.Show("An error occurred during initialization: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

#End Region
#Region "DICTIONARY"
    Sub LoadDictionary(path As String) ' Load the word list from CSV file
        Dim line As String
        Dim entry As String
        Dim Word As String
        Dim Clu As String
        Dictionary.Clear()
        Try

            Using sr As New StreamReader(path)
                While Not sr.EndOfStream
                    line = sr.ReadLine()
                    Dim parts = line.Split(","c, 2)
                    Dim w = parts(0).Trim().ToUpper()
                    Clu = If(parts.Length > 1, parts(1).Trim(), "")
                    If Len(w) > 12 Then Continue While ' Skip lines that are too long
                    Dictionary.Add(New Clue With {.Word = w, .Clue = Clu})
                End While
                sr.Close()
            End Using
            RandomiseDictionary(100) ' Load a random selection of 100 words from the dictionary for puzzle generation. This helps ensure variety in the generated puzzles and can improve performance by working with a smaller set of words during placement.
        Catch ex As Exception
            MessageBox.Show("Failed to load dictionary: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Sub RandomiseDictionary(count As Integer)
        Dim selected As New List(Of Clue)
        Dim rnd As New Random()
        Dim usedIndexes As New HashSet(Of Integer)
        Dim i As Integer = 0

        Try
            While selected.Count < 100 And i < count * 10 ' Add a safety limit to prevent infinite loops
                Dim index = rnd.Next(Dictionary.Count)
                If Not usedIndexes.Contains(index) Then
                    usedIndexes.Add(index)
                    selected.Add(Dictionary(index))
                End If
                i += 1
            End While

            If i >= count * 10 Then
                MsgBox("Selected " & selected.Count.ToString() & " unique words out of requested " & count.ToString() & ". Consider increasing the word list or reducing the requested count.", MessageBoxButtons.OK, "Warning")
                Dim response = MsgBox("Do you want to reload the dictionary?", MessageBoxButtons.YesNo, "Reload Dictionary")
                If response = DialogResult.Yes Then
                    Dim Message As String = "If you choose to download a new dictionary, a copy of your current dictionary will be saved with an auto-generated name to avoid overwriting it. Do you want to proceed with downloading a new dictionary?"
                    Dim response2 = MsgBox(Message, MessageBoxButtons.YesNo Or vbQuestion, "Download New Dictionary")
                    If response2 = DialogResult.Yes Then
                        Dim newfilepath = CopyWithAutoName(PhraseFilePath) ' Save a copy of the current phrase list with an auto-generated name to avoid overwriting the existing one.
                        DownloadPhraseList() ' Get a new Phrase List from Camsoft.au

                        LoadDictionary(DictFilePath) ' Reload the dictionary and try again
                    End If
                    Return
                End If
            End If

            Dictionary = selected.OrderByDescending(Function(x) x.Word.Length).ToList() ' Sort by length for better puzzle generation

        Catch ex As Exception
            MessageBox.Show("An error occurred while randomizing the dictionary: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Function CopyWithAutoName(sourceFile As String) As String
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

#Region "DOWNLOAD PHRASE LIST FROM CAMSOFT.AU"
    ' Downloads the list of phrases from Camsoft.au website.

    'Probably not be used now, as the phrase list is included in the installer, but it's here if we want to update the list in the future without having to release a new version of the software. It can also be used to get a new list of phrases if the user wants to refresh their list.
    Sub DownloadPhraseList()

        Try
            Using client As New WebClient()
                client.DownloadFile(ipAddress, PhraseFilePath)
            End Using

            Dim DictList As New List(Of Clue)

            Using sr As New StreamReader(PhraseFilePath)
                While Not sr.EndOfStream
                    Dim line As String = sr.ReadLine()
                    Dim parts As String() = line.Split(","c)
                    If parts.Length >= 2 Then
                        Dim entry As New Clue With {
                            .Word = parts(0).Trim(),
                            .Clue = "Auto generated Clue"
                        }
                        If Len(entry.Word) <= 12 AndAlso Len(entry.Word) > 4 Then 'Select only phrases of appropriate length for the puzzle
                            DictList.Add(entry)
                        End If
                    End If
                End While
            End Using

            ShuffleDictList(DictList)
            Dim selected1000 As List(Of Clue) = DictList.Take(1000).ToList()

            SaveDictionary(PhraseFilePath, selected1000)

        Catch ex As Exception
            MsgBox("Failed to Load Phrase List from Camsoft.au")
        End Try
    End Sub
    Private Sub ShuffleDictList(list As List(Of Clue))

        Dim rnd As New Random()
        Try
            For i As Integer = list.Count - 1 To 1 Step -1
                Dim j As Integer = rnd.Next(i + 1)

                Dim temp = list(i)
                list(i) = list(j)
                list(j) = temp
            Next

        Catch ex As Exception
            MessageBox.Show("An error occurred while shuffling the dictionary: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub


    Private Sub SaveDictionary(filepath As String, wlist As List(Of Clue))
        Try
            Using writer As New StreamWriter(filepath, False)
                For Each entry In wlist
                    If entry.Word = "Word" Then Continue For ' Skip any words that contain "Word" 
                    writer.WriteLine($"{entry.Word},{entry.Clue}")
                Next
                writer.Close()
            End Using
        Catch ex As Exception
            MessageBox.Show("Failed to save Dictionary.")
        End Try

    End Sub

#End Region

#Region "GRID SETUP"

    Private Sub SetupGrid()
        Try

            ' ---------- Grid panel ----------
            pnlGrid = New Panel With {
        .Dock = DockStyle.Left,
        .Width = 756}

            ' ---------- DataGridView ----------
            DgvGrid = New DataGridView With {
            .Name = "dgvGrid",
            .Dock = DockStyle.Fill,
            .AllowUserToAddRows = False,
            .AllowUserToResizeColumns = False,
            .AllowUserToResizeRows = False,
            .RowHeadersVisible = False,
            .ColumnHeadersVisible = False,
            .ScrollBars = ScrollBars.None,
            .Font = New Font("Segoe UI", 14, FontStyle.Bold)
        }

            ' ---------- Columns ----------
            DgvGrid.Columns.Clear()
            For i = 0 To GridSize - 1
                DgvGrid.Columns.Add($"C{i}", "")
                DgvGrid.Columns(i).Width = 50
            Next

            ' ---------- Rows ----------
            DgvGrid.RowCount = GridSize
            For Each r As DataGridViewRow In DgvGrid.Rows
                r.Height = 49
            Next

            ' ---------- Assemble ----------
            pnlGrid.Controls.Add(DgvGrid)
            Me.Controls.Add(pnlGrid)

            ReDim cD(GridSize - 1, GridSize - 1)

            '---------- Initialize CellData array ----------
            For r = 0 To GridSize - 1
                For c = 0 To GridSize - 1
                    cD(r, c) = New CellData With {
                        .Letter = "",
                        .AcrossNumber = 0,
                        .DownNumber = 0
                    }
                Next
            Next
        Catch ex As Exception
            MessageBox.Show("An error occurred while setting up the grid: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub


    Private Sub SetupUI()

        Try
            ' ---------- Bottom panel ----------
            pnlBottom = New Panel With {
        .Dock = DockStyle.Bottom,
        .Width = Me.ClientSize.Width,
        .Height = 90}

            ' ---------- Edit Dictionary ----------
            btnEditDict.Font = New Font("Segoe UI", 12, FontStyle.Bold)
            btnEditDict.Text = "Edit Dictionary"
            btnEditDict.AutoSize = True
            btnEditDict.Location = New Point(10, 20)

            '----------Puzzle type label----------
            RbPhrase.Font = New Font("Segoe UI", 10, FontStyle.Bold)
            RbPhrase.Text = "PhraseWord"
            RbPhrase.AutoSize = True
            RbPhrase.Location = New Point(200, 10)

            RbCodeword.Font = New Font("Segoe UI", 10, FontStyle.Bold)
            RbCodeword.Text = "CodeWord"
            RbCodeword.AutoSize = True
            RbCodeword.Location = New Point(200, 30)

            RbCrossword.Font = New Font("Segoe UI", 10, FontStyle.Bold)
            RbCrossword.Text = "CrossWord"
            RbCrossword.AutoSize = True
            RbCrossword.Location = New Point(200, 50)

            If Puzzle = "cWord" Then RbCodeword.Text = "CodeWord"
            If Puzzle = "pWord" Then RbPhrase.Text = "PhraseWord"
            If Puzzle = "xWord" Then RbCrossword.Text = "CrossWord"


            ' ---------- Number of clues ----------
            lblNrClues.Font = New Font("Segoe UI", 12, FontStyle.Bold)
            lblNrClues.Text = "Number of Clues:"
            lblNrClues.AutoSize = True
            lblNrClues.Location = New Point(330, 25)

            cmbNrClues.Font = New Font("Segoe UI", 12, FontStyle.Bold)
            cmbNrClues.Items.AddRange({"0", "1", "2", "3", "4", "5"})
            cmbNrClues.SelectedIndex = 3
            cmbNrClues.DropDownStyle = ComboBoxStyle.DropDownList
            cmbNrClues.Width = 60
            cmbNrClues.Location = New Point(480, 22)

            ' ---------- New button ----------
            btnNew.Font = New Font("Segoe UI", 12, FontStyle.Bold)
            btnNew.Text = "New"
            btnNew.Size = New Size(90, 32)
            btnNew.Location = New Point(550, 20)

            ' ---------- Print button ----------
            btnPrint.Font = New Font("Segoe UI", 12, FontStyle.Bold)
            btnPrint.Text = "Print"
            btnPrint.Size = New Size(90, 32)
            btnPrint.Location = New Point(650, 20)

            ' ---------- Add to bottom panel ----------
            Me.Controls.Add(pnlBottom)
            pnlBottom.Controls.AddRange({
            btnEditDict,
            RbPhrase,
            RbCrossword,
            RbCodeword,
            lblNrClues,
            cmbNrClues,
            btnNew,
            btnPrint})

            ' ---------- Crossword clue panels ----------
            If Puzzle = "xWord" OrElse Puzzle = "pWord" Then

                pnlClues = New Panel With {
                .Dock = DockStyle.Right,
                .Width = 260,
                .Height = 400
            }

                lstAcross.Font = New Font("Segoe UI", 10)
                lstAcross.Dock = DockStyle.Top
                lstAcross.Width = pnlClues.Width
                lstAcross.ScrollAlwaysVisible = False

                lstDown.Font = New Font("Segoe UI", 10)
                lstDown.Width = pnlClues.Width
                lstDown.ScrollAlwaysVisible = False
                lstDown.Location = New Point(0, pnlClues.Height \ 2)

                pnlClues.Controls.Add(lstDown)
                pnlClues.Controls.Add(lstAcross)
                Me.Controls.Add(pnlClues)

                Me.Width = 1050
            Else
                Me.Width = 760
            End If

            ' ---------- Handlers ----------
            AddHandler btnNew.Click, AddressOf BtnNew_Click
            AddHandler btnPrint.Click, AddressOf BtnPrint_Click
            AddHandler btnEditDict.Click, AddressOf EditDictionary
            AddHandler RbCodeword.Click, AddressOf RadioButton_CheckedChanged
            AddHandler RbPhrase.Click, AddressOf RadioButton_CheckedChanged
            AddHandler RbCrossword.Click, AddressOf RadioButton_CheckedChanged
            AddHandler cmbNrClues.SelectedIndexChanged, AddressOf cmbNrClues_IndexChanged
            AddHandler DgvGrid.CellPainting, AddressOf dgvGrid_CellPainting

        Catch ex As Exception
            MessageBox.Show("An error occurred while setting up the UI: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub


    Sub SetupLetterGrid() ' Setup the letter-number mapping grid

        Try
            letterdg = New DataGridView With {
            .Name = "letterdg",
            .Dock = DockStyle.Bottom,
            .Width = 500,
            .AllowUserToAddRows = False,
            .AllowUserToResizeColumns = False,
            .AllowUserToResizeRows = False,
            .RowHeadersVisible = False,
            .ColumnHeadersVisible = False,
            .Font = New Font("Segoe UI", 12, FontStyle.Bold)
        }

            For i = 0 To 12
                letterdg.Columns.Add("", "")
                letterdg.Columns(i).Width = 45
            Next

            letterdg.Rows.Add(4)
            For Each r As DataGridViewRow In letterdg.Rows
                r.Height = 45
            Next

            Me.Controls.Add(letterdg)

            For i = 0 To 12
                letterdg.Rows(0).Cells(i).Value = i + 1
                letterdg.Rows(2).Cells(i).Value = i + 14
            Next

            letterdg.Visible = False ' Hide the grid on the screen. Only need it on the printed version

        Catch ex As Exception
            MessageBox.Show("An error occurred while setting up the letter grid: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Sub SetupTextBox() ' Setup the alphabet textbox
        Try
            txt_Alphabet = New TextBox With {
            .Multiline = False,
            .Width = 500,
            .Height = 50,
            .Font = New Font("Consolas", 12),
            .Location = New Point(300, 850),
            .ScrollBars = ScrollBars.Vertical
        }
            txt_Alphabet.ReadOnly = True
            txt_Alphabet.Text = "A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z"
            Me.Controls.Add(txt_Alphabet)
        Catch ex As Exception
            MessageBox.Show("An error occurred while setting up the alphabet textbox: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub InitGrid() ' Initialize the grid and UI for a new puzzle
        Try
            PuzzleNumber = rnd.Next(1, 1000)
            Me.Text = "VB.NET Codeword" & " - Puzzle #" & PuzzleNumber.ToString()

            lstClues.Clear()
            PlacedWords.Clear()
            lstAcross.Items.Clear()
            lstAcross.Items.Add("Across Clues:")
            lstDown.Items.Clear()
            lstDown.Items.Add("Down Clues:")


            For r = 0 To GridSize - 1
                For c = 0 To GridSize - 1
                    Grid(r, c) = EMPTY
                Next
            Next

            For Each col As DataGridViewColumn In DgvGrid.Columns
                col.Width = 50
            Next

            For Each row As DataGridViewRow In DgvGrid.Rows
                For Each cell As DataGridViewCell In row.Cells
                    cell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                Next
            Next

        Catch ex As Exception
            MessageBox.Show("An error occurred while initializing the grid: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


#End Region

#Region "CROSSWORD ENGINE"

    Sub GeneratePuzzle() ' This is the main puzzle generation routine
        Try

            If Dictionary.Count = 0 Then
                MessageBox.Show("Dictionary is empty. Check the word list file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            Dim shuffled = Dictionary.OrderBy(Function(x) rnd.Next()).ToList()
            Dim Words = shuffled.Take(60).OrderByDescending(Function(x) x.Word.Length).ToList()
            If Words.Count = 0 Then Return

            ' Place longest of the random set in the middle of the grid, horizontally
            PlaceWord(Words(0), GridSize \ 2, (GridSize - Words(0).Word.Length) \ 2, DirectionType.Across)
            ' Place remaining words
            For i = 1 To Words.Count - 1
                If WordUsed(Words(i).Word) Then Continue For
                TryPlaceWord(Words(i))
            Next

            RenderGrid()

        Catch ex As Exception
            MessageBox.Show("An error occurred while generating the puzzle: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    '===================== PLACEMENT LOGIC =====================

    Function TryPlaceWord(entry As Clue) As Boolean
        Try

            Dim word = entry.Word.ToUpper()

            For r = 0 To GridSize - 1
                For c = 0 To GridSize - 1

                    ' Try ACROSS
                    For i = 0 To word.Length - 1
                        If Grid(r, c) = word(i) Or Grid(r, c) = EMPTY Then
                            Dim startCol = c - i
                            If startCol >= 0 AndAlso CanPlace(word, r, startCol, DirectionType.Across) Then
                                PlaceWord(entry, r, startCol, DirectionType.Across)
                                Return True
                            End If
                        End If
                    Next

                    ' Try DOWN
                    For i = 0 To word.Length - 1
                        If Grid(r, c) = word(i) Or Grid(r, c) = EMPTY Then
                            Dim startRow = r - i
                            If startRow >= 0 AndAlso CanPlace(word, startRow, c, DirectionType.Down) Then
                                PlaceWord(entry, startRow, c, DirectionType.Down)
                                Return True
                            End If
                        End If
                    Next

                Next
            Next

            Return False
        Catch ex As Exception
            MessageBox.Show("An error occurred while trying to place a word: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try

    End Function

    Function CanPlace(word As String, row As Integer, col As Integer, dir As DirectionType) As Boolean

        Try
            If dir = DirectionType.Across Then

                ' --- Bounds ---
                If col < 0 OrElse col + word.Length > GridSize Then Return False

                ' --- Cell before start ---
                If col > 0 AndAlso Grid(row, col - 1) <> EMPTY Then Return False

                ' --- Cell after end ---
                If col + word.Length < GridSize AndAlso
               Grid(row, col + word.Length) <> EMPTY Then Return False

                For i As Integer = 0 To word.Length - 1

                    Dim r = row
                    Dim c = col + i
                    Dim gridChar = Grid(r, c)
                    Dim letter = word(i)

                    ' --- Letter compatibility ---
                    If gridChar <> EMPTY AndAlso gridChar <> letter Then
                        Return False
                    End If

                    ' --- If NOT crossing, validate vertical fragment ---
                    If gridChar = EMPTY Then
                        If CreatesInvalidVerticalWord(r, c, letter) Then
                            Return False
                        End If
                    End If

                Next

            Else ' DOWN

                ' --- Bounds ---
                If row < 0 OrElse row + word.Length > GridSize Then Return False

                ' --- Cell before start ---
                If row > 0 AndAlso Grid(row - 1, col) <> EMPTY Then Return False

                ' --- Cell after end ---
                If row + word.Length < GridSize AndAlso
               Grid(row + word.Length, col) <> EMPTY Then Return False

                For i As Integer = 0 To word.Length - 1

                    Dim r = row + i
                    Dim c = col
                    Dim gridChar = Grid(r, c)
                    Dim letter = word(i)

                    ' --- Letter compatibility ---
                    If gridChar <> EMPTY AndAlso gridChar <> letter Then
                        Return False
                    End If

                    ' --- If NOT crossing, validate horizontal fragment ---
                    If gridChar = EMPTY Then
                        If CreatesInvalidHorizontalWord(r, c, letter) Then
                            Return False
                        End If
                    End If

                Next

            End If

            Return True

        Catch ex As Exception
            MessageBox.Show("An error occurred while checking if a word can be placed: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try

    End Function

    Function CreatesInvalidVerticalWord(row As Integer,
                                    col As Integer,
                                    newLetter As Char) As Boolean
        Try
            Dim r = row
            Dim word As String = ""

            ' Move up
            While r > 0 AndAlso Grid(r - 1, col) <> EMPTY
                r -= 1
            End While

            ' Build downward
            Dim tempRow = r

            While tempRow < GridSize AndAlso
              (Grid(tempRow, col) <> EMPTY OrElse tempRow = row)

                If tempRow = row Then
                    word &= newLetter
                Else
                    word &= Grid(tempRow, col)
                End If
                tempRow += 1
            End While

            If word.Length > 1 AndAlso Not WordLookup.Contains(word) Then
                Return True
            End If

            Return False
        Catch ex As Exception
            MessageBox.Show("An error occurred while checking vertical word validity: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return True ' Return true to prevent placement if there's an error
        End Try

    End Function

    Function CreatesInvalidHorizontalWord(row As Integer, col As Integer, newLetter As Char) As Boolean
        Try
            Dim c = col
            Dim word As String = ""

            ' Move left
            While c > 0 AndAlso Grid(row, c - 1) <> EMPTY
                c -= 1
            End While

            ' Build rightward
            Dim tempCol = c

            While tempCol < 15 AndAlso' Width AndAlso
              (Grid(row, tempCol) <> EMPTY OrElse tempCol = col)

                If tempCol = col Then
                    word &= newLetter
                Else
                    word &= Grid(row, tempCol)
                End If

                tempCol += 1
            End While

            If word.Length > 1 AndAlso Not WordLookup.Contains(word) Then
                Return True
            End If

            Return False

        Catch ex As Exception
            MessageBox.Show("An error occurred while checking horizontal word validity: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return True ' Return true to prevent placement if there's an error
        End Try

    End Function


    Function WordUsed(Word As String) 'Check if word has already been placed
        Try
            For Each item In PlacedWords
                If Word = item.ToString() Then 'word has been used,
                    Return True
                End If
            Next
            Return False

        Catch ex As Exception
            MessageBox.Show("An error occurred while checking if a word has already been used: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return True
        End Try
    End Function

    '============== END OF PLACEMENT LOGIC ==============

    '============== PLACEMENT ==============
    'Sub PlaceWord(entry As ClueEntry, row As Integer, col As Integer, dir As    DirectionType)
    Sub PlaceWord(entry As Clue, row As Integer, col As Integer, dir As DirectionType)


        Dim word = entry.Word.ToUpper()
        Dim NrOfLetters As String = word.Length.ToString()
        NrOfLetters = " (" & NrOfLetters & ") "

        Try
            If dir = DirectionType.Across Then
                For i = 0 To word.Length - 1
                    Grid(row, col + i) = word(i)
                Next
            Else ' DOWN
                For i = 0 To word.Length - 1
                    Grid(row + i, col) = word(i)
                Next
            End If

            If Puzzle = "xWord" OrElse Puzzle = "pWord" Then
                If dir = DirectionType.Across Then
                    lstClues.Add(New Clue With {.Word = word, .Row = row, .Col = col, .IsAcross = True, .Clue = entry.Clue})
                    PlacedWords.Add(word)
                Else
                    lstClues.Add(New Clue With {.Word = word, .Row = row, .Col = col, .IsAcross = False, .Clue = entry.Clue})
                    PlacedWords.Add(word)
                End If
            End If

            Dim ex As Exception

        Catch ex As Exception
            MessageBox.Show("An error occurred while placing a word: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub


    '============== END OF PLACEMENT LOGIC ==============

#End Region

#Region "RENDER"
    '=============Up until now all operations have been on the Grid() array. This routine renders the grid to the DataGridView and applies cell coloring. It also assigns clue numbers for crosswords and random letter-number mappings for codeword puzzles. =============

    Sub RenderGrid() ' Render the grid to the DataGridView

        Dim ShowSolution As Boolean = True 'If true shows letters in cells. Set to false for codeword puzzles, true for crosswords

        For r = 0 To GridSize - 1
            For c = 0 To GridSize - 1
                Dim cell = DgvGrid.Rows(r).Cells(c)

                If Grid(r, c) = "."c Then
                    cell.Style.BackColor = Color.Black
                    cell.Value = ""
                Else
                    cell.Style.BackColor = Color.White
                    cell.Value = If(ShowSolution, Grid(r, c), " ")
                End If

            Next
        Next

        If Puzzle = "xWord" OrElse Puzzle = "pWord" Then
            Dim ClueNumber As Integer = 0
            reorderClueList() ' Ensure clues are in ascending order based on their position on the grid before assigning numbers
            For i = 0 To lstClues.Count - 1
                Dim row = lstClues(i).Row
                Dim col = lstClues(i).Col
                Dim across = lstClues(i).IsAcross
                ClueNumber = lstClues(i).ClueNumber

                If across Then
                    cD(row, col).AcrossNumber = ClueNumber
                    lstClues(i).ClueNumber = ClueNumber
                    lstClues(i).IsAcross = True
                Else
                    cD(row, col).DownNumber = ClueNumber
                    lstClues(i).ClueNumber = ClueNumber
                    lstClues(i).IsAcross = False
                End If
            Next

            AddCluesToList()
        Else 'Puzzle is codeword, so we need to assign random numbers to letters and populate the letter grid
            AssignNumbersToLetters()
        End If

    End Sub
    Sub AddCluesToList()

        For Each clue In lstClues
            If clue.IsAcross Then
                lstAcross.Items.Add($"{clue.ClueNumber}: {clue.Clue}")
            Else
                lstDown.Items.Add($"{clue.ClueNumber}:  {clue.Clue}")
            End If
        Next

        lstAcross.Height = lstAcross.Items.Count * 20 + 30 ' Adjust height based on number of clues
        lstDown.Height = lstDown.Items.Count * 20 + 30 ' Adjust height based on number of clues
        lstDown.Location = New Point(0, pnlClues.Height \ 2) ' Keep down clues in the bottom half of the panel

    End Sub

    Public Sub AssignNumbersToLetters() ' Assign random numbers to letters A-Z in the grid and populate letter grid with numbers.
        Dim Letter As String = ""
        Dim Number As Integer
        Dim ch As Char
        Dim rnd As New Random()
        Dim r As Integer
        Dim c As Integer
        Dim i As Integer

        'Generate a list of letters A-Z and assign each a random number 1-26
        Dim Alphabet As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Dim combined As New List(Of LetterNumber)

        Dim lst_letters As New List(Of Char)
        For i = 0 To 25
            ch = Alphabet(i)
            lst_letters.Add(ch)
        Next

        Dim lst_Numbers As New List(Of Integer)
        For i = 1 To 26
            lst_Numbers.Add(i)
        Next
        ' Shuffle the numbers
        lst_Numbers = lst_Numbers.OrderBy(Function(x) rnd.Next()).ToList()

        'Combine the two lists
        If lst_letters.Count <> lst_Numbers.Count Then
            Throw New InvalidOperationException("Lists must be the same length")
            Exit Sub
        End If

        For i = 0 To Math.Min(lst_letters.Count, lst_Numbers.Count) - 1
            combined.Add(New LetterNumber With {
        .Letter = lst_letters(i),
        .Number = lst_Numbers(i)
    })
        Next

        With DgvGrid
            For r = 0 To GridSize - 1
                For c = 0 To GridSize - 1
                    Dim v = .Rows(r).Cells(c).Value
                    If v Is Nothing Then Continue For

                    For Each item In combined
                        If item.Letter = v Then
                            Number = item.Number
                            .Rows(r).Cells(c).Value = CStr(Number)
                            Exit For
                        End If
                    Next
                Next
            Next
        End With

        '===============Replace letters with numbers on the grid===========

        Dim UsedNumbers As New List(Of Integer)
        Dim hits As Integer = 0 ' count of letters placed
        Dim rowIndex As Integer = 1
        Dim iterations As Integer = GridSize * GridSize 'to prevent endless loops
        Dim nrLoops As Integer = 0

        If combined Is Nothing OrElse combined.Count = 0 Then Exit Sub
        ' Shuffle the combined list
        combined = combined.OrderBy(Function(x) rnd.Next()).ToList()

        ' Clear Letters from LetterGrid()
        With letterdg
            For c = 0 To .ColumnCount - 1
                .Rows(1).Cells(c).Value = Nothing
                .Rows(3).Cells(c).Value = Nothing
            Next
        End With

        i = 1
        NrOfClues = Int(cmbNrClues.SelectedIndex)
        Do Until hits = NrOfClues Or nrLoops = iterations
            nrLoops += 1
            'take an entry from combined list
            Dim Clue As LetterNumber
            Clue = combined(i)
            Letter = Clue.Letter
            Number = Clue.Number   ' 1–26

            If UsedNumbers.Contains(Number) Then Continue Do ' skip duplicates
            UsedNumbers.Add(Number)

            'Confirm that the letter is used in the puzzle
            Dim LetterUsed As Boolean = False

            For r = 0 To GridSize - 1
                For c = 0 To GridSize - 1
                    If Grid(r, c).ToString() = Letter Then
                        LetterUsed = True
                        If Number > 13 Then 'it goes in row 3
                            rowIndex = 3
                            letterdg.Rows(rowIndex).Cells(Number - 14).Value = Letter
                        Else
                            rowIndex = 1 ' it goes in row 1
                            letterdg.Rows(rowIndex).Cells(Number - 1).Value = Letter
                        End If
                        hits += 1
                        c = GridSize - 1 'force exit from inner loop
                        r = GridSize - 1 'force exit from outer loop
                    End If
                Next
            Next

            i += 1 'increment the index
            If i >= 25 Then i = 0 ' reset to 0 if exceeding list size
        Loop

    End Sub
    Sub reorderClueList()
        ' This routine ensures that the clues are listed in ascending order based on their position on the grid
        Dim ClueNumber As Integer = 0
        lstClues = lstClues.OrderBy(Function(c) c.Row).ThenBy(Function(c) c.Col).ToList()

        For Each c In lstClues
            ClueNumber += 1
            c.ClueNumber = ClueNumber
        Next

    End Sub

    Function GenerateRandomNumber(n As Integer)
        Dim rnd As New Random()
        Dim nums As New HashSet(Of Integer)
        Dim counter As Integer = 0
        While counter < n
            nums.Add(rnd.Next(1, n))
            counter += 1
        End While
        Dim result = nums.ToList()
        Return result
    End Function


#End Region

    '==================== CELL PAINTING FOR CLUE NUMBERS ====================
#Region "Cell Painting"
    Private Sub dgvGrid_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs)
        Try
            If e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Return
            Dim Row = e.RowIndex
            Dim Col = e.ColumnIndex
            ' Let grid paint background & borders
            e.Paint(e.CellBounds, DataGridViewPaintParts.All)

            Using f As New Font("Segoe UI", 7, FontStyle.Regular), b As New SolidBrush(Color.Black)

                ' Across number (top-right)
                If cD(Row, Col).AcrossNumber > 0 Then
                    e.Graphics.DrawString(
                        cD(Row, Col).AcrossNumber.ToString(),
                        f, b,
                        e.CellBounds.Left + 33,
                        e.CellBounds.Top + 2
                )
                End If

                ' Down number (bottom-left)
                If cD(Row, Col).DownNumber > 0 Then
                    e.Graphics.DrawString(
                    cD(Row, Col).DownNumber.ToString(),
                    f, b,
                    e.CellBounds.Left + 2,
                    e.CellBounds.Bottom - f.Height - 2
                )
                End If
            End Using

            e.Handled = True

        Catch ex As Exception
            MessageBox.Show("An error occurred while painting the grid cells: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


#End Region

#Region "Button Handlers"
    ' ===================== BUTTONS =====================
    Sub EditDictionary(sender As Object, e As EventArgs)
        Me.Hide()
        DictionaryGenerator.Show()
    End Sub
    Sub RadioButton_CheckedChanged(sender As Object, e As EventArgs)
        If sender Is RbCodeword Then
            Me.Close()
            RestartApp("cWord")
        ElseIf sender Is RbCrossword Then
            Me.Close()
            RestartApp("xWord")
        ElseIf sender Is RbPhrase Then
            Me.Close()
            RestartApp("pWord")
        End If
    End Sub
    Private Sub BtnNew_Click(sender As Object, e As EventArgs)
        RestartApp(Puzzle)
    End Sub
    Private Sub RestartApp(puzzleType As String)
        Dim exePath As String = Application.ExecutablePath
        Dim psi As New ProcessStartInfo(exePath, puzzleType)
        Process.Start(psi)
        Application.Exit()
    End Sub
    Private Sub BtnPrint_Click(sender As Object, e As EventArgs)

        Dim dlg As New PrintDialog With {.Document = pd}
        reorderClueList() ' Ensure clues are listed in ascending order on the grid

        If dlg.ShowDialog() = DialogResult.OK Then
            PrintPageIndex = 0
            If Puzzle = "cWord" Then
                Dim response = MsgBox("Print Placed Words List", vbYesNo, "Printing")
                If response = vbYes Then PrintWords = True Else PrintWords = False
            End If

            If Puzzle = "xWord" OrElse Puzzle = "pWord" Then
                pd.DefaultPageSettings.Landscape = True
                Dim Number As Integer = (lstClues.Count) - 1
                LstRnr = GenerateRandomNumber(Number)
            Else
                pd.DefaultPageSettings.Landscape = False
            End If

            pd.Print()
        End If

    End Sub

    Private Sub cmbNrClues_IndexChanged(sender As Object, e As EventArgs)
        NrOfClues = Int(cmbNrClues.SelectedIndex)
    End Sub

#End Region

#Region "Printing"
    ' ===================== PRINTING =====================

    Private Sub pd_PrintPage(sender As Object, e As PrintPageEventArgs) Handles pd.PrintPage

        '-----This is the main print page handler-----
        '-----This routine handles multi-page printing for codeword puzzles-----

        '    Dim n As New Random(15)
        Try

            Select Case PrintPageIndex
                Case 0
                    PrintPuzzlePage(e)
                    PrintLetterGrid(e)
                    PrintAlphabetGrid(e)
                    PrintClueLists(e)
                    e.HasMorePages = True
                Case 1
                    PrintAnswerKeyPage(e)
                    PrintPlacedWords(e)
                    PrintxWordAnswers(e)
                    PrintClueLists(e)
                    e.HasMorePages = False
            End Select
            PrintPageIndex += 1
        Catch ex As Exception
            MessageBox.Show("An error occurred during printing: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub


    Sub PrintClueLists(e As PrintPageEventArgs)
        'This routine prints the clues pages. Not used in codeword puzzles
        If Puzzle <> "xWord" And Puzzle <> "pWord" Then Return

        Dim g = e.Graphics
        Dim font As New Font("Segoe UI", 9)
        Dim y = 20
        Dim ClueNumber As Integer = 0
        Try

            g.DrawString("Across Clues: ", font, Brushes.Black, 800, y)
            y += 20

            For Each c In lstClues.Where(Function(x) x.IsAcross)
                ClueNumber = c.ClueNumber
                g.DrawString(ClueNumber & " :  " & c.Clue, font, Brushes.Black, 800, y)
                y += 15
            Next

            y += 20
            g.DrawString("Down Clues: ", font, Brushes.Black, 800, y)
            y += 20
            For Each c In lstClues.Where(Function(x) Not x.IsAcross)
                ClueNumber = c.ClueNumber
                g.DrawString(ClueNumber & " :  " & c.Clue, font, Brushes.Black, 800, y)
                y += 15
            Next

        Catch ex As Exception
            MessageBox.Show("An error occurred while printing clues: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Private Sub PrintPuzzlePage(e As PrintPageEventArgs)
        Try
            Dim g = e.Graphics
            g.DrawString("CODEWORD PUZZLE" & " - Puzzle #" & PuzzleNumber.ToString(),
                     New Font("Segoe UI", 16, FontStyle.Bold),
                     Brushes.Black, 50, 20)

            PrintGrid(e, printLetters:=False)
        Catch ex As Exception
            MessageBox.Show("An error occurred while printing the puzzle page: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Private Sub PrintAnswerKeyPage(e As PrintPageEventArgs)
        Try
            If Puzzle <> "cWord" Then Return

            'pd.DefaultPageSettings.Landscape = False

            Dim g = e.Graphics
            g.DrawString("ANSWER KEY" & " - Puzzle #" & PuzzleNumber.ToString(),
                     New Font("Segoe UI", 16, FontStyle.Bold),
                     Brushes.Black, 50, 20)

            PrintGrid(e, printLetters:=True)
        Catch ex As Exception
            MessageBox.Show("An error occurred while printing the answer key page: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub PrintxWordAnswers(e As PrintPageEventArgs)
        Try
            Dim g = e.Graphics
            g.DrawString("CROSS WORD PUZZLE" & " - Puzzle #" & PuzzleNumber.ToString(),
                     New Font("Segoe UI", 16, FontStyle.Bold),
                     Brushes.Black, 50, 20)

            PrintGrid(e, printLetters:=True)
        Catch ex As Exception
            MessageBox.Show("An error occurred while printing the crossword answers: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub PrintPlacedWords(e As PrintPageEventArgs)
        Try
            If PrintWords = False Then Return

            Dim g = e.Graphics
            Dim y = 750
            Dim font As New Font("Segoe UI", 12)
            g.DrawString("Placed Words:", New Font(font, FontStyle.Bold), Brushes.Black, 50, y)
            y += 30
            If PlacedWords.Count = 0 Then
                g.DrawString("No words placed.", font, Brushes.Black, 50, y)
                Return
            End If
            Dim i As Integer = 0
            Dim l As Integer = 50

            For Each w In PlacedWords
                g.DrawString(w, font, Brushes.Black, l, y)
                l += Len(PlacedWords(0)) * 16
                i += 1
                If i Mod 5 = 0 Then 'Print five words per line
                    y += 20
                    l = 50
                End If
            Next
        Catch ex As Exception
            MessageBox.Show("An error occurred while printing the placed words: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub PrintAlphabetGrid(e As PrintPageEventArgs)
        Try
            If Puzzle <> "cWord" Then Return 'Only print alphabet grid for codeword puzzles
            pd.DefaultPageSettings.Landscape = False

            Dim g = e.Graphics
            Dim size = 41
            Dim x = 140
            Dim y = 700

            Using p As New Pen(Color.Black)
                ' draw alphabet BELOW the grid
                Dim textY As Integer = y + size + 10
                g.DrawString(txt_Alphabet.Text,
                     New Font("Segoe UI", 16, FontStyle.Bold), Brushes.Black, x, textY)
            End Using
        Catch ex As Exception
            MessageBox.Show("An error occurred while printing the alphabet grid: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Private Sub PrintLetterGrid(e As PrintPageEventArgs)

        Try
            If Puzzle <> "cWord" Then Return

            pd.DefaultPageSettings.Landscape = False

            Dim g = e.Graphics        ' ✅ THIS is the printer page
            Dim size = 41
            Dim sx = 150
            Dim sy = 850              ' OK: lower on the page
            Dim f As New Font("Segoe UI", 12, FontStyle.Bold)

            Using p As New Pen(Color.Black)

                For r = 0 To 3
                    For c = 0 To 12

                        Dim x = sx + c * size
                        Dim y = sy + r * size
                        Dim rect As New Rectangle(x, y, size, size)

                        ' Always draw the cell border
                        g.DrawRectangle(p, rect)

                        If letterdg.Rows(r).Cells(c).Value Is Nothing Then
                            g.FillRectangle(Brushes.White, rect)
                        Else
                            g.DrawString(
                            letterdg.Rows(r).Cells(c).Value.ToString(),
                            f,
                            Brushes.Black,
                            x + 10,
                            y + 6
                        )
                        End If
                    Next
                Next
            End Using

        Catch ex As Exception
            MessageBox.Show("An error occurred while printing the letter grid: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub


    Private Sub PrintGrid(e As PrintPageEventArgs, printLetters As Boolean)

        Try
            Dim g = e.Graphics
            Dim size = 46
            Dim sx = 50
            Dim sy = 50
            Dim letterFont As New Font("Segoe UI", 12, FontStyle.Bold)
            Dim numberFont As New Font("Segoe UI", 8, FontStyle.Bold)
            Dim count As Integer = 0

            Using p As New Pen(Color.Black)

                For r = 0 To GridSize - 1
                    For c = 0 To GridSize - 1

                        Dim x = sx + c * size
                        Dim y = sy + r * size
                        Dim rect As New Rectangle(x, y, size, size)
                        'Dim text As String

                        If Puzzle = "cWord" Then
                            If printLetters Then
                                Dim num As String = DgvGrid(c, r).Value
                                If num <> "" Then g.DrawString(num, numberFont, Brushes.Black, x + 2, y + 1)
                            Else
                                g.DrawString(Grid(r, c).ToString(), letterFont, Brushes.Black, x + 6, y + 5)
                            End If

                            If Grid(r, c) = "."c Then
                                g.FillRectangle(Brushes.Black, rect)
                            Else
                                g.DrawRectangle(p, rect)
                            End If

                        Else ' Its a crossword or a PhraseWord.  just print  clue numbers

                            If Grid(r, c) = "."c Then
                                g.FillRectangle(Brushes.Black, rect)
                            Else
                                g.DrawRectangle(p, rect)
                            End If

                            If printLetters Then ' Print the answer sheet
                                Dim Letter As String = DgvGrid(c, r).Value
                                If Letter <> "" Then g.DrawString(Letter, letterFont, Brushes.Black, x + 12, y + 12)
                            End If

                            Dim ClueNumber As Integer = 0

                            For i = 0 To lstClues.Count - 1
                                Dim row = lstClues(i).Row
                                Dim col = lstClues(i).Col
                                Dim across = lstClues(i).IsAcross
                                Dim placeLetter = lstClues(i).PlaceLetter
                                ClueNumber += 1
                                If row = r And col = c Then
                                    If across Then
                                        g.DrawString(ClueNumber.ToString(), numberFont, Brushes.Black, x + 30, y + 2)
                                        count += 1
                                        If placeLetter Then ' if this clue is selected to have its letter placed on the crossword, print it in the cell
                                            Dim Letter As String = DgvGrid(c, r).Value
                                            g.DrawString(Letter, letterFont, Brushes.Red, x + 12, y + 12)
                                        End If
                                    Else
                                        g.DrawString(ClueNumber.ToString(), numberFont, Brushes.Black, x, y + 30)
                                        count += 1
                                        If placeLetter Then ' if this clue is selected to have its letter placed on the crossword, print it in the cell
                                            Dim Letter As String = DgvGrid(c, r).Value
                                            g.DrawString(Letter, letterFont, Brushes.Red, x + 12, y + 12)
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    Next
                Next
            End Using
        Catch ex As Exception
            MessageBox.Show("An error occurred while printing the grid: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()
    End Sub
End Class

