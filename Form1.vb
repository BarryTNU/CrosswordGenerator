Imports System.IO
Imports System.Drawing.Printing
Imports System.Configuration
Imports System.Diagnostics

Public Class Form1

    Public DefaultPath As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + Path.DirectorySeparatorChar + "Crosswords"
    Public WordFilePath As String = Path.Combine(DefaultPath, "WordList.csv")

#Region "DATA STRUCTURES"

    Friend WithEvents DgvGrid As DataGridView
    Const GridSize As Integer = 15
    Dim Grid(GridSize - 1, GridSize - 1) As Char
    Dim lstClues As New List(Of Clue)
    Dim Dictionary As New List(Of ClueEntry)
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

    Public Class ClueEntry
        Public Word As String
        Public Clue As String
    End Class


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
    Private btnPuzzle As New Button
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


#End Region

    ' ===================== UI SETUP =====================

#Region "FORM LOAD"

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = "VB.NET Crossword"
        Me.Size = New Size(850, 850)
        Me.Location = New Point(100, 50)

        If My.Application.CommandLineArgs.Count > 0 Then
            Puzzle = My.Application.CommandLineArgs(0)
        End If

        ' Create directory for word list if it doesn't exist
        If Not Directory.Exists(DefaultPath) Then
            Directory.CreateDirectory(DefaultPath)
        End If

        'Check if there is a dictionary. If not, Create one from the included word list
        If Not File.Exists(WordFilePath) Then
            createDictionary() ' Ensure dictionary exists
        End If

        If Puzzle = "" Then Puzzle = "cWord" ' Default to codeword if not set

        LoadDictionary(WordFilePath)
        SetupGrid()
        SetupUI()
        If Puzzle = "cWord" Then
            SetupLetterGrid()
            SetupTextBox()
        End If
        GeneratePuzzle()

    End Sub

    Private Sub RestartApp(puzzleType As String)

        Dim exePath As String = Application.ExecutablePath
        Dim psi As New ProcessStartInfo(exePath, puzzleType)
        Process.Start(psi)
        Application.Exit()
    End Sub


#End Region
#Region "DICTIONARY"
    Sub LoadDictionary(path As String) ' Load the word list from CSV file

            Dictionary.Clear()

            For Each line In File.ReadAllLines(path).Skip(1)
                Dim p = line.Split(","c, 2)
                Dictionary.Add(New ClueEntry With {.Word = p(0).Trim().ToUpper(), .Clue = p(1).Trim().ToUpper()})
            Next

        End Sub

#End Region

#Region "GRID SETUP"

    Private Sub SetupGrid()

        ' ---------- Grid panel ----------
        pnlGrid = New Panel With {
        .Dock = DockStyle.Left,
        .Width = 755
    }

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
            r.Height = 50
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

    End Sub


    Private Sub SetupUI()

        ' ---------- Bottom panel ----------
        pnlBottom = New Panel With {
        .Dock = DockStyle.Bottom,
        .Height = 70
    }

        ' ---------- Edit Dictionary ----------
        btnEditDict.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        btnEditDict.Text = "Edit Dictionary"
        btnEditDict.AutoSize = True
        btnEditDict.Location = New Point(10, 20)

        '----------Puzzle type label----------
        btnPuzzle.Font = New Font("Segoe UI", 12, FontStyle.Bold)

        btnPuzzle.Text = "CodeWord"
        btnPuzzle.AutoSize = True
        btnPuzzle.Location = New Point(160, 20)
        pnlBottom.Controls.Add(btnPuzzle)
        If Puzzle = "cWord" Then btnPuzzle.Text = "Crossword"
        If Puzzle = "xWord" Then btnPuzzle.Text = "CodeWord"


        ' ---------- Number of clues ----------
        lblNrClues.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        lblNrClues.Text = "Number of Clues:"
        lblNrClues.AutoSize = True
        lblNrClues.Location = New Point(300, 25)

        cmbNrClues.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        cmbNrClues.Items.AddRange({"0", "1", "2", "3", "4", "5"})
        cmbNrClues.SelectedIndex = 3
        cmbNrClues.DropDownStyle = ComboBoxStyle.DropDownList
        cmbNrClues.Width = 60
        cmbNrClues.Location = New Point(460, 22)

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
        pnlBottom.Controls.AddRange({
        btnEditDict,
        lblNrClues,
        cmbNrClues,
        btnNew,
        btnPrint
    })

        Me.Controls.Add(pnlBottom)

        ' ---------- Crossword clue panels ----------
        If Puzzle = "xWord" Then

            pnlClues = New Panel With {
            .Dock = DockStyle.Right,
            .Width = 260,
            .Height = 400
        }

            lstAcross.Font = New Font("Segoe UI", 10)
            lstAcross.Dock = DockStyle.Top
            lstAcross.Width = pnlClues.Width
            lstAcross.ScrollAlwaysVisible = False
            lstAcross.Items.Add("Across Clues:")

            lstDown.Font = New Font("Segoe UI", 10)
            lstDown.Width = pnlClues.Width
            lstDown.ScrollAlwaysVisible = False
            lstDown.Location = New Point(0, pnlClues.Height \ 2)
            lstDown.Items.Add("Down Clues:")

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
        AddHandler btnPuzzle.Click, AddressOf btnPuzzle_Click
        AddHandler cmbNrClues.SelectedIndexChanged, AddressOf cmbNrClues_IndexChanged
        AddHandler DgvGrid.CellPainting, AddressOf dgvGrid_CellPainting

    End Sub


    Sub SetupLetterGrid() ' Setup the letter-number mapping grid

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

        End Sub

    Sub SetupTextBox() ' Setup the alphabet textbox
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
    End Sub

#End Region

#Region "CROSSWORD ENGINE"

    Sub GeneratePuzzle() ' This is the main puzzle generation routine

        InitGrid()

        ' Shuffle dictionary
        Dim shuffled = Dictionary.OrderBy(Function(x) rnd.Next()).ToList()

        ' Take 60, then sort those by length
        Dim Words = shuffled.Take(60).OrderByDescending(Function(x) x.Word.Length).ToList()

        ' Place longest of the random set in the middle of the grid, horizontally
        PlaceWord(Words(0), GridSize \ 2, (GridSize - Words(0).Word.Length) \ 2, True)

        ' Place remaining words
        For i = 1 To Words.Count - 1
            If WordUsed(Words(i).Word) Then Continue For
            TryPlaceIntersect(Words(i))
        Next

        RenderGrid()

    End Sub


    Sub InitGrid() ' Initialize the grid and UI for a new puzzle
            PuzzleNumber = rnd.Next(1, 1000)

        lstClues.Clear()
        lstAcross.Items.Clear()
        lstAcross.Items.Add("Across Clues:")
        lstDown.Items.Clear()
        lstDown.Items.Add("Down Clues:")

        Me.Text = "VB.NET Codeword" & " - Puzzle #" & PuzzleNumber.ToString()

        For r = 0 To GridSize - 1 'Start with all empty cells
                For c = 0 To GridSize - 1
                    Grid(r, c) = "."c 'cell will be rendered as black
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
            lstClues.Clear()
            PlacedWords.Clear()
        End Sub

        Function WordUsed(Word As String) 'Check if word has already been placed
            For Each item In PlacedWords
                If Word = item.ToString() Then 'word has been used,
                    Return True
                End If
            Next
            Return False
        End Function
        Function TryPlaceIntersect(Clue As ClueEntry) As Boolean ' Try to place a word by intersecting with existing letters

            For r = 1 To GridSize - 1
                For c = 0 To GridSize - 1
                    If Grid(r, c) = "."c Then Continue For

                    For i = 0 To Clue.Word.Length - 1
                        If Clue.Word(i) = Grid(r, c) Then

                            If CanPlace(Clue.Word, r, c - i, True) Then 'Try horizontal first
                                PlaceWord(Clue, r, c - i, True)
                                Return True
                            End If

                            If CanPlace(Clue.Word, r - i, c, False) Then 'Then try vertical
                                PlaceWord(Clue, r - i, c, False)
                                Return True
                            End If

                        End If
                    Next
                Next
            Next

            Return False

        End Function

        Function CanPlace(word As String, row As Integer, col As Integer, horiz As Boolean) As Boolean ' Check if a word can be placed at the specified location
            Dim r, c As Integer

            ' Reject negative starts

            If row < 0 OrElse col < 0 Then Return False

            If horiz AndAlso col + word.Length > GridSize Then Return False
            If Not horiz AndAlso row + word.Length > GridSize Then Return False

            For i = 0 To word.Length - 1
                r = If(horiz, row, row + i)
                c = If(horiz, col + i, col)

                If Grid(r, c) <> "."c AndAlso Grid(r, c) <> word(i) Then
                    Return False
                End If
            Next

            Return True
        End Function

        Function PlaceWord(entry As ClueEntry, row As Integer, col As Integer, horizontal As Boolean) As Boolean ' Place a word on the grid

            Dim Word = entry.Word
            Dim Clue = entry.Clue

            ' Bounds check
            If horizontal Then
                If col < 0 OrElse col + Word.Length > GridSize Then Return False
            Else
                If row < 0 OrElse row + Word.Length > GridSize Then Return False
            End If

            ' Collision check
            For i = 0 To Word.Length - 1
                Dim r = If(horizontal, row, row + i)
                Dim c = If(horizontal, col + i, col)

                If Grid(r, c) <> "."c AndAlso Grid(r, c) <> Word(i) Then
                    Return False
                End If
            Next

            ' Place letters
            For i = 0 To Word.Length - 1
                Dim r = If(horizontal, row, row + i)
                Dim c = If(horizontal, col + i, col)
                Dim cVal As String = Word(i)
                Grid(r, c) = cVal
            Next

            ' Force black cell AFTER word
            If horizontal Then
                Dim afterCol = col + Word.Length
                If afterCol < GridSize AndAlso Grid(row, afterCol) = "."c Then
                    Grid(row, afterCol) = "."c
                End If
            Else
                Dim afterRow = row + Word.Length
                If afterRow < GridSize AndAlso Grid(afterRow, col) = "."c Then
                    Grid(afterRow, col) = "."c
                End If
            End If

        If Puzzle = "xWord" Then
            lstClues.Add(New Clue With {.Word = Word, .Row = row, .Col = col, .IsAcross = horizontal, .Clue = entry.Clue})
            If horizontal Then
                lstAcross.Items.Add($"{lstClues.Count }. {entry.Clue}")
            Else
                lstDown.Items.Add($"{lstClues.Count }. {entry.Clue}")
            End If

            lstAcross.Height = lstAcross.Items.Count * 20 + 30 ' Adjust height based on number of clues
            lstDown.Height = lstDown.Items.Count * 20 + 30 ' Adjust height based on number of clues
            lstDown.Location = New Point(0, pnlClues.Height \ 2) ' Keep down clues in the bottom half of the panel
        End If
        ' Record placed word
        PlacedWords.Add(Word)
        Return True

    End Function

#End Region

#Region "RENDER"

    Sub RenderGrid() ' Render the grid to the DataGridView

        Dim ShowSolution As Boolean = True 'If true shows letters in cells. Set to false for blank grid.
        'ShowSolution = False ' Set to false for codeword puzzles, true for crosswords

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


        If Puzzle = "xWord" Then
            Dim ClueNumber As Integer = 0
            For i = 0 To lstClues.Count - 1
                Dim row = lstClues(i).Row
                Dim col = lstClues(i).Col
                Dim across = lstClues(i).IsAcross
                ClueNumber += 1

                If across Then
                    cD(row, col).AcrossNumber = ClueNumber
                    lstClues(i).ClueNumber = ClueNumber
                Else
                    cD(row, col).DownNumber = ClueNumber
                    lstClues(i).ClueNumber = ClueNumber
                End If

            Next
        Else
            AssignNumbersToLetters()
        End If

    End Sub

    Private Sub dgvGrid_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs) 'Handles DgvGrid.CellPainting

        If e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Return
        Dim Row = e.RowIndex
        Dim Col = e.ColumnIndex
        ' Let grid paint background & borders
        e.Paint(e.CellBounds, DataGridViewPaintParts.All)

        Using f As New Font("Segoe UI", 7, FontStyle.Regular),
          b As New SolidBrush(Color.Black)

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
    End Sub


#End Region

#Region "NUMBER ASSIGNMENT"
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

    Function GenerateRandomNumber(n As Integer) ' Example of a random number generator function
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

    Sub AssignClues(n As Integer)
        ' This routine assigns the clues to the grid based on the random numbers generated for the crossword puzzle
        Dim x As Integer = LstRnr.Count
        For i = 0 To n - 1
            Dim rnd As New Random()
            Dim q As Integer = rnd.Next(1, x + 1)
            Dim clue = lstClues(q)
            clue.PlaceLetter = True
        Next

    End Sub

#End Region

#Region "Button Handlers"
    ' ===================== BUTTONS =====================
    Sub EditDictionary(sender As Object, e As EventArgs)
        Me.Hide()
        DictionaryGenerator.Show()
    End Sub
    Sub btnPuzzle_Click(sender As Object, e As EventArgs)
        If Puzzle = "cWord" Then
            RestartApp("xWord")
        Else
            RestartApp("cWord")
        End If
    End Sub
    Private Sub BtnNew_Click(sender As Object, e As EventArgs)
            GeneratePuzzle()
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

            If Puzzle = "xWord" Then
                pd.DefaultPageSettings.Landscape = True
                Dim Number As Integer = (lstClues.Count)
                LstRnr = GenerateRandomNumber(Number)
                AssignClues(NrOfClues)
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

        Dim n As New Random(15)


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
    Sub PrintClueLists(e As PrintPageEventArgs)
        'This routine prints the clues pages. Not used in codeword puzzles
        If Puzzle <> "xWord" Then Return

        Dim g = e.Graphics
        Dim font As New Font("Segoe UI", 9)
        Dim y = 20
        Dim ClueNumber As Integer = 0

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

    End Sub
    Private Sub PrintPuzzlePage(e As PrintPageEventArgs)

        Dim g = e.Graphics
        g.DrawString("CODEWORD PUZZLE" & " - Puzzle #" & PuzzleNumber.ToString(),
                 New Font("Segoe UI", 16, FontStyle.Bold),
                 Brushes.Black, 50, 20)

        PrintGrid(e, printLetters:=False)

    End Sub
    Private Sub PrintAnswerKeyPage(e As PrintPageEventArgs)

        If Puzzle <> "cWord" Then Return

        'pd.DefaultPageSettings.Landscape = False

        Dim g = e.Graphics
        g.DrawString("ANSWER KEY" & " - Puzzle #" & PuzzleNumber.ToString(),
                 New Font("Segoe UI", 16, FontStyle.Bold),
                 Brushes.Black, 50, 20)

        PrintGrid(e, printLetters:=True)
    End Sub

    Private Sub PrintxWordAnswers(e As PrintPageEventArgs)

        Dim g = e.Graphics
        g.DrawString("CROSS WORD PUZZLE" & " - Puzzle #" & PuzzleNumber.ToString(),
                 New Font("Segoe UI", 16, FontStyle.Bold),
                 Brushes.Black, 50, 20)

        PrintGrid(e, printLetters:=True)

    End Sub

    Private Sub PrintPlacedWords(e As PrintPageEventArgs)

            If PrintWords = False Then Return
        ' pd.DefaultPageSettings.Landscape = False

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
        End Sub

    Private Sub PrintAlphabetGrid(e As PrintPageEventArgs)

        If Puzzle <> "cWord" Then Return
        ' pd.DefaultPageSettings.Landscape = False

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

    End Sub
    Private Sub PrintLetterGrid(e As PrintPageEventArgs)
        If Puzzle <> "cWord" Then Return

        ' pd.DefaultPageSettings.Landscape = False

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

        End Sub


    Private Sub PrintGrid(e As PrintPageEventArgs, printLetters As Boolean)

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

                    Else ' Its a crossword. For crossword, just print  clue numbers

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

    End Sub

#End Region

End Class

