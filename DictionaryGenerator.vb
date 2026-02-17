Imports System.IO
Imports System.Drawing.Printing
Imports System.Diagnostics.Eventing.Reader

Public Class DictionaryGenerator

    Public DictFilePath As String
    ' Public dictionary As New Dictionary Of Clue

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


    Public Word As String
    Public UniqueWord As Boolean = False
    Public PrintPageIndex As Integer = 0
    Public WordList As New List(Of Clue)
    Dim rnd As New Random()
    Public WithEvents pd As New PrintDocument
    Private currentWord As String = ""
    Private CurrentRow As Integer = 0
    Private CurrentCol As Integer = 0

    ' ===================== UI =====================
    Private lbl_Dictionary As New Label
    Private lbl_New As New Label
    Private btnUpdate As New Button
    '====================== LISTBOXES =====================
    Public WithEvents lv_Dictionary As New ListView
    Public txt_NewWords As New TextBox
    Public txt_NewClues As New TextBox
    Public Puzzle As String = Form1.Puzzle



#End Region

#Region "FORM LOAD"
    Private Sub Dictionary_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        DictFilePath = Form1.DictFilePath

        Me.Text = "VB.NET Dictionary Generator"
        Me.Size = New Size(490, 800)
        Me.Location = New Point(300, 50)
        SetupUI()
        LoadFile(sender, e)

    End Sub

    Private Sub SetupUI()

        '====================TEXT BOXES FOR NEW ENTRIES===================

        Try
            With txt_NewWords
                .Font = New Font("Segoe UI", 12, FontStyle.Bold)
                .CharacterCasing = CharacterCasing.Upper
                .Location = New Point(50, 70)
                .Width = 130
                .Text = "    New words."
                .Select()
                .Focus()
            End With

            With txt_NewClues
                .Font = New Font("Segoe UI", 12, FontStyle.Bold)
                .Location = New Point(180, 70)
                .Width = 230
                .Text = "            New clues."
            End With

            '===================== DICTIONARY LIST BOX  =====================
            '===================== DICTIONARY LABLE=====================

            With lbl_Dictionary
                .Font = New Font("Segoe UI", 14, FontStyle.Bold)
                .Text = "Dictionary"
                .ForeColor = Color.Black
                .BackColor = Color.White
                .BorderStyle = BorderStyle.FixedSingle
                .TextAlign = ContentAlignment.MiddleCenter
                .Size = New Size(200, 40)
                '.AutoSize = True
                .Location = New Point(120, 20)
            End With

            With lv_Dictionary
                .Font = New Font("Segoe UI", 12)
                .Location = New Point(50, 100)
                .Size = New Size(375, 600)
                .View = View.Details
                .FullRowSelect = True
                .GridLines = True
                .Columns.Clear()
                .Columns.Add("Word", 130, HorizontalAlignment.Left)
                .Columns.Add("Clue", 230, HorizontalAlignment.Left)

                If Puzzle = "pWord" Then
                    Me.Size = New Size(610, 800)
                    .Size = New Size(520, 600)
                    .Columns(0).Width = 200
                    .Columns(1).Width = 300
                    txt_NewWords.Location = New Point(50, 70)
                    txt_NewWords.Width = 200
                    txt_NewClues.Location = New Point(250, 70)
                    txt_NewClues.Width = 300
                    lbl_Dictionary.Location = New Point(160, 20)
                End If
            End With

            Controls.AddRange({lbl_Dictionary, lv_Dictionary, txt_NewWords, txt_NewClues})
            AddHandler txt_NewWords.KeyDown, AddressOf Shared_KeyDown
            AddHandler txt_NewClues.KeyDown, AddressOf Shared_KeyDown
            AddHandler txt_NewWords.GotFocus, AddressOf txtBox_HasFocus
            AddHandler txt_NewClues.GotFocus, AddressOf txtBox_HasFocus

        Catch ex As Exception
            MessageBox.Show("Error setting up the UI.")
        End Try

    End Sub

#End Region

#Region "Load and Save DICTIONARY"

    Private Async Sub LoadFile(sender As Object, e As EventArgs)
        '
        Dim loadingForm As New FrmLoading()

        Cursor = Cursors.WaitCursor
        Try
            With loadingForm
                .StartPosition = FormStartPosition.CenterParent
            End With

            loadingForm.Show(Me)

            Dim progress = New Progress(Of Integer)(Sub(p) loadingForm.UpdateProgress(p))
            Dim items = Await Task.Run(Function() CreateItemsWithProgress(DictFilePath, progress))
            Dim result = Await Task.Run(Function() CreateItemsWithProgress(DictFilePath, progress))
            WordList = result.wlist
            lv_Dictionary.BeginUpdate()
            lv_Dictionary.Items.AddRange(result.items.ToArray())
            lv_Dictionary.EndUpdate()

            loadingForm.Close()
            Cursor = Cursors.Default
        Catch ex As Exception
            MessageBox.Show("Failed to load Dictionary.")
            loadingForm.Close()
            Cursor = Cursors.Default
        End Try

    End Sub


    Private Function CreateItemsWithProgress(path As String, progress As IProgress(Of Integer)) As (items As List(Of ListViewItem), wlist As List(Of Clue))
        Dim items As New List(Of ListViewItem)
        Dim wlist As New List(Of Clue)
        Try
            Using sr As New StreamReader(path)
                Dim totalBytes = sr.BaseStream.Length
                While Not sr.EndOfStream
                    Dim line = sr.ReadLine()
                    Dim p = line.Split(","c, 2)
                    Dim w = If(p.Length > 1, p(0).Trim(), "")
                    Dim c = If(p.Length > 1, p(1).Trim(), "")
                    If w = "Word" AndAlso c = "Clue" Then Continue While
                    Dim item As New ListViewItem(w)
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
        Catch ex As Exception
            MessageBox.Show("Failed to load Dictionary.")
        End Try

        Return (items, wlist)
    End Function

    Public Sub SaveDictionary(Wordfilepath As String, wlist As List(Of Clue))
        Try
            Using writer As New StreamWriter(Wordfilepath, False)
                writer.WriteLine("Word,Clue")
                For Each entry In wlist
                    writer.WriteLine($"{entry.Word},{entry.Clue}")
                Next
                writer.Close()
            End Using
        Catch ex As Exception
            MessageBox.Show("Failed to save Dictionary.")
        End Try

    End Sub

#End Region


#Region "ADD WORDS TO DICTIONARY"

    Private Sub lv_Dictionary_MouseDoubleClick(sender As Object, e As MouseEventArgs) _
    Handles lv_Dictionary.MouseDoubleClick

        Dim hit = lv_Dictionary.HitTest(e.Location)

        If hit.Item Is Nothing Then Exit Sub
        Try
            '===============If double click in column(0) then delete word and clue================'
            If hit.SubItem Is hit.Item.SubItems(0) Then
                Dim word As String = hit.Item.Text
                Dim clue As String = hit.SubItem.Text

                Dim response = MsgBox("Are you sure you want to delete " & word, MsgBoxStyle.YesNo, "Confirm Deletion")
                If response = vbYes Then

                    hit = lv_Dictionary.HitTest(e.Location)
                    If hit.Item Is Nothing Then Exit Sub

                    ' Only allow removal when double-clicking the word column
                    If hit.SubItem Is hit.Item.SubItems(0) Then
                        lv_Dictionary.Items.Remove(hit.Item)
                        'And rebuild the word list
                        WordList.RemoveAll(Function(c) c.Word = word AndAlso c.Clue = clue)
                        SaveDictionary(DictFilePath, WordList)
                    End If
                End If
            Else
                ' If double-clicked on the clue column, Edit the clue
                Word = hit.Item.Text
                txt_NewWords.Text = Word
                txt_NewClues.Text = hit.SubItem.Text
                lv_Dictionary.Items.Remove(hit.Item)
                WordList.RemoveAll(Function(c) c.Word = Word AndAlso c.Clue = hit.SubItem.Text)
                txt_NewClues.Focus()
            End If
        Catch ex As Exception
            MessageBox.Show("Error processing the selected item.")
        End Try

    End Sub

    Private Function CheckForDuplicates(Word As String)

        Try
            Word = Word.Trim().ToUpper()

            If Word = "" Then Return False

            For Each item In lv_Dictionary.Items
                If Word = item.ToString() Then
                    MessageBox.Show("Duplicate word.")
                    Return False
                End If
            Next

        Catch ex As Exception
            MessageBox.Show("Invalid word.")
            Exit Function
        End Try

        Return True

    End Function

    Private Sub Shared_KeyDown(sender As Object, e As KeyEventArgs)

        Dim tb = DirectCast(sender, TextBox)
        Try
            If e.KeyCode = Keys.Enter OrElse e.KeyCode = Keys.Tab Then
                e.SuppressKeyPress = True
                txt_NewWords.Text.Trim().ToUpper()

                If tb Is txt_NewWords Then
                    Word = txt_NewWords.Text.Trim().ToUpper()
                    If Len(Word) < 2 OrElse Len(Word) > 15 Then
                        MessageBox.Show("Word must be 2–15 letters (A–Z only).")
                        txt_NewWords.Clear()
                        txt_NewWords.Focus()
                        Exit Sub
                    End If
                    UniqueWord = CheckForDuplicates(Word)
                    If UniqueWord Then

                        txt_NewClues.Focus()
                    Else
                        txt_NewWords.Clear()
                        txt_NewWords.Focus()
                    End If

                ElseIf tb Is txt_NewClues Then
                    Dim Clue As String = txt_NewClues.Text

                    ' SaveDictionary(DictFilePath, WordList)

                    Word = txt_NewWords.Text.Trim().ToUpper()

                    If Len(Clue) < 5 OrElse Len(Clue) > 30 Then
                        MessageBox.Show("Clue must be 5-30 characters.")
                        txt_NewClues.Clear()
                        txt_NewClues.Focus()
                        Exit Sub
                    End If
                    Dim item As New ListViewItem(Word)      ' Word Column 
                    item.SubItems.Add(Clue)                 ' Clue Column 
                    lv_Dictionary.Items.Insert(0, item) 'Add the word and clue to the first row in the list
                    WordList.Add(New Clue With {.Word = Word, .Clue = Clue})
                    SaveDictionary(DictFilePath, WordList)

                    txt_NewWords.Clear()
                    txt_NewClues.Clear()
                    txt_NewWords.Focus()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Error processing input.")
        End Try

    End Sub

    Sub txtBox_HasFocus(sender As Object, e As EventArgs)

        Dim tb = DirectCast(sender, TextBox)
        If tb Is txt_NewWords Then
            txt_NewWords.Clear()
        ElseIf tb Is txt_NewClues Then
            txt_NewClues.Clear()
        End If
    End Sub

#End Region
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        Form1.Show()
    End Sub

End Class