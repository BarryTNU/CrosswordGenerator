Imports System.IO
Imports System.Drawing.Printing
Imports System.Diagnostics.Eventing.Reader

Public Class DictionaryGenerator

    Public DefaultPath As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + Path.DirectorySeparatorChar + "Crosswords"
    Public WordFilePath As String = Path.Combine(DefaultPath, "WordList.csv")

#Region "DATA STRUCTURES"

    Public Class ClueEntry
        Public Word As String
        Public Clue As String

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
    End Class

    Public Word As String
    Public UniqueWord As Boolean = False
    Public PrintPageIndex As Integer = 0
    Public WordList As List(Of ClueEntry)
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



#End Region

#Region "FORM LOAD"
    Private Sub Dictionary_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = "VB.NET Dictionary Generator"
        Me.Size = New Size(490, 800)
        Me.Location = New Point(300, 50)
        SetupUI()
        LoadDictionary(WordFilePath)
    End Sub

    Private Sub SetupUI()

        '====================TEXT BOXES FOR NEW ENTRIES===================
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
        End With
        ' ===================== DICTIONARY LABLE=====================
        With lbl_Dictionary
            .Font = New Font("Segoe UI", 14, FontStyle.Bold)
            .Text = "Dictionary"
            .ForeColor = Color.Black
            .BackColor = Color.White
            .BorderStyle = BorderStyle.FixedSingle
            .AutoSize = True
            .Location = New Point(150, 30)
        End With

        Controls.AddRange({lbl_Dictionary, lv_Dictionary, txt_NewWords, txt_NewClues})
        AddHandler txt_NewWords.KeyDown, AddressOf Shared_KeyDown
        AddHandler txt_NewClues.KeyDown, AddressOf Shared_KeyDown
        AddHandler txt_NewWords.GotFocus, AddressOf txtBox_HasFocus
        AddHandler txt_NewClues.GotFocus, AddressOf txtBox_HasFocus

    End Sub

#End Region

#Region "Load and Save DICTIONARY"
    Sub LoadDictionary(path As String)

        Try
            WordList = New List(Of ClueEntry)

            For Each line In File.ReadAllLines(path).Skip(1)

                Dim p = line.Split(","c, 2)
                Dim word = If(p.Length > 1, p(0).Trim(), "")
                Dim Clue = If(p.Length > 1, p(1).Trim(), "")

                If Not WordList.Any(Function(w) w.Word = word) Then
                    WordList.Add(New ClueEntry With {.Word = word, .Clue = Clue})
                    Dim item As New ListViewItem(word)      ' Word Column 
                    item.SubItems.Add(Clue)       ' Clue Column 
                    lv_Dictionary.Items.Add(item)
                End If
            Next

        Catch ex As Exception
            MessageBox.Show("Failed to load Dictionary.")
        End Try
    End Sub


    Private Sub lv_Dictionary_MouseDoubleClick(sender As Object, e As MouseEventArgs) _
    Handles lv_Dictionary.MouseDoubleClick

        Dim hit = lv_Dictionary.HitTest(e.Location)

        If hit.Item Is Nothing Then Exit Sub

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
                    WordList = New List(Of ClueEntry)
                End If
            End If
        Else
            ' If double-clicked on the clue column, Edit the clue
            Word = hit.Item.Text
            txt_NewWords.Text = Word
            txt_NewClues.Text = hit.SubItem.Text
            lv_Dictionary.Items.Remove(hit.Item)
            txt_NewClues.Focus()
        End If

    End Sub


    Private Sub SaveDictionary(Wordfilepath As String, wlist As List(Of ClueEntry))
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

        If e.KeyCode = Keys.Enter OrElse e.KeyCode = Keys.Tab Then
            e.SuppressKeyPress = True
            txt_NewWords.Text.Trim().ToUpper()

            If tb Is txt_NewWords Then
                Word = txt_NewWords.Text.Trim().ToUpper()
                If Len(Word) < 2 OrElse Len(Word) > 10 Then
                    MessageBox.Show("Word must be 2–10 letters (A–Z only).")
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
                'Dim Word As String = txt_NewWords.Text.Trim().ToUpper()
                Dim Clue As String = txt_NewClues.Text
                If Len(Clue) < 5 OrElse Len(Clue) > 30 Then
                    MessageBox.Show("Clue must be 5-30 characters.")
                    txt_NewClues.Clear()
                    txt_NewClues.Focus()
                    Exit Sub
                End If
                Dim item As New ListViewItem(Word)      ' Word Column 
                item.SubItems.Add(Clue)                 ' Clue Column 
                lv_Dictionary.Items.Insert(0, item) 'Add the word and clue to the first row in the list
                WordList.Add(New ClueEntry With {.Word = Word, .Clue = Clue})
                SaveDictionary(WordFilePath, WordList)

                txt_NewWords.Clear()
                txt_NewClues.Clear()
                txt_NewWords.Focus()
            End If
        End If

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
            Form1.Show()
        End Sub

    End Class