
Imports System.IO

Module ClueDatabase
    Public DefaultPath As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + Path.DirectorySeparatorChar + "Crosswords"
    Public WordFilePath As String = Path.Combine(DefaultPath, "Word_List.csv")
    Public dB1 As String = Path.Combine(DefaultPath, "ClueDB1.txt")
    Public dB2 As String = Path.Combine(DefaultPath, "ClueDB2.txt")

    Public Class Clue
        Public Word As String
        Public Row As Integer
        Public Col As Integer
        Public IsAcross As Boolean
        Public Clue As String
        Public ClueNumber As Integer
        Public PlaceLetter As Boolean = False ' Used in crossword puzzles to indicate whether the letter should be revealed on the grid (for codeword puzzles, all letters are hidden)
    End Class

    ' Public Class ClueEntry
    Public Word As String
    ''     Public Clue As String
    ' End Class

    ' Public Enum DirectionType
    '    Across
    '     Down
    ' End Enum


    '  Public Const Down As Boolean = False
    '  Public Puzzle As String = ""
    '  Public PrintPageIndex As Integer = 0
    '  Public WordList As List(Of String)
    '  Public PlacedWords As New List(Of String)
    '  Public Condensed As Boolean = False
    '  Dim rnd As New Random()

    ' Public Class LetterNumber
    'Public Letter As String
    'Public Number As Integer
    ' End Class

    ' Public Class CellData
    ' Public Letter As String
    '  Public AcrossNumber As Integer
    ' Public DownNumber As Integer
    'End Class
    '
    Public Sub LoadClueDatabase()

        If Not File.Exists(dB1) Then
            File.Create(dB1).Close()
        End If
        If Not File.Exists(dB2) Then
            File.Create(dB2).Close()
        End If
        LoadClueEntries()

    End Sub

    Public Sub LoadClueEntries()

        Dim clueEntries As New List(Of ClueEntry)
        Using reader As New StreamReader(dB1)
            While Not reader.EndOfStream
                Dim line As String = reader.ReadLine()
                Dim parts As String() = line.Split(";"c)
                If parts.Length >= 2 Then
                    Dim entry As New ClueEntry With {
                        .Word = parts(0).Trim(),
                        .Clue = parts(1).Trim()
                    }
                    clueEntries.Add(entry)
                End If
            End While
        End Using

        SaveClueEntries(clueEntries)
        CurateClueEntries()
    End Sub

    Public Sub CurateClueEntries()
        Dim clueEntries As New List(Of ClueEntry)
        Dim Word As String
        Using reader As New StreamReader(WordFilePath)
            While Not reader.EndOfStream
                Word = reader.ReadLine()
                Dim parts As String() = Word.Split(","c)
                If parts(0).Length >= 5 AndAlso parts(0).Length <= 15 Then
                    Dim entry As New ClueEntry With {
                        .Word = parts(0).Trim(),
                        .Clue = parts(1).Trim()
                    }
                    clueEntries.Add(entry)
                End If

            End While
        End Using
        SaveClueEntries(clueEntries)
    End Sub
    Public Sub SaveClueEntries(clueEntries As List(Of ClueEntry))

        Using writer As New StreamWriter(WordFilePath, False)
            For Each entry As ClueEntry In clueEntries
                writer.WriteLine($"{entry.Word},{entry.Clue}")
            Next
        End Using
    End Sub
End Module
