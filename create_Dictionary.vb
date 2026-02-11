
Imports System.IO
    Module Dictionary
        Public DefaultPath As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + Path.DirectorySeparatorChar + "Crosswords"
        Public WordFilePath As String = Path.Combine(DefaultPath, "WordList.csv")

        Class ClueEntry
            Public Word As String
            Public Clue As String
        End Class

    Sub createDictionary()

        Dim wordlist As List(Of ClueEntry)
        wordlist = New List(Of ClueEntry)
        Dim seen As New HashSet(Of String)

        Dim Words As New List(Of ClueEntry) From {
    New ClueEntry With {.Word = "HOUSE", .Clue = "(5) A place where people live"},
    New ClueEntry With {.Word = "THREAD", .Clue = "(6) Used for sewing"},
    New ClueEntry With {.Word = "BREAD", .Clue = "(5) Baked food made from flour"},
    New ClueEntry With {.Word = "LETTER", .Clue = "(6) A character in the alphabet"},
    New ClueEntry With {.Word = "SYSTEM", .Clue = "(6) A set of connected things"},
    New ClueEntry With {.Word = "ISLAND", .Clue = "(6) Land surrounded by water"},
    New ClueEntry With {.Word = "DRUM", .Clue = "(4) A percussion instrument"},
    New ClueEntry With {.Word = "MOTION", .Clue = "(6) The act of moving"},
    New ClueEntry With {.Word = "FORCE", .Clue = "(5) A push or a pull"},
    New ClueEntry With {.Word = "OBJECT", .Clue = "(6) A visible or tangible thing"},
    New ClueEntry With {.Word = "DESERT", .Clue = "(6) A dry, barren region"},
    New ClueEntry With {.Word = "POWER", .Clue = "(5) Ability to do work"},
    New ClueEntry With {.Word = "FUNCTION", .Clue = "(8) A specific purpose or role"},
    New ClueEntry With {.Word = "INDEX", .Clue = "(5) An alphabetical list"},
    New ClueEntry With {.Word = "SPEED", .Clue = "(5) Rate of movement"},
    New ClueEntry With {.Word = "BATTERY", .Clue = "(7) A device that stores energy"},
    New ClueEntry With {.Word = "CHAIR", .Clue = "(5) Furniture for sitting"},
    New ClueEntry With {.Word = "TRIANGLE", .Clue = "(8) A three-sided shape"},
    New ClueEntry With {.Word = "ENGINE", .Clue = "(6) Machine that produces power"},
    New ClueEntry With {.Word = "STREAM", .Clue = "(6) A small flowing body of water"},
    New ClueEntry With {.Word = "GREEN", .Clue = "(5) A colour"},
    New ClueEntry With {.Word = "COPPER", .Clue = "(6) A reddish metal"},
    New ClueEntry With {.Word = "CLIENT", .Clue = "(6) A customer"},
    New ClueEntry With {.Word = "IRON", .Clue = "(4) A strong metal"},
    New ClueEntry With {.Word = "CIRCLE", .Clue = "(6) A round shape"},
    New ClueEntry With {.Word = "GALAXY", .Clue = "(6) A system of stars"},
    New ClueEntry With {.Word = "JAVA", .Clue = "(4) A programming language"},
    New ClueEntry With {.Word = "PYTHON", .Clue = "(6) A programming language"},
    New ClueEntry With {.Word = "STORAGE", .Clue = "(7) A place to keep things"},
    New ClueEntry With {.Word = "VALLEY", .Clue = "(6) Low land between hills"}
}

        wordlist.AddRange(Words)


        Using writer As New StreamWriter(WordFilePath, False)
            writer.WriteLine("Word,Clue")
            For Each entry In wordlist
                writer.WriteLine($"{entry.Word},{entry.Clue}")
            Next

            writer.Close()
        End Using


    End Sub
End Module

