
Imports System.IO
    Module Dictionary
        Public DefaultPath As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + Path.DirectorySeparatorChar + "Crosswords"
        Public WordFilePath As String = Path.Combine(DefaultPath, "WordList.csv")

        Class ClueEntry
            Public Word As String
            Public Clue As String
        End Class

    Public Sub createDictionary(Puzzle As String, fPath As String)

        Dim wordlist As List(Of ClueEntry)
        wordlist = New List(Of ClueEntry)
        Dim seen As New HashSet(Of String)


        If Puzzle = "xWord" OrElse Puzzle = "cWord" Then

            Dim Words As New List(Of ClueEntry) From {
    New ClueEntry With {.Word = "HOUSE", .Clue = "Place where people live"},
    New ClueEntry With {.Word = "THREAD", .Clue = "Used for sewing"},
    New ClueEntry With {.Word = "BREAD", .Clue = "Baked food made from flour"},
    New ClueEntry With {.Word = "LETTER", .Clue = "Character in the alphabet"},
    New ClueEntry With {.Word = "SYSTEM", .Clue = "Set of connected things"},
    New ClueEntry With {.Word = "ISLAND", .Clue = "Land surrounded by water"},
    New ClueEntry With {.Word = "DRUM", .Clue = "Percussion instrument"},
    New ClueEntry With {.Word = "MOTION", .Clue = "The act of moving"},
    New ClueEntry With {.Word = "FORCE", .Clue = "Push or a pull"},
    New ClueEntry With {.Word = "OBJECT", .Clue = "Visible or tangible thing"},
    New ClueEntry With {.Word = "DESERT", .Clue = "Dry, barren region"},
    New ClueEntry With {.Word = "POWER", .Clue = "Ability to do work"},
    New ClueEntry With {.Word = "FUNCTION", .Clue = "Specific purpose or role"},
    New ClueEntry With {.Word = "INDEX", .Clue = "Alphabetical list"},
    New ClueEntry With {.Word = "SPEED", .Clue = "Rate of movement"},
    New ClueEntry With {.Word = "BATTERY", .Clue = "Device that stores energy"},
    New ClueEntry With {.Word = "CHAIR", .Clue = "Furniture for sitting"},
    New ClueEntry With {.Word = "TRIANGLE", .Clue = "Three-sided shape"},
    New ClueEntry With {.Word = "ENGINE", .Clue = "Machine that produces power"},
    New ClueEntry With {.Word = "STREAM", .Clue = "Small flowing body of water"},
    New ClueEntry With {.Word = "GREEN", .Clue = "Colour"},
    New ClueEntry With {.Word = "COPPER", .Clue = "Reddish metal"},
    New ClueEntry With {.Word = "CLIENT", .Clue = "Customer"},
    New ClueEntry With {.Word = "IRON", .Clue = "Strong metal"},
    New ClueEntry With {.Word = "CIRCLE", .Clue = "Round shape"},
    New ClueEntry With {.Word = "GALAXY", .Clue = "System of stars"},
    New ClueEntry With {.Word = "JAVA", .Clue = "Programming language"},
    New ClueEntry With {.Word = "PYTHON", .Clue = "Programming language"},
    New ClueEntry With {.Word = "STORAGE", .Clue = "Place to keep things"},
    New ClueEntry With {.Word = "VALLEY", .Clue = "Low land between hills"},
       New ClueEntry With {.Word = "AN", .Clue = "Indefinite article"},
New ClueEntry With {.Word = "AS", .Clue = "In the role of"},
New ClueEntry With {.Word = "AT", .Clue = "In a place"},
New ClueEntry With {.Word = "BE", .Clue = "To exist"},
New ClueEntry With {.Word = "BY", .Clue = "Near"},
New ClueEntry With {.Word = "DO", .Clue = "Perform"},
New ClueEntry With {.Word = "GO", .Clue = "Move"},
New ClueEntry With {.Word = "HE", .Clue = "Male pronoun"},
New ClueEntry With {.Word = "IF", .Clue = "Conditional word"},
New ClueEntry With {.Word = "IN", .Clue = "Inside"},
New ClueEntry With {.Word = "IS", .Clue = "Third person of BE"},
New ClueEntry With {.Word = "IT", .Clue = "Neutral pronoun"},
New ClueEntry With {.Word = "ME", .Clue = "Object pronoun"},
New ClueEntry With {.Word = "MY", .Clue = "Possessive pronoun"},
New ClueEntry With {.Word = "NO", .Clue = "Negative reply"},
New ClueEntry With {.Word = "OF", .Clue = "Belonging to"},
New ClueEntry With {.Word = "ON", .Clue = "Resting upon"},
New ClueEntry With {.Word = "OR", .Clue = "Choice word"},
New ClueEntry With {.Word = "SO", .Clue = "Therefore"},
New ClueEntry With {.Word = "TO", .Clue = "Toward"},
New ClueEntry With {.Word = "UP", .Clue = "Opposite of down"},
New ClueEntry With {.Word = "US", .Clue = "Plural pronoun"},
New ClueEntry With {.Word = "WE", .Clue = "Plural subject pronoun"},
New ClueEntry With {.Word = "AND", .Clue = "Joining word"},
New ClueEntry With {.Word = "ARE", .Clue = "Plural of IS"},
New ClueEntry With {.Word = "FOR", .Clue = "In favor of"},
New ClueEntry With {.Word = "NOT", .Clue = "Negative word"},
New ClueEntry With {.Word = "THE", .Clue = "Definite article"},
New ClueEntry With {.Word = "YOU", .Clue = "Second person pronoun"},
New ClueEntry With {.Word = "SUN", .Clue = "Star at center of solar system"},
New ClueEntry With {.Word = "MOON", .Clue = "Earth's satellite"},
New ClueEntry With {.Word = "STAR", .Clue = "Luminous celestial body"},
New ClueEntry With {.Word = "TREE", .Clue = "Tall plant with trunk"},
New ClueEntry With {.Word = "ROCK", .Clue = "Hard stone"},
New ClueEntry With {.Word = "WIND", .Clue = "Moving air"},
New ClueEntry With {.Word = "FIRE", .Clue = "Produces heat and light"},
New ClueEntry With {.Word = "WATER", .Clue = "H2O"},
New ClueEntry With {.Word = "EARTH", .Clue = "Our planet"},
New ClueEntry With {.Word = "SAND", .Clue = "Tiny grains of rock"},
New ClueEntry With {.Word = "SNOW", .Clue = "Frozen rain"},
New ClueEntry With {.Word = "RAIN", .Clue = "Water from clouds"},
New ClueEntry With {.Word = "CLOUD", .Clue = "Visible vapor in sky"},
New ClueEntry With {.Word = "ROAD", .Clue = "Path for vehicles"},
New ClueEntry With {.Word = "HOUSE", .Clue = "Place to live"},
New ClueEntry With {.Word = "CHAIR", .Clue = "Seat with back"},
New ClueEntry With {.Word = "TABLE", .Clue = "Furniture with flat top"},
New ClueEntry With {.Word = "LIGHT", .Clue = "Opposite of dark"},
New ClueEntry With {.Word = "SOUND", .Clue = "What you hear"},
New ClueEntry With {.Word = "PLANT", .Clue = "Living organism in soil"},
New ClueEntry With {.Word = "RIVER", .Clue = "Large flowing water"},
New ClueEntry With {.Word = "MOUSE", .Clue = "Small rodent"},
New ClueEntry With {.Word = "HORSE", .Clue = "Riding animal"},
New ClueEntry With {.Word = "SHEEP", .Clue = "Animal with wool"},
New ClueEntry With {.Word = "SNAKE", .Clue = "Legless reptile"},
New ClueEntry With {.Word = "BRICK", .Clue = "Building block"},
New ClueEntry With {.Word = "METAL", .Clue = "Strong material"},
New ClueEntry With {.Word = "GLASS", .Clue = "Transparent material"},
New ClueEntry With {.Word = "CLOCK", .Clue = "Tells time"},
New ClueEntry With {.Word = "PHONE", .Clue = "Communication device"},
New ClueEntry With {.Word = "RADIO", .Clue = "Audio broadcast device"},
New ClueEntry With {.Word = "TRAIN", .Clue = "Rail transport"},
New ClueEntry With {.Word = "PLANE", .Clue = "Flying vehicle"},
New ClueEntry With {.Word = "TRUCK", .Clue = "Large road vehicle"},
New ClueEntry With {.Word = "COMPUTER", .Clue = "Electronic machine"},
New ClueEntry With {.Word = "PRINTER", .Clue = "Outputs paper copies"},
New ClueEntry With {.Word = "KEYBOARD", .Clue = "Typing device"},
New ClueEntry With {.Word = "MONITOR", .Clue = "Display screen"},
New ClueEntry With {.Word = "SOFTWARE", .Clue = "Programs and applications"},
New ClueEntry With {.Word = "HARDWARE", .Clue = "Physical computer parts"},
New ClueEntry With {.Word = "NETWORK", .Clue = "Connected system"},
New ClueEntry With {.Word = "PROGRAM", .Clue = "Set of instructions"},
New ClueEntry With {.Word = "LANGUAGE", .Clue = "System of communication"},
New ClueEntry With {.Word = "PUZZLE", .Clue = "Brain teaser"},
New ClueEntry With {.Word = "CROSSWORD", .Clue = "Word puzzle grid"},
New ClueEntry With {.Word = "MOUNTAIN", .Clue = "Very high hill"},
New ClueEntry With {.Word = "OCEAN", .Clue = "Large body of salt water"},
New ClueEntry With {.Word = "FOREST", .Clue = "Large area of trees"},
New ClueEntry With {.Word = "DESERT", .Clue = "Dry barren land"},
New ClueEntry With {.Word = "VALLEY", .Clue = "Low land between hills"},
New ClueEntry With {.Word = "ISLAND", .Clue = "Land surrounded by water"},
New ClueEntry With {.Word = "BATTERY", .Clue = "Stores electrical energy"},
New ClueEntry With {.Word = "ENGINE", .Clue = "Machine that produces power"},
New ClueEntry With {.Word = "SYSTEM", .Clue = "Set of connected parts"}
}
            wordlist.AddRange(Words)

        ElseIf Puzzle = "pWord" Then
            Dim Words As New List(Of ClueEntry) From {
                New ClueEntry With {.Word = "APPLE", .Clue = "A fruit"},
                New ClueEntry With {.Word = "BRAVE", .Clue = "Courageous"},
                New ClueEntry With {.Word = "CRANE", .Clue = "A type of bird"},
                New ClueEntry With {.Word = "DANCE", .Clue = "Move rhythmically"},
                New ClueEntry With {.Word = "EARTH", .Clue = "Our planet"},
                New ClueEntry With {.Word = "FLAME", .Clue = "Fire"},
                New ClueEntry With {.Word = "GRAPE", .Clue = "A small fruit"},
                New ClueEntry With {.Word = "HEART", .Clue = "Organ that pumps blood"},
                New ClueEntry With {.Word = "ISLAND", .Clue = "Land surrounded by water"},
                New ClueEntry With {.Word = "JOKER", .Clue = "A wild card"},
                New ClueEntry With {.Word = "KITE", .Clue = "A flying toy"},
                New ClueEntry With {.Word = "LEMON", .Clue = "A sour fruit"},
                New ClueEntry With {.Word = "MOUSE", .Clue = "A small rodent"},
                New ClueEntry With {.Word = "NIGHT", .Clue = "Opposite of day"},
                New ClueEntry With {.Word = "OCEAN", .Clue = "Large body of salt water"},
                New ClueEntry With {.Word = "PLANT", .Clue = "Living organism"},
                New ClueEntry With {.Word = "QUEEN", .Clue = "Female monarch"},
                New ClueEntry With {.Word = "RIVER", .Clue = "Flowing water"},
                New ClueEntry With {.Word = "SNAKE", .Clue = "Legless reptile"},
                New ClueEntry With {.Word = "TRAIN", .Clue = "Rail transport"},
                New ClueEntry With {.Word = "UNIVERSE", .Clue = "All of space"},
                New ClueEntry With {.Word = "VIOLET", .Clue = "A purple flower"},
                New ClueEntry With {.Word = "WHALE", .Clue = "Large marine mammal"},
                New ClueEntry With {.Word = "XENON", .Clue = "Noble gas"},
                New ClueEntry With {.Word = "YACHT", .Clue = "Luxury boat"},
                New ClueEntry With {.Word = "ZEBRA", .Clue = "Striped animal"},
               New ClueEntry With {.Word = "EUNUCH", .Clue = "Sterile Male"},
               New ClueEntry With {.Word = "GOOSING", .Clue = "Grab Ass"},
               New ClueEntry With {.Word = "IRISHLINEN", .Clue = "Good sheets"},
               New ClueEntry With {.Word = "MEDICATED", .Clue = "Dosed up"},
               New ClueEntry With {.Word = "SORBET", .Clue = "Pallet refresher"},
               New ClueEntry With {.Word = "CROWDFUNDER", .Clue = "Give a little"},
               New ClueEntry With {.Word = "COINMAGIC", .Clue = "Magicians do this"},
               New ClueEntry With {.Word = "EXTRACTION", .Clue = "Dentist's job"},
               New ClueEntry With {.Word = "ROYALFLUSH", .Clue = "Poker players dream"},
               New ClueEntry With {.Word = "SOUVENEER", .Clue = "Brought home from holiday"},
               New ClueEntry With {.Word = "ILLTRYREDIALING", .Clue = "If you get the wrong number"},
               New ClueEntry With {.Word = "WATERINGHOLE", .Clue = "Local Bar"},
               New ClueEntry With {.Word = "FALLLINE", .Clue = "A line straight down a slope"},
               New ClueEntry With {.Word = "BUSKINGTABLES", .Clue = "Casual Waiter"},
               New ClueEntry With {.Word = "CARPEDIEM", .Clue = "Seize the day"},
               New ClueEntry With {.Word = "THESNIP", .Clue = "Vasectomy"},
               New ClueEntry With {.Word = "METARZAN", .Clue = "YouJane"},
               New ClueEntry With {.Word = "SOWEIRD", .Clue = "Very unusual"},
               New ClueEntry With {.Word = "PADDYFIELD", .Clue = "Rice growing area"},
               New ClueEntry With {.Word = "FELONIOUSMONK", .Clue = "Rasputin"},
               New ClueEntry With {.Word = "RITETOBEARARMS", .Clue = "Second Amendment"},
               New ClueEntry With {.Word = "SAWNOFFSHOTGUN", .Clue = "Criminals weapon of choice"},
               New ClueEntry With {.Word = "FAWNOVER", .Clue = "Excessive admiration"},
               New ClueEntry With {.Word = "FORHINDQUARTERS", .Clue = "Go to the local Butcher"},
               New ClueEntry With {.Word = "ALLFORIT", .Clue = "I agree, Lets do it"},
               New ClueEntry With {.Word = "NINETEEN", .Clue = "Going on Twenty"},
               New ClueEntry With {.Word = "LICKOFSENSE", .Clue = "What unruly kids don't have"},
               New ClueEntry With {.Word = "BACONARTIST", .Clue = "Draws using Pigskin"},
               New ClueEntry With {.Word = "NIGHTSDREAM", .Clue = "Mid Summer ???"},
               New ClueEntry With {.Word = "HOUSEDETECTIVE", .Clue = "Hotel security agent"},
               New ClueEntry With {.Word = "DIVORCEE", .Clue = "Now flying solo"},
               New ClueEntry With {.Word = "IRAQWAR", .Clue = "Second Gulf war"},
               New ClueEntry With {.Word = "SANCTUMS", .Clue = "The seat of power"}
                           }

            wordlist.AddRange(Words)
        End If



        Using writer As New StreamWriter(fPath, False)
            writer.WriteLine("Word,Clue")
            For Each entry In wordlist
                writer.WriteLine($"{entry.Word},{entry.Clue}")
            Next

            writer.Close()
        End Using

    End Sub
End Module

