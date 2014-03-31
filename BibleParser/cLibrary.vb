Imports System.ComponentModel
Imports System.IO
Public Class cLibrary
    Implements INotifyPropertyChanged

#Region "Public Events"
    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged
#End Region 'Public Events
    Private mLibraryPath As String
    Private mVersions As Dictionary(Of String, cVersion)
    Private mBibleBooks As Dictionary(Of String, cBibleBook)
    Private mCurrentVersion As cVersion
    Private mBibleReference As cBibleReference

    Public Function IsBookNameValid(BookName As String) As Boolean
        Dim result As Boolean = False

        For Each b In BibleBooks
            result = result Or b.Value.HasName(BookName)
        Next

        Return result
    End Function
    ''' <summary>
    '''  
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Property LibraryPath() As String
        Get
            Return mLibraryPath
        End Get
        Set(ByVal value As String)
            Dim v As cVersion

            If value = mLibraryPath Then
                Exit Property
            Else
                mLibraryPath = value
            End If

            FileIO.FileSystem.FileExists(value)
            If FileIO.FileSystem.FileExists(value) Then
                'get the list of files in this path
                Dim fileList As List(Of String) = New List(Of String)(System.IO.Directory.EnumerateFiles(value))
                Dim f As String

                For Each f In fileList
                    v = New cVersion
                    If Len(Path.GetExtension(f)) > 0 Then
                        Path.GetFileNameWithoutExtension(f)
                        v.Name = Path.GetFileNameWithoutExtension(f)
                    Else
                        v.Name = Path.GetFileName(f)
                    End If
                    v.Path = f
                    v.Library = Me

                    mVersions.Add(v.Name, v)
                Next
            End If

            'set default version
            If mVersions.ContainsKey("KJV") Then
                CurrentVersion = mVersions("KJV")
            ElseIf mVersions.Count > 0 Then
                CurrentVersion = mVersions(0)
            Else
                CurrentVersion = Nothing
            End If

            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("LibraryPath"))
        End Set
    End Property
    ''' <summary>
    '''  
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Property Versions() As Dictionary(Of String, cVersion)
        Get
            Return mVersions
        End Get
        Set(ByVal value As Dictionary(Of String, cVersion))
            mVersions = value
        End Set
    End Property
    ''' <summary>
    '''  
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property BibleBooks() As Dictionary(Of String, cBibleBook)
        Get
            Return mBibleBooks
        End Get
    End Property
    ''' <summary>
    '''  
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Property CurrentVersion() As cVersion
        Get
            Return mCurrentVersion
        End Get
        Set(ByVal value As cVersion)
            mCurrentVersion = value
            If Not mCurrentVersion.FileReadComplete Then
                mCurrentVersion.ReadBible()
            End If
        End Set
    End Property
    ''' <summary>
    '''  
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property BibleReference() As cBibleReference
        Get
            Return mBibleReference
        End Get
    End Property

    Public Function ShortBibleReference(BookChapterVerseIndex As String) As String
        Dim bn As Integer
        Dim cn As Integer
        Dim vn As Integer

        '01001001 - Genesis 1:1
        bn = BookChapterVerseIndex.Substring(0, 2)
        cn = BookChapterVerseIndex.Substring(2, 3)
        vn = BookChapterVerseIndex.Substring(Math.Max(0, BookChapterVerseIndex.Length - 3))'Right(BookChapterVerseIndex, 3)

        ShortBibleReference = Me.BibleBooks(BibleBooks.Keys(bn - 1)).ShortName & " " & cn & ":" & vn
    End Function
    Public Function GetBookNumber(BookName As String) As Integer
        Dim foundNum As Integer = 0

        For Each b In mBibleBooks
            If b.Value.HasName(BookName) Then
                foundNum = b.Value.Number
                Exit For
            End If
        Next

        Return foundNum
    End Function
    Public Function GetBookShortName(BookName As String) As String
        Dim foundName As String = ""

        For Each b In mBibleBooks
            If b.Value.HasName(BookName) Then
                foundName = b.Value.ShortName
                Exit For
            End If
        Next

        Return foundName
    End Function
    Public Function BookChapterVerseIndex(BookNumber As Integer, ChapterNumber As Integer, VerseNumber As Integer) As String
        BookChapterVerseIndex = Format(BookNumber, "00") & Format(ChapterNumber, "000") & Format(VerseNumber, "000")
    End Function
    Public Sub New()
        'Dim delim As String
        Dim nb As cBibleBook
        mVersions = New Dictionary(Of String, cVersion)
        mBibleBooks = New Dictionary(Of String, cBibleBook)
        mBibleReference = New cBibleReference

        mBibleReference.Library = Me

        'load up information about bible books

        '1. GENESIS
        nb = New cBibleBook
        nb.LongName = "Genesis"
        nb.ShortName = "Gen"
        nb.Abbreviations = "_gn_"
        nb.Chapters = 50
        nb.Number = 1
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '2. EXODUS
        nb = New cBibleBook
        nb.LongName = "Exodus"
        nb.ShortName = "Ex"
        nb.Abbreviations = "_exd_"
        nb.Chapters = 40
        nb.Number = 2
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '3. LEVITICUS
        nb = New cBibleBook
        nb.LongName = "Leviticus"
        nb.ShortName = "Lev"
        nb.Abbreviations = "_lv_"
        nb.Chapters = 27
        nb.Number = 3
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '4. NUMBERS
        nb = New cBibleBook
        nb.LongName = "Numbers"
        nb.ShortName = "Num"
        nb.Abbreviations = "_nm_nb_"
        nb.Chapters = 36
        nb.Number = 4
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '5. DEUTERONOMY
        nb = New cBibleBook
        nb.LongName = "Deuteronomy"
        nb.ShortName = "Deut"
        nb.Abbreviations = "_dt_"
        nb.Chapters = 34
        nb.Number = 5
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '6. JOSHUA
        nb = New cBibleBook
        nb.LongName = "Joshua"
        nb.ShortName = "Jos"
        nb.Abbreviations = "_jsh_josue_"
        nb.Chapters = 24
        nb.Number = 6
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '7. JUDGES
        nb = New cBibleBook
        nb.LongName = "Judges"
        nb.ShortName = "Jud"
        nb.Abbreviations = "_jdg_jg_jdgs_jug_"
        nb.Chapters = 21
        nb.Number = 7
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '8. RUTH
        nb = New cBibleBook
        nb.LongName = "Ruth"
        nb.ShortName = "Ru"
        nb.Abbreviations = "_rth_"
        nb.Chapters = 4
        nb.Number = 8
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '9. 1 SAMUEL
        nb = New cBibleBook
        nb.LongName = "1 Samuel"
        nb.ShortName = "1 Sam"
        nb.Abbreviations = "_1 sm_1sm_ism_1st samuel_first samuel_isam_1 sam_i sam_"
        nb.Chapters = 31
        nb.Number = 9
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '10. 2 SAMUEL
        nb = New cBibleBook
        nb.LongName = "2 Samuel"
        nb.ShortName = "2 Sam"
        nb.Abbreviations = "_2 sm_2sm_iism_2nd samuel_second samuel_2 sam_ii sam_"
        nb.Chapters = 24
        nb.Number = 10
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '11. 1 KINGS
        nb = New cBibleBook
        nb.LongName = "1 Kings"
        nb.ShortName = "1 Kgs"
        nb.Abbreviations = "_1 kgs_1kgs_1stkgs_1stkings_firstkings_firstkgs_ikgs_1 king_i king_"
        nb.Chapters = 22
        nb.Number = 11
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '12. 2 KINGS
        nb = New cBibleBook
        nb.LongName = "2 Kings"
        nb.ShortName = "2 Kgs"
        nb.Abbreviations = "_2 kgs_2kgs_2ndkgs_2ndkings_secondkings_secondkgs_iikgs_2 king_ii king_"
        nb.Chapters = 25
        nb.Number = 12
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '13. 1 CHRONICLES
        nb = New cBibleBook
        nb.LongName = "1 Chronicles"
        nb.ShortName = "1 Chr"
        nb.Abbreviations = "_1stchronicles_firstchronicles_1 paralipomenon_1 chron_i chron_"
        nb.Chapters = 29
        nb.Number = 13
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '14. 2 CHRONICLES
        nb = New cBibleBook
        nb.LongName = "2 Chronicles"
        nb.ShortName = "2 Chr"
        nb.Abbreviations = "_2ndchronicles_secondchronicles_2 paralipomenon_2 chron_ii chron_"
        nb.Chapters = 36
        nb.Number = 14
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '15. EZRA
        nb = New cBibleBook
        nb.LongName = "Ezra"
        nb.ShortName = "Ezr"
        nb.Abbreviations = "_1 esdras_"
        nb.Chapters = 10
        nb.Number = 15
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '16. NEHEMIAH
        nb = New cBibleBook
        nb.LongName = "Nehemiah"
        nb.ShortName = "Neh"
        nb.Abbreviations = "_2 esdras_"
        nb.Chapters = 13
        nb.Number = 16
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '17. ESTHER
        nb = New cBibleBook
        nb.LongName = "Esther"
        nb.ShortName = "Est"
        nb.Abbreviations = "_est_"
        nb.Chapters = 10
        nb.Number = 17
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '18. JOB
        nb = New cBibleBook
        nb.LongName = "Job"
        nb.ShortName = "Job"
        nb.Abbreviations = "_jb_"
        nb.Chapters = 42
        nb.Number = 18
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '19. PSALMS
        nb = New cBibleBook
        nb.LongName = "Psalms"
        nb.ShortName = "Ps"
        nb.Abbreviations = "_pslm_ps_psm_pss_"
        nb.Chapters = 150
        nb.Number = 19
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '20. PROVERBS
        nb = New cBibleBook
        nb.LongName = "Proverbs"
        nb.ShortName = "Prov"
        nb.Abbreviations = "_prv_"
        nb.Chapters = 31
        nb.Number = 20
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '21. ECCLESIASTES
        nb = New cBibleBook
        nb.LongName = "Ecclesiastes"
        nb.ShortName = "Ecc"
        nb.Abbreviations = "_qoh_qoheleth_"
        nb.Chapters = 12
        nb.Number = 21
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '22. SONG OF SOLOMON
        nb = New cBibleBook
        nb.LongName = "Song of Solomon"
        nb.ShortName = "Song"
        nb.Abbreviations = "_song_so_canticleofcanticles_canticles_song of songs_sos_"
        nb.Chapters = 8
        nb.Number = 22
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '23. ISAIAH
        nb = New cBibleBook
        nb.LongName = "Isaiah"
        nb.ShortName = "Isa"
        nb.Abbreviations = "_isa_"
        nb.Chapters = 66
        nb.Number = 23
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '24. JEREMIAH
        nb = New cBibleBook
        nb.LongName = "Jeremiah"
        nb.ShortName = "Jer"
        nb.Abbreviations = "_jr_"
        nb.Chapters = 52
        nb.Number = 24
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '25. LAMENTATIONS
        nb = New cBibleBook
        nb.LongName = "Lamentations"
        nb.ShortName = "Lam"
        nb.Abbreviations = "_lamentations of jeremiah_"
        nb.Chapters = 5
        nb.Number = 25
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '26. EZEKIEL
        nb = New cBibleBook
        nb.LongName = "Ezekiel"
        nb.ShortName = "Eze"
        nb.Abbreviations = "_ezk_"
        nb.Chapters = 48
        nb.Number = 26
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '27. DANIEL
        nb = New cBibleBook
        nb.LongName = "Daniel"
        nb.ShortName = "Dan"
        nb.Abbreviations = "_dn_"
        nb.Chapters = 12
        nb.Number = 27
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '28. HOSEA
        nb = New cBibleBook
        nb.LongName = "Hosea"
        nb.ShortName = "Hos"
        nb.Abbreviations = "_ho_"
        nb.Chapters = 14
        nb.Number = 28
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '29. JOEL
        nb = New cBibleBook
        nb.LongName = "Joel"
        nb.ShortName = "Joel"
        nb.Abbreviations = "_jl_"
        nb.Chapters = 3
        nb.Number = 29
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '30. AMOS
        nb = New cBibleBook
        nb.LongName = "Amos"
        nb.ShortName = "Amos"
        nb.Abbreviations = "_am_"
        nb.Chapters = 9
        nb.Number = 30
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '31. OBADIAH
        nb = New cBibleBook
        nb.LongName = "Obadiah"
        nb.ShortName = "Oba"
        nb.Abbreviations = "_oba_"
        nb.Chapters = 1
        nb.Number = 31
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '32. JONAH
        nb = New cBibleBook
        nb.LongName = "Jonah"
        nb.ShortName = "Jon"
        nb.Abbreviations = "_jnh_"
        nb.Chapters = 4
        nb.Number = 32
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '33. MICAH
        nb = New cBibleBook
        nb.LongName = "Micah"
        nb.ShortName = "Mic"
        nb.Abbreviations = "_mic_"
        nb.Chapters = 7
        nb.Number = 33
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '34. NAHUM
        nb = New cBibleBook
        nb.LongName = "Nahum"
        nb.ShortName = "Nah"
        nb.Abbreviations = "_na_"
        nb.Chapters = 3
        nb.Number = 34
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '35. HABAKKUK
        nb = New cBibleBook
        nb.LongName = "Habakkuk"
        nb.ShortName = "Hab"
        nb.Abbreviations = "_hab_"
        nb.Chapters = 3
        nb.Number = 35
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '36. ZEPHANIAH
        nb = New cBibleBook
        nb.LongName = "Zephaniah"
        nb.ShortName = "Zep"
        nb.Abbreviations = "_zp_"
        nb.Chapters = 3
        nb.Number = 36
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '37. HAGGAI
        nb = New cBibleBook
        nb.LongName = "Haggai"
        nb.ShortName = "Hag"
        nb.Abbreviations = "_hg_"
        nb.Chapters = 2
        nb.Number = 37
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '38. ZECHARIAH
        nb = New cBibleBook
        nb.LongName = "Zechariah"
        nb.ShortName = "Zec"
        nb.Abbreviations = "_zc_"
        nb.Chapters = 14
        nb.Number = 38
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '39. MALACHI
        nb = New cBibleBook
        nb.LongName = "Malachi"
        nb.ShortName = "Mal"
        nb.Abbreviations = "_ml_"
        nb.Chapters = 4
        nb.Number = 39
        nb.Testament = cBibleBook.enTestament.OldTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '40. MATTHEW
        nb = New cBibleBook
        nb.LongName = "Matthew"
        nb.ShortName = "Mt"
        nb.Abbreviations = "_mt_"
        nb.Chapters = 28
        nb.Number = 40
        nb.Testament = cBibleBook.enTestament.NewTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '41. MARK
        nb = New cBibleBook
        nb.LongName = "Mark"
        nb.ShortName = "Mk"
        nb.Abbreviations = "_mrk_mk_mr_mak_"
        nb.Chapters = 16
        nb.Number = 41
        nb.Testament = cBibleBook.enTestament.NewTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '42. LUKE
        nb = New cBibleBook
        nb.LongName = "Luke"
        nb.ShortName = "Lk"
        nb.Abbreviations = "_lk_"
        nb.Chapters = 24
        nb.Number = 42
        nb.Testament = cBibleBook.enTestament.NewTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '43. JOHN
        nb = New cBibleBook
        nb.LongName = "John"
        nb.ShortName = "Jn"
        nb.Abbreviations = "_jn_jhn_"
        nb.Chapters = 21
        nb.Number = 43
        nb.Testament = cBibleBook.enTestament.NewTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '44. ACTS
        nb = New cBibleBook
        nb.LongName = "Acts"
        nb.ShortName = "Acts"
        nb.Abbreviations = "_acts of the apostles_"
        nb.Chapters = 28
        nb.Number = 44
        nb.Testament = cBibleBook.enTestament.NewTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '45. ROMANS
        nb = New cBibleBook
        nb.LongName = "Romans"
        nb.ShortName = "Rom"
        nb.Abbreviations = "_rm_"
        nb.Chapters = 16
        nb.Number = 45
        nb.Testament = cBibleBook.enTestament.NewTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '46. 1 CORINTHIANS
        nb = New cBibleBook
        nb.LongName = "1 Corinthians"
        nb.ShortName = "1 Cor"
        nb.Abbreviations = "_1stcorinthians_first corinthians_1 cor_i cor_"
        nb.Chapters = 16
        nb.Number = 46
        nb.Testament = cBibleBook.enTestament.NewTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '47. 2 CORINTHIANS
        nb = New cBibleBook
        nb.LongName = "2 Corinthians"
        nb.ShortName = "2 Cor"
        nb.Abbreviations = "_2ndcorinthians_second corinthians_2 cor_ii cor_"
        nb.Chapters = 13
        nb.Number = 47
        nb.Testament = cBibleBook.enTestament.NewTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '48. GALATIANS
        nb = New cBibleBook
        nb.LongName = "Galatians"
        nb.ShortName = "Gal"
        nb.Abbreviations = "_ga_"
        nb.Chapters = 6
        nb.Number = 48
        nb.Testament = cBibleBook.enTestament.NewTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '49. EPHESIANS
        nb = New cBibleBook
        nb.LongName = "Ephesians"
        nb.ShortName = "Eph"
        nb.Abbreviations = "_eph_"
        nb.Chapters = 6
        nb.Number = 49
        nb.Testament = cBibleBook.enTestament.NewTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '50. PHILIPPIANS
        nb = New cBibleBook
        nb.LongName = "Philippians"
        nb.ShortName = "Phil"
        nb.Abbreviations = "_php_phl_"
        nb.Chapters = 4
        nb.Number = 50
        nb.Testament = cBibleBook.enTestament.NewTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '51. COLOSSIANS
        nb = New cBibleBook
        nb.LongName = "Colossians"
        nb.ShortName = "Col"
        nb.Abbreviations = "_col_"
        nb.Chapters = 4
        nb.Number = 51
        nb.Testament = cBibleBook.enTestament.NewTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '52. 1 THESSALONIANS
        nb = New cBibleBook
        nb.LongName = "1 Thessalonians"
        nb.ShortName = "1 Thes"
        nb.Abbreviations = "1stthessalonians_first thessalonians_1ts_1 ts_1 the_i the_"
        nb.Chapters = 5
        nb.Number = 52
        nb.Testament = cBibleBook.enTestament.NewTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '53. 2 THESSALONIANS
        nb = New cBibleBook
        nb.LongName = "2 Thessalonians"
        nb.ShortName = "2 Thes"
        nb.Abbreviations = "_2ndthessalonians_second thessalonians_2ts_2 ts_2 the_ii the_"
        nb.Chapters = 3
        nb.Number = 53
        nb.Testament = cBibleBook.enTestament.NewTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '54. 1 TIMOTHY
        nb = New cBibleBook
        nb.LongName = "1 Timothy"
        nb.ShortName = "1 Tim"
        nb.Abbreviations = "_1sttimothy_first timothy_1 tm_i tm_1tm_1 tim_i tim_"
        nb.Chapters = 6
        nb.Number = 54
        nb.Testament = cBibleBook.enTestament.NewTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '55. 2 TIMOTHY
        nb = New cBibleBook
        nb.LongName = "2 Timothy"
        nb.ShortName = "2 Tim"
        nb.Abbreviations = "_2ndtimothy_second timothy_2 tm_ii tm_2tm_2 tim_ii tim_"
        nb.Chapters = 4
        nb.Number = 55
        nb.Testament = cBibleBook.enTestament.NewTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '56. TITUS
        nb = New cBibleBook
        nb.LongName = "Titus"
        nb.ShortName = "Tit"
        nb.Abbreviations = "_tit_"
        nb.Chapters = 3
        nb.Number = 56
        nb.Testament = cBibleBook.enTestament.NewTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '57. PHILEMON
        nb = New cBibleBook
        nb.LongName = "Philemon"
        nb.ShortName = "Phile"
        nb.Abbreviations = "_phm_"
        nb.Chapters = 1
        nb.Number = 57
        nb.Testament = cBibleBook.enTestament.NewTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '58. HEBREWS
        nb = New cBibleBook
        nb.LongName = "Hebrews"
        nb.ShortName = "Heb"
        nb.Abbreviations = "_heb_"
        nb.Chapters = 13
        nb.Number = 58
        nb.Testament = cBibleBook.enTestament.NewTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '59. JAMES
        nb = New cBibleBook
        nb.LongName = "James"
        nb.ShortName = "Jam"
        nb.Abbreviations = "_jas_jm_"
        nb.Chapters = 5
        nb.Number = 59
        nb.Testament = cBibleBook.enTestament.NewTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '60. 1 PETER
        nb = New cBibleBook
        nb.LongName = "1 Peter"
        nb.ShortName = "1 Pet"
        nb.Abbreviations = "_ipt_1pt_1stpeter_first peter_1 pet_i pet_"
        nb.Chapters = 5
        nb.Number = 60
        nb.Testament = cBibleBook.enTestament.NewTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '61. 2 PETER
        nb = New cBibleBook
        nb.LongName = "2 Peter"
        nb.ShortName = "2 Pet"
        nb.Abbreviations = "_iipt_2pt_2ndpeter_second peter_2 pet_ii pet_"
        nb.Chapters = 3
        nb.Number = 61
        nb.Testament = cBibleBook.enTestament.NewTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '62. 1 JOHN
        nb = New cBibleBook
        nb.LongName = "1 John"
        nb.ShortName = "1 Jn"
        nb.Abbreviations = "_1 jhn_1 jn_1jhn_1jn_i jhn_i jn_ijhn_ijh_1stjohn_first john_ijn_1 joh_i joh_"
        nb.Chapters = 5
        nb.Number = 62
        nb.Testament = cBibleBook.enTestament.NewTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '63. 2 JOHN
        nb = New cBibleBook
        nb.LongName = "2 John"
        nb.ShortName = "2 Jn"
        nb.Abbreviations = "_2 jhn_2 jn_2jhn_2jn_ii jhn_ii jn_iijhn_iijh_2ndjohn_second john_iijn_2 joh_ii joh_"
        nb.Chapters = 1
        nb.Number = 63
        nb.Testament = cBibleBook.enTestament.NewTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '64. 3 JOHN
        nb = New cBibleBook
        nb.LongName = "3 John"
        nb.ShortName = "3 Jn"
        nb.Abbreviations = "_3 jhn_3 jn_3jhn_3jn_iii jhn_iii jn_iiijhn_iiijh_3rdjohn_third john_iiijn_3 joh_iii joh_"
        nb.Chapters = 1
        nb.Number = 64
        nb.Testament = cBibleBook.enTestament.NewTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '65. JUDE
        nb = New cBibleBook
        nb.LongName = "Jude"
        nb.ShortName = "Jude"
        nb.Abbreviations = "_jud_"
        nb.Chapters = 1
        nb.Number = 65
        nb.Testament = cBibleBook.enTestament.NewTestament
        mBibleBooks.Add(nb.ShortName, nb)

        '66. REVELATION
        nb = New cBibleBook
        nb.LongName = "Revelation"
        nb.ShortName = "Rev"
        nb.Abbreviations = "_the revelation_"
        nb.Chapters = 22
        nb.Number = 66
        nb.Testament = cBibleBook.enTestament.NewTestament
        mBibleBooks.Add(nb.ShortName, nb)

        'Read the generic bible 
        Dim gb As New cVersion

        gb.CreateEmptyBible()
        gb.Name = "Generic"
        gb.Library = Me
        Me.Versions.Add("GENERIC", gb)
        Me.CurrentVersion = gb
    End Sub


End Class
