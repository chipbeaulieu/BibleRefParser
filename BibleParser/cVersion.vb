Imports System.ComponentModel
Imports System.IO
Public Class cVersion
    Implements INotifyPropertyChanged

#Region "Public Events"
    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged
#End Region 'Public Events

    Private mLibrary As cLibrary
    Private mName As String
    Private mPath As String
    Private mInitialized As Boolean
    Private mVerses As Dictionary(Of String, String)
    Private mVerseCountOfChapters As Dictionary(Of String, Integer)
    Private mFileReadComplete As Boolean

    ''' <summary>
    '''  
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Property Library() As cLibrary
        Get
            Return mLibrary
        End Get
        Set(ByVal value As cLibrary)
            mLibrary = value
        End Set
    End Property
    ''' <summary>
    '''  
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Property Name() As String
        Get
            Return mName
        End Get
        Set(ByVal value As String)
            mName = value.trim
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Name"))
        End Set
    End Property
    ''' <summary>
    '''  
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Property Path() As String
        Get
            Return mPath
        End Get
        Set(ByVal value As String)
            mPath = value.trim
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Path"))
        End Set
    End Property
    ''' <summary>
    '''  
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Property Initialized() As Boolean
        Get
            Return mInitialized
        End Get
        Set(ByVal value As Boolean)
            mInitialized = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Initialized"))
        End Set
    End Property
    ''' <summary>
    '''  
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Property Verses() As Dictionary(Of String, String)
        Get
            Return mVerses
        End Get
        Set(ByVal value As Dictionary(Of String, String))
            mVerses = value
        End Set
    End Property

    ''' <summary>
    '''  
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property VerseCountOfChapters() As Dictionary(Of String, Integer)
        Get
            Return mVerseCountOfChapters
        End Get
    End Property
    ''' <summary>
    '''  
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property FileReadComplete() As Boolean
        Get
            Return mFileReadComplete
        End Get
    End Property

    Public Function IsValid(bookChapterVerseIndex As String) As Boolean
        If mVerses Is Nothing OrElse mVerses.Count = 0 Then
            Return True
        Else
            Return mVerses.ContainsKey(bookChapterVerseIndex)
        End If
    End Function
    Public Function GetPreviousVerseIndex(bn As Integer, cn As Integer, vn As Integer) As String
        Dim cv As Integer
        Dim done As Boolean

        On Error Resume Next
        vn = vn - 1
        If vn = 0 Then
            cn = cn - 1
            If cn = 0 Then
                bn = bn - 1
                If bn = 0 Then
                    'before beginning of bible
                Else
                    'get last chapter of new book
                    cn = mLibrary.BibleBooks(mLibrary.BibleBooks.Keys(bn - 1)).Chapters
                    vn = VerseCountOfChapter(bn, cn)
                End If
            Else
                vn = VerseCountOfChapter(bn, cn)
            End If
        End If
        Return mLibrary.BookChapterVerseIndex(bn, cn, vn)
    End Function
    Public Function GetNextVerseIndex(ByRef bn As Integer, ByRef cn As Integer, ByRef vn As Integer) As String
        Dim cv As Integer

        On Error Resume Next
        vn = vn + 1
        cv = VerseCountOfChapter(bn, cn)
        If cv > 0 Then
            If vn > cv Then
                vn = 1
                cn = cn + 1
                If cn > mLibrary.BibleBooks(mLibrary.BibleBooks.Keys(bn - 1)).Chapters Then
                    cn = 1
                    bn = bn + 1
                    If bn > 66 Then
                        'Err.Raise 20003, "cVersion.GetNextVerseIndex", "Verse exceeds Bible length error"
                    End If
                End If
            End If
        End If
        Return mLibrary.BookChapterVerseIndex(bn, cn, vn)
    End Function
    Public Function VerseCountOfChapter(BookNum As Integer, ChapterNum As Integer) As Integer
        On Error Resume Next
        Return mVerseCountOfChapters(Format(BookNum, "00") & Format(ChapterNum, "000"))
    End Function
    Public Sub ReadBible()
        Dim line As String
        Dim verseIndex As String
        Dim chapterIndex As String
        Dim lastChapterIndex As String = ""
        Dim lastVerseIndex As String = ""
        Dim verseText As String
        'Dim i As Long
        Dim msg As String

        mVerses = New Dictionary(Of String, String)
        mVerseCountOfChapters = New Dictionary(Of String, Integer)
        mFileReadComplete = False
        Try
            If FileIO.FileSystem.FileExists(mPath) Then
                Using ts As New StreamReader(mPath)
                    While Not ts.EndOfStream
                        line = ts.ReadLine
                        If line <> "" Then
                            verseIndex = line.Substring(0, 8)
                            chapterIndex = verseIndex.Substring(0, 5)
                            verseText = line.Substring(9)
                            mVerses.Add(verseIndex, verseText)

                            If chapterIndex <> lastChapterIndex Then
                                'see if this is first chapter or not
                                If lastChapterIndex <> "" Then
                                    'there is a prior chapter. Make sure I know the number of verses
                                    mVerseCountOfChapters.Add(lastChapterIndex, CInt(Right(lastVerseIndex, 3)))
                                End If
                            End If
                            lastVerseIndex = verseIndex
                            lastChapterIndex = chapterIndex
                        End If
                        'i = i + 1
                    End While

                    If lastChapterIndex <> "" Then
                        'add the last chapter
                        mVerseCountOfChapters.Add(lastChapterIndex, CInt(Right(lastVerseIndex, 3)))
                    End If
                    mFileReadComplete = True
                End Using
            Else
                mFileReadComplete = False
            End If
        Catch ex As Exception
            'Select Case Err.Number
            '    Case 457
            '        'Repeat verse number
            '        msg = "Repeat verse index at " & lastVerseIndex
            '    Case 2077
            '        msg = Err.Description
            '    Case Else
            '        msg = Err.Description & " at lastVerseIndex of " & lastVerseIndex
            'End Select
            msg = ex.Message & " at lastVerseIndex of " & lastVerseIndex
            'messagebox.show("Error reading Bible version: " & vbCrLf & vbCrLf & msg, vbOKOnly, "Error Reading Bible")
            Debug.Print("In cVersion.ReadBible: " & ex.Message & " at lastVerseIndex of " & lastVerseIndex)
            mFileReadComplete = False
        End Try
    End Sub
    ''' <summary>
    ''' Creates a default empty bible for use with parsing bible references w/o text
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CreateEmptyBible()
        Const chapterVerses As String = "001031002025003024004026005032006022007024008022009029010032011032012020013018014024015021016016017027018033019038020018021034022024023020024067025034026035027046028022029035030043031055032032033020034031035029036043037036038030039023040023041057042038043034044034045028046034047031048022049033050026" &
        "001022002025003022004031005023006030007025008032009035010029011010012051013022014031015027016036017016018027019025020026021036022031023033024018025040026037027021028043029046030038031018032035033023034035035035036038037029038031039043040038001017002016003017004035005019006030007038008036009024010020" &
        "011047012008013059014057015033016034017016018030019037020027021024022033023044024023025055026046027034001054002034003051004049005031006027007089008026009023010036011035012016013033014045015041016050017013018032019022020029021035022041023030024025025018026065027023028031029040030016031054032042033056" &
        "034029035034036013001046002037003029004049005033006025007026008020009029010022011032012032013018014029015023016022017020018022019021020020021023022030023025024022025019026019027026028068029029030020031030032052033029034012001018002024003017004024005015006027007026008035009027010043011023012024013033" &
        "014015015063016010017018018028019051020009021045022034023016024033001036002023003031004024005031006040007025008035009057010018011040012015013025014020015020016031017013018031019030020048021025001022002023003018004022001028002036003021004022005012006021007017008022009027010027011015012025013023014052" &
        "015035016023017058018030019024020042021015022023023029024022025044026025027012028025029011030031031013001027002032003039004012005025006023007029008018009013010019011027012031013039014033015037016023017029018033019043020026021022022051023039024025001053002046003028004034005018006038007051008066009028" &
        "010029011043012033013034014031015034016034017024018046019021020043021029022053001018002025003027004044005027006033007020008029009037010036011021012021013025014029015038016020017041018037019037020021021026022020023037024020025030001054002055003024004043005026006081007040008040009044010014011047012040" &
        "013014014017015029016043017027018017019019020008021030022019023032024031025031026032027034028021029030001017002018003017004022005014006042007022008018009031010019011023012016013022014015015019016014017019018034019011020037021020022012023021024027025028026023027009028027029036030027031021032033033025" &
        "034033035027036023001011002070003013004024005017006022007028008036009015010044001011002020003032004023005019006019007073008018009038010039011036012047013031001022002023003015004017005014006014007010008017009032010003001022002013003026004021005027006030007021008022009035010022011020012025013028014022" &
        "015035016022017016018021019029020029021034022030023017024025025006026014027023028028029025030031031040032022033033034037035016036033037024038041039030040024041034042017001006002012003008004008005012006010007017008009009020010018011007012008013006014007015005016011017015018050019014020009021013022031" &
        "023006024010025022026012027014028009029011030012031024032011033022034022035028036012037040038022039013040017041013042011043005044026045017046011047009048014049020050023051019052009053006054007055023056013057011058011059017060012061008062012063011064010065013066020067007068035069036070005071024072020" &
        "073028074023075010076012077020078072079013080019081016082008083018084012085013086017087007088018089052090017091016092015093005094023095011096013097012098009099009100005101008102028103022104035105045106048107043108013109031110007111010112010113009114008115018116019117002118029119176120007121008122009" &
        "123004124008125005126006127005128006129008130008131003132018133003134003135021136026137009138008139024140013141010142007143012144015145021146010147020148014149009150006001033002022003035004027005023006035007027008036009018010032011031012028013025014035015033016033017028018024019029020030021031022029" &
        "023035024034025028026028027027028028029027030033031031001018002026003022004016005020006012007029008017009018010020011010012014001017002017003011004016005016006013007013008014001031002022003026004006005030006013007025008022009021010034011016012006013022014032015009016014017014018007019025020006021017" &
        "022025023018024023025012026021027013028029029024030033031009032020033024034017035010036022037038038022039008040031041029042025043028044028045025046013047015048022049026050011051023052015053012054017055013056012057021058014059021060022061011062012063019064012065025066024001019002037003025004031005031" &
        "006030007034008022009026010025011023012017013027014022015021016021017027018023019015020018021014022030023040024010025038026024027022028017029032030024031040032044033026034022035019036032037021038028039018040016041018042022043013044030045005046028047007048047049039050046051064052034001022002022003066" &
        "004022005022001028002010003027004017005017006014007027008018009011010022011025012028013023014023015008016063017024018032019014020049021032022031023049024027025017026021027036028026029021030026031018032032033033034031035015036038037028038023039029040049041026042020043027044031045025046024047023048035" &
        "001021002049003030004037005031006028007028008027009027010021011045012013001011002023003005004019005015006011007016008014009017010015011012012014013016014009001020002032003021001015002016003015004013005027006014007017008014009015001021001017002010003010004011001016002013003012004013005015006016007020" &
        "001015002013003019001017002020003019001018002015003020001015002023001021002013003010004014005011006015007014008023009017010012011017012014013009014021001014002017003018004006001025002023003017004025005048006034007029008034009038010042011030012050013058014036015039016028017027018035019030020034021046" &
        "022046023039024051025046026075027066028020001045002028003035004041005043006056007037008038009050010052011033012044013037014072015047016020001080002052003038004044005039006049007050008056009062010042011054012059013035014035015032016031017037018043019048020047021038022071023056024053001051002025003036" &
        "004054005047006071007053008059009041010042011057012050013038014031015027016033017026018040019042020031021025001026002047003026004037005042006015007060008040009043010048011030012025013052014028015041016040017034018028019041020038021040022030023035024027025027026032027044028031001032002029003031004025" &
        "005021006023007025008039009033010021011036012021013014014023015033016027001031002016003023004021005013006020007040008013009027010033011034012031013013014040015058016024001024002017003018004018005021006018007016008024009015010018011033012021013014001024002021003029004031005026006018001023002022003021" &
        "004032005033006024001030002030003021004023001029002023003025004018001010002020003013004018005028001012002017003018001020002015003016004016005025006021001018002026003017004022001016002015003015001025001014002018003019004016005014006020007028008013009028010039011040012029013025001027002026003018004017" &
        "005020001025002025003022004019005014001021002022003018001010002029003024004021005021001013001014001025001020002029003022004011005014006017007017008013009021010011011019012017013018014020015008016021017018018024019021020015021027022021"

        Dim lastIndex As String = "999"
        Dim cvIndex As String
        Dim i As Integer, j As Integer
        Dim bookID As Integer = 0
        Dim chapterID As String
        Dim maxVerse As Integer
        Dim verseIndex As String
        Dim chapterIndex As String
        Dim lastChapterIndex As String = ""
        Dim lastVerseIndex As String = ""

        mVerses = New Dictionary(Of String, String)
        mVerseCountOfChapters = New Dictionary(Of String, Integer)
        mFileReadComplete = False

        For i = 0 To chapterVerses.Length - 1 Step 6
            'Debug.Print(i)
            cvIndex = chapterVerses.Substring(i, 6)
            chapterID = cvIndex.Substring(0, 3)

            If lastIndex >= chapterID Then
                'a new book
                bookID += 1
            End If

            maxVerse = cvIndex.Substring(3, 3)
            chapterIndex = String.Format("{0}{1}", bookID.ToString("00"), chapterID)

            For j = 1 To maxVerse
                verseIndex = String.Format("{0}{1}{2}", bookID.ToString("00"), chapterID, j.ToString("000"))
                mVerses.Add(verseIndex, "")

                If chapterIndex <> lastChapterIndex Then
                    'see if this is first chapter or not
                    If lastChapterIndex <> "" Then
                        'there is a prior chapter. Make sure I know the number of verses
                        mVerseCountOfChapters.Add(lastChapterIndex, CInt(Right(lastVerseIndex, 3)))
                    End If
                End If
                lastVerseIndex = verseIndex
                lastChapterIndex = chapterIndex
            Next
            lastIndex = chapterID
        Next
        If lastChapterIndex <> "" Then
            'add the last chapter
            mVerseCountOfChapters.Add(lastChapterIndex, CInt(Right(lastVerseIndex, 3)))
        End If
        mFileReadComplete = True
    End Sub
    Public Sub New()
        mVerses = New Dictionary(Of String, String)
        mVerseCountOfChapters = New Dictionary(Of String, Integer)
    End Sub
End Class
