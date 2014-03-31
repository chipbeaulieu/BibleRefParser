Imports System.ComponentModel
Imports System.Text

Public Class cBibleReference
    Implements INotifyPropertyChanged

#Region "Public Events"
    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged
#End Region 'Public Events

    Private mReferenceTokens As List(Of cReferenceToken)
    Private mLibrary As cLibrary
    Private mInvalidReference As Boolean
    Private mVerses As List(Of cVerse)
    Private mBadReferences As String
    Private mShortReference As String
    Private mLongReference As String

    Public ReadOnly Property VerseList As String
        Get
            Dim result As String = ""
            Dim sb As New StringBuilder

            For Each v In mVerses
                sb.Append(v.ShortReference)
                sb.Append(",")
            Next
            result = sb.ToString
            If result <> "" Then
                result = result.Substring(0, result.Length - 1)
            End If

            Return result
        End Get
    End Property
    ''' <summary>
    '''  
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Property ReferenceTokens() As List(Of cReferenceToken)
        Get
            Return mReferenceTokens
        End Get
        Set(ByVal value As List(Of cReferenceToken))
            mReferenceTokens = value
        End Set
    End Property
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
    Public ReadOnly Property InvalidReference() As Boolean
        Get
            Return mInvalidReference
        End Get
    End Property
    ''' <summary>
    '''  
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Verses() As List(Of cVerse)
        Get
            Return mVerses
        End Get
    End Property
    ''' <summary>
    '''  
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property BadReferences() As String
        Get
            Return mBadReferences
        End Get
    End Property
    ''' <summary>
    '''  
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ShortReference() As String
        Get
            Return mShortReference
        End Get
    End Property
    ''' <summary>
    '''  
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property LongReference() As String
        Get
            Return mLongReference
        End Get
    End Property

    Public ReadOnly Property versesHaveNoQuotes() As Boolean
        Get
            Dim v As cVerse
            Dim noVerses As Boolean

            noVerses = True
            For Each v In mVerses
                noVerses = noVerses And InStr(v.VerseText, Chr(34)) = 0
            Next
            Return noVerses
        End Get
    End Property

    Public Sub ParseReference(InValue As String, findFirst2VersesOnly As Boolean)
        'given a string, parse it out
        Dim ref As String
        Dim i As Integer
        Dim r As String
        Dim prevToken As cReferenceToken = Nothing
        Dim curToken As cReferenceToken
        Dim nextToken As cReferenceToken
        Dim maxTokens As Integer
        Dim prevBookToken As cReferenceToken = Nothing, prevChapterToken As cReferenceToken = Nothing, prevVerseToken As cReferenceToken = Nothing
        Dim done As Boolean
        Dim addFullChapter As Boolean
        Dim addFullBook As Boolean
        Dim rangeDelimToken As New cReferenceToken
        Dim verse1Token As New cReferenceToken
        Dim chapter1Token As New cReferenceToken


        Try
            ref = PreparseReference(InValue)

            rangeDelimToken.ReferencePart = "-"
            rangeDelimToken.TokenType = cReferenceToken.BibleRefTokenType.RangeDelimiter
            verse1Token.ReferencePart = 1
            verse1Token.TokenType = cReferenceToken.BibleRefTokenType.verse
            chapter1Token.ReferencePart = 1
            chapter1Token.TokenType = cReferenceToken.BibleRefTokenType.Chapter

            mReferenceTokens = New List(Of cReferenceToken)

            'at this point, I think I have a reference.
            curToken = New cReferenceToken
            For i = 0 To Len(ref) - 1
                curToken = prevToken
                r = ref.Substring(i, 1)

                Select Case r
                    Case ":"
                        'a new token
                        If prevToken.TokenType = cReferenceToken.BibleRefTokenType.Space Then
                            'get rid of spaces before a colon
                            mReferenceTokens.RemoveAt(mReferenceTokens.Count - 1)
                            If mReferenceTokens.Count > 0 Then
                                prevToken = mReferenceTokens(mReferenceTokens.Count - 1)
                            Else
                                prevToken = Nothing
                            End If

                            curToken = prevToken
                        End If
                        curToken = New cReferenceToken
                        mReferenceTokens.Add(curToken)
                        curToken.ReferencePart = r
                        curToken.TokenType = cReferenceToken.BibleRefTokenType.ChapterVerseDelim

                        'alway set previous token to chapter
                        prevToken.TokenType = cReferenceToken.BibleRefTokenType.Chapter
                        'Case "."
                        '    If prevToken.IsANumber Or prevToken.TokenType = cReferenceToken.BibleRefTokenType.Chapter Then
                        '        'convert this to a ":"
                        '        curToken = New cReferenceToken
                        '        mReferenceTokens.Add(curToken)
                        '        curToken.ReferencePart = ":"
                        '        curToken.TokenType = cReferenceToken.BibleRefTokenType.ChapterVerseDelim
                        '    End If
                    Case ",", ";", "&"
                        If prevToken.TokenType = cReferenceToken.BibleRefTokenType.Space Then
                            'get rid of spaces before a comma
                            mReferenceTokens.RemoveAt(mReferenceTokens.Count - 1)
                            If mReferenceTokens.Count > 0 Then
                                prevToken = mReferenceTokens(mReferenceTokens.Count - 1)
                            Else
                                prevToken = Nothing
                            End If

                            curToken = prevToken
                        End If
                        'a new token
                        curToken = New cReferenceToken
                        mReferenceTokens.Add(curToken)
                        curToken.ReferencePart = ","
                        If prevToken.TokenType = cReferenceToken.BibleRefTokenType.book Or _
                           prevToken.TokenType = cReferenceToken.BibleRefTokenType.Chapter Or _
                           prevToken.TokenType = cReferenceToken.BibleRefTokenType.verse Then
                            curToken.TokenType = cReferenceToken.BibleRefTokenType.ListDelimiter
                        Else
                            curToken.TokenType = cReferenceToken.BibleRefTokenType.Unknown
                        End If
                    Case "-"
                        If prevToken.TokenType = cReferenceToken.BibleRefTokenType.Space Then
                            'get rid of spaces before a dash
                            mReferenceTokens.RemoveAt(mReferenceTokens.Count - 1)
                            If mReferenceTokens.Count > 0 Then
                                prevToken = mReferenceTokens(mReferenceTokens.Count - 1)
                            Else
                                prevToken = Nothing
                            End If

                            curToken = prevToken
                        End If
                        'a new token
                        curToken = New cReferenceToken
                        mReferenceTokens.Add(curToken)
                        curToken.ReferencePart = r
                        If prevToken.TokenType = cReferenceToken.BibleRefTokenType.book Or _
                           prevToken.TokenType = cReferenceToken.BibleRefTokenType.Chapter Or _
                           prevToken.TokenType = cReferenceToken.BibleRefTokenType.verse Then
                            curToken.TokenType = cReferenceToken.BibleRefTokenType.RangeDelimiter
                        Else
                            curToken.TokenType = cReferenceToken.BibleRefTokenType.Unknown
                        End If
                    Case " "
                        If prevToken.TokenType = cReferenceToken.BibleRefTokenType.ListDelimiter Or prevToken.TokenType = cReferenceToken.BibleRefTokenType.RangeDelimiter Or prevToken.TokenType = cReferenceToken.BibleRefTokenType.ChapterVerseDelim Or prevToken.TokenType = cReferenceToken.BibleRefTokenType.Space Then
                            'skip spaces after list/range delimiters and duplicate spaces
                        Else
                            If prevToken.TokenType = cReferenceToken.BibleRefTokenType.book And
                               prevToken.ReferencePart = "vs" Then
                                'convert previous token to chapterverse delimiter
                                prevToken.TokenType = cReferenceToken.BibleRefTokenType.ChapterVerseDelim
                                prevToken.ReferencePart = ":"
                            ElseIf prevToken.TokenType = cReferenceToken.BibleRefTokenType.book And
                                Not mLibrary.IsBookNameValid(prevToken.ReferencePart) Then
                                'clear out prevtoken
                                prevToken.ReferencePart = ""
                                curToken.ReferencePart = ""
                            Else
                                curToken = New cReferenceToken
                                mReferenceTokens.Add(curToken)
                                curToken.ReferencePart = r
                                curToken.TokenType = cReferenceToken.BibleRefTokenType.Space
                            End If

                        End If
                    Case "."
                        If prevToken.TokenType = cReferenceToken.BibleRefTokenType.Chapter Then
                            'this is a chapter verse delim
                            curToken = New cReferenceToken
                            mReferenceTokens.Add(curToken)
                            curToken.ReferencePart = ":"
                            curToken.TokenType = cReferenceToken.BibleRefTokenType.ChapterVerseDelim
                        ElseIf prevToken.TokenType = cReferenceToken.BibleRefTokenType.verse Then
                            'this is a list delim
                            curToken = New cReferenceToken
                            mReferenceTokens.Add(curToken)
                            curToken.ReferencePart = ","
                            curToken.TokenType = cReferenceToken.BibleRefTokenType.ListDelimiter
                        End If
                    Case Else
                        'dealing with numbers or text
                        If IsNumeric(r) Then
                            If prevToken Is Nothing Then
                                'start of a new token
                                curToken = New cReferenceToken
                                mReferenceTokens.Add(curToken)
                                curToken.ReferencePart = r
                                curToken.TokenType = cReferenceToken.BibleRefTokenType.book
                            Else
                                Select Case prevToken.TokenType
                                    Case cReferenceToken.BibleRefTokenType.Chapter, cReferenceToken.BibleRefTokenType.verse
                                        'append this to previous token
                                        prevToken.ReferencePart = prevToken.ReferencePart & r
                                    Case cReferenceToken.BibleRefTokenType.Space
                                        'see if I need to keep the space
                                        curToken = New cReferenceToken
                                        mReferenceTokens.Add(curToken)
                                        curToken.ReferencePart = r
                                        If mReferenceTokens(mReferenceTokens.Count - 3).TokenType = cReferenceToken.BibleRefTokenType.book Then
                                            curToken.TokenType = cReferenceToken.BibleRefTokenType.Chapter
                                            prevToken.TokenType = cReferenceToken.BibleRefTokenType.BookChapterDelim
                                        ElseIf mReferenceTokens(mReferenceTokens.Count - 3).TokenType = cReferenceToken.BibleRefTokenType.Chapter Then
                                            curToken.TokenType = cReferenceToken.BibleRefTokenType.verse
                                            prevToken.TokenType = cReferenceToken.BibleRefTokenType.ChapterVerseDelim
                                        ElseIf mReferenceTokens(mReferenceTokens.Count - 3).TokenType = cReferenceToken.BibleRefTokenType.verse Then
                                            curToken.TokenType = cReferenceToken.BibleRefTokenType.verse
                                        Else
                                            curToken.TokenType = cReferenceToken.BibleRefTokenType.Unknown
                                        End If
                                    Case cReferenceToken.BibleRefTokenType.book, cReferenceToken.BibleRefTokenType.ChapterListDelim, cReferenceToken.BibleRefTokenType.ChapterRangeDelim
                                        curToken = New cReferenceToken
                                        mReferenceTokens.Add(curToken)
                                        curToken.ReferencePart = r
                                        curToken.TokenType = cReferenceToken.BibleRefTokenType.Chapter
                                    Case cReferenceToken.BibleRefTokenType.ChapterVerseDelim, cReferenceToken.BibleRefTokenType.VerseListDelim, cReferenceToken.BibleRefTokenType.VerseRangeDelim
                                        curToken = New cReferenceToken
                                        mReferenceTokens.Add(curToken)
                                        curToken.ReferencePart = r
                                        curToken.TokenType = cReferenceToken.BibleRefTokenType.verse
                                    Case cReferenceToken.BibleRefTokenType.ListDelimiter
                                        curToken = New cReferenceToken
                                        mReferenceTokens.Add(curToken)
                                        curToken.ReferencePart = r
                                        If mReferenceTokens.Count > 2 Then
                                            'assume it's the same type as the prior number
                                            If mReferenceTokens(mReferenceTokens.Count - 3).IsANumber Then
                                                curToken.TokenType = mReferenceTokens(mReferenceTokens.Count - 3).TokenType
                                            Else
                                                'it's just a number for now
                                                curToken.TokenType = cReferenceToken.BibleRefTokenType.Number
                                            End If
                                        Else
                                            'it's just a number for now
                                            curToken.TokenType = cReferenceToken.BibleRefTokenType.Number
                                        End If
                                    Case cReferenceToken.BibleRefTokenType.RangeDelimiter
                                        curToken = New cReferenceToken
                                        mReferenceTokens.Add(curToken)
                                        curToken.ReferencePart = r
                                        If mReferenceTokens.Count > 2 Then
                                            'assume it's the same type as the prior number
                                            If mReferenceTokens(mReferenceTokens.Count - 3).IsANumber Then
                                                curToken.TokenType = mReferenceTokens(mReferenceTokens.Count - 3).TokenType
                                            Else
                                                'it's just a number for now
                                                curToken.TokenType = cReferenceToken.BibleRefTokenType.Number
                                            End If
                                        Else
                                            'it's just a number for now
                                            curToken.TokenType = cReferenceToken.BibleRefTokenType.Number
                                        End If
                                End Select
                            End If
                        Else
                            'it's a string
                            If Asc(r) >= Asc("a") And Asc(r) <= Asc("z") And r <> "q" Then
                                If prevToken Is Nothing Then
                                    'this is a new book reference
                                    curToken = New cReferenceToken
                                    mReferenceTokens.Add(curToken)
                                    curToken.ReferencePart = r
                                    curToken.TokenType = cReferenceToken.BibleRefTokenType.book
                                ElseIf prevToken.TokenType = cReferenceToken.BibleRefTokenType.book Then
                                    'this is a current reference
                                    prevToken.ReferencePart = prevToken.ReferencePart & r

                                    'special case for specific book with numeric prefixes
                                    Select Case prevToken.ReferencePart.ToLower
                                        Case "sam", "kin", "kgs", "chro", "chr", "cor", "tim", "the", "pet", "john", "jn"
                                            If mReferenceTokens.Count > 1 Then
                                                If mReferenceTokens(mReferenceTokens.Count - 2).TokenType = cReferenceToken.BibleRefTokenType.Space Then
                                                    If mReferenceTokens.Count > 2 Then
                                                        If (IsNumeric(mReferenceTokens(mReferenceTokens.Count - 3).ReferencePart) And _
                                                            (mReferenceTokens(mReferenceTokens.Count - 3).ReferencePart = "1" Or _
                                                             mReferenceTokens(mReferenceTokens.Count - 3).ReferencePart = "2" Or _
                                                             mReferenceTokens(mReferenceTokens.Count - 3).ReferencePart = "3")) Or _
                                                             mReferenceTokens(mReferenceTokens.Count - 3).ReferencePart = "i" Or _
                                                             mReferenceTokens(mReferenceTokens.Count - 3).ReferencePart = "ii" Then
                                                            'merge all these
                                                            prevToken.ReferencePart = mReferenceTokens(mReferenceTokens.Count - 3).ReferencePart & mReferenceTokens(mReferenceTokens.Count - 2).ReferencePart & prevToken.ReferencePart
                                                            mReferenceTokens.RemoveAt(mReferenceTokens.Count - 2)
                                                            mReferenceTokens.RemoveAt(mReferenceTokens.Count - 2)
                                                        End If
                                                    End If
                                                ElseIf (IsNumeric(mReferenceTokens(mReferenceTokens.Count - 2).ReferencePart) And _
                                                        (mReferenceTokens(mReferenceTokens.Count - 2).ReferencePart = "1" Or _
                                                         mReferenceTokens(mReferenceTokens.Count - 2).ReferencePart = "2" Or _
                                                         mReferenceTokens(mReferenceTokens.Count - 2).ReferencePart = "3")) Or _
                                                         mReferenceTokens(mReferenceTokens.Count - 2).ReferencePart = "i" Or _
                                                         mReferenceTokens(mReferenceTokens.Count - 2).ReferencePart = "ii" Then
                                                    'merge  these
                                                    prevToken.ReferencePart = mReferenceTokens(mReferenceTokens.Count - 2).ReferencePart & prevToken.ReferencePart
                                                    mReferenceTokens.RemoveAt(mReferenceTokens.Count - 2)
                                                End If
                                            End If
                                    End Select

                                ElseIf prevToken.TokenType = cReferenceToken.BibleRefTokenType.Space Then
                                    'see if two tokens ago was a book
                                    If mReferenceTokens(mReferenceTokens.Count - 2).TokenType = cReferenceToken.BibleRefTokenType.book Then
                                        'get rid of the space, and continue with this book
                                        mReferenceTokens.RemoveAt(mReferenceTokens.Count - 1)
                                        prevToken = mReferenceTokens(mReferenceTokens.Count - 1)
                                        curToken = prevToken
                                        If prevToken.ReferencePart.ToLower = "and" Then
                                            'get rid of ands entirely
                                            prevToken.ReferencePart = r
                                        Else
                                            prevToken.ReferencePart = prevToken.ReferencePart & " " & r
                                        End If
                                        'ElseIf mReferenceTokens(mReferenceTokens.Count - 2).TokenType = cReferenceToken.BibleRefTokenType.Chapter And
                                        '       (curToken.ReferencePart.ToLower = "vs") Then
                                        '    'just convert current token to a chapterVerse delim
                                    Else
                                        curToken = New cReferenceToken
                                        mReferenceTokens.Add(curToken)
                                        curToken.ReferencePart = r
                                        curToken.TokenType = cReferenceToken.BibleRefTokenType.book
                                    End If
                                Else
                                    'now add a new book token
                                    curToken = New cReferenceToken
                                    mReferenceTokens.Add(curToken)
                                    curToken.ReferencePart = r
                                    curToken.TokenType = cReferenceToken.BibleRefTokenType.book
                                End If
                            Else
                                'Debug.Print "Found extraneous character in reference: " & r
                            End If
                        End If
                End Select
                prevToken = curToken
            Next

            'get rid of odd stuff at beginning and end
            While Not done
                If mReferenceTokens.Count = 0 Then
                    done = True
                ElseIf mReferenceTokens(0).TokenType = cReferenceToken.BibleRefTokenType.Space Or mReferenceTokens(0).TokenType <> cReferenceToken.BibleRefTokenType.book Then
                    mReferenceTokens.RemoveAt(0)
                ElseIf (mReferenceTokens(mReferenceTokens.Count - 1).TokenType <> cReferenceToken.BibleRefTokenType.book And mReferenceTokens(mReferenceTokens.Count - 1).TokenType <> cReferenceToken.BibleRefTokenType.Chapter And mReferenceTokens(mReferenceTokens.Count - 1).TokenType <> cReferenceToken.BibleRefTokenType.verse) Then
                    mReferenceTokens.RemoveAt(mReferenceTokens.Count - 1)
                Else
                    done = True
                End If
            End While

            'everything should be ok... now clean up
            '1. Remove internal spaces
            i = 1    'can't start with a space or end with one
            maxTokens = mReferenceTokens.Count - 2
            While i <= maxTokens
                curToken = mReferenceTokens(i)
                prevToken = mReferenceTokens(i - 1)
                nextToken = mReferenceTokens(i + 1)
                If curToken.TokenType = cReferenceToken.BibleRefTokenType.Space Then
                    If prevToken.TokenType = cReferenceToken.BibleRefTokenType.book And nextToken.TokenType = cReferenceToken.BibleRefTokenType.book Then
                        curToken.TokenType = cReferenceToken.BibleRefTokenType.BookRangeDelim
                    ElseIf prevToken.TokenType = cReferenceToken.BibleRefTokenType.Chapter And nextToken.TokenType = cReferenceToken.BibleRefTokenType.verse Then
                        curToken.TokenType = cReferenceToken.BibleRefTokenType.ChapterVerseDelim
                    ElseIf prevToken.TokenType = cReferenceToken.BibleRefTokenType.verse And nextToken.TokenType = cReferenceToken.BibleRefTokenType.verse Then
                        curToken.TokenType = cReferenceToken.BibleRefTokenType.VerseListDelim
                    Else
                        'don't need this space
                        mReferenceTokens.RemoveAt(i)
                        i = i - 1
                        maxTokens = maxTokens - 1
                    End If
                ElseIf prevToken.TokenType = cReferenceToken.BibleRefTokenType.book And
                       prevToken.ReferencePart.ToLower.StartsWith("chap") Then
                    'this could be "chapter 1 of daniel"
                    '                  p   in23 4   5
                    If maxTokens <= i + 5 Then
                        If curToken.TokenType = cReferenceToken.BibleRefTokenType.Space And
                           nextToken.IsANumber And
                           mReferenceTokens(i + 2).TokenType = cReferenceToken.BibleRefTokenType.Space And
                           mReferenceTokens(i + 3).TokenType = cReferenceToken.BibleRefTokenType.book And
                           mReferenceTokens(i + 3).ReferencePart.ToLower = "of" And
                           mReferenceTokens(i + 4).TokenType = cReferenceToken.BibleRefTokenType.Space And
                           mReferenceTokens(i + 5).TokenType = cReferenceToken.BibleRefTokenType.book Then

                            'made it here!
                            '1. replace "chapter with bookname
                            prevToken.ReferencePart = mReferenceTokens(i + 5).ReferencePart
                            '2 make the next number a chapter
                            nextToken.TokenType = cReferenceToken.BibleRefTokenType.Chapter
                            '3 delete the tokens not needed
                            mReferenceTokens.RemoveAt(i + 2) 'space
                            mReferenceTokens.RemoveAt(i + 2) 'of
                            mReferenceTokens.RemoveAt(i + 2) 'space
                            mReferenceTokens.RemoveAt(i + 2) 'daniel
                            maxTokens -= 4
                        End If
                    End If
                ElseIf prevToken.TokenType = cReferenceToken.BibleRefTokenType.Space And
                       curToken.TokenType = cReferenceToken.BibleRefTokenType.book And
                       Not mLibrary.IsBookNameValid(curToken.ReferencePart) And
                       nextToken.TokenType = cReferenceToken.BibleRefTokenType.Space Then
                    'get rid of current token and next token

                    mReferenceTokens.RemoveAt(i)
                    mReferenceTokens.RemoveAt(i)
                    i -= 1
                    maxTokens -= 2
                ElseIf prevToken.TokenType = cReferenceToken.BibleRefTokenType.book Then
                    'check for special case book names
                    prevToken = CheckTokenForSpecialWords(prevToken)
                    If prevToken.TokenType = cReferenceToken.BibleRefTokenType.InvalidCharacter Then
                        'don't need this token
                        mReferenceTokens.RemoveAt(i - 1)
                        i = i - 1
                        maxTokens -= i
                    End If
                ElseIf curToken.TokenType = cReferenceToken.BibleRefTokenType.book Then
                    'check for special case book names
                    curToken = CheckTokenForSpecialWords(curToken)
                    If curToken.TokenType = cReferenceToken.BibleRefTokenType.InvalidCharacter Then
                        'don't need this token
                        mReferenceTokens.RemoveAt(i)
                        i -= 1
                        maxTokens -= 1
                    End If
                ElseIf nextToken.TokenType = cReferenceToken.BibleRefTokenType.book Then
                    'check for special case book names
                    nextToken = CheckTokenForSpecialWords(nextToken)
                    If nextToken.TokenType = cReferenceToken.BibleRefTokenType.InvalidCharacter Then
                        'don't need this token
                        mReferenceTokens.RemoveAt(i + 1)
                        maxTokens -= 1
                    End If
                End If

                i = i + 1
            End While

            '2. Get rid of bookchapter delims, chapterverse delims.
            i = 1    'can't start or end with one
            maxTokens = mReferenceTokens.Count - 2
            While i <= maxTokens
                curToken = mReferenceTokens(i)

                If curToken.TokenType = cReferenceToken.BibleRefTokenType.BookChapterDelim Or _
                   curToken.TokenType = cReferenceToken.BibleRefTokenType.ChapterVerseDelim Then
                    mReferenceTokens.RemoveAt(i)
                    i = i - 1
                    maxTokens = maxTokens - 1
                End If
                i = i + 1
            End While

            '3. now make partial list items, full list items.
            'i.e. John 3:16, 20 = John 3:16, John 3:20
            'or   John 3:16, 20:34 = John 3:16, John 20:34
            'i.e. chapters should always be preceeded by a book and followed by a verse
            'i = 1
            i = 0
            maxTokens = mReferenceTokens.Count - 1
            While i <= maxTokens
                'Debug.Print i
                curToken = mReferenceTokens(i)
                If i > 0 Then
                    prevToken = mReferenceTokens(i - 1)
                Else
                    prevToken = Nothing
                End If
                If i < maxTokens Then
                    nextToken = mReferenceTokens(i + 1)
                Else
                    nextToken = Nothing
                End If

                If curToken.TokenType = cReferenceToken.BibleRefTokenType.book Then
                    'should always be followed by a chapter
                    addFullBook = False
                    If nextToken Is Nothing Then
                        addFullBook = True
                    ElseIf nextToken.TokenType <> cReferenceToken.BibleRefTokenType.Chapter Then
                        addFullBook = True
                    End If
                    If addFullBook Then
                        'insert chapter 1
                        mReferenceTokens.Insert(i + 1, chapter1Token.Duplicate)
                        'insert verse 1
                        mReferenceTokens.Insert(i + 2, verse1Token.Duplicate)
                        i = i + 2
                        maxTokens = maxTokens + 2
                    End If

                    prevBookToken = curToken
                ElseIf curToken.TokenType = cReferenceToken.BibleRefTokenType.Chapter Then
                    addFullChapter = False
                    prevChapterToken = curToken

                    If prevToken.TokenType <> cReferenceToken.BibleRefTokenType.book Then
                        'need to insert a book right here
                        mReferenceTokens.Insert(i, prevBookToken.Duplicate)
                        i = i + 1
                        maxTokens = maxTokens + 1
                    End If
                    If nextToken Is Nothing Then
                        addFullChapter = True
                    ElseIf nextToken.TokenType <> cReferenceToken.BibleRefTokenType.verse Then
                        'one more check for single chapter books
                        If NeedToInsertChapter1ForSingleChapterBook(prevToken.ReferencePart, curToken.ReferencePart) Then
                            '1. Convert current token to verse
                            curToken.TokenType = cReferenceToken.BibleRefTokenType.verse
                            Dim tmpToken As New cReferenceToken With {.TokenType = cReferenceToken.BibleRefTokenType.ChapterVerseDelim, .ReferencePart = ":"}
                            mReferenceTokens.Insert(i, tmpToken)
                            mReferenceTokens.Insert(i, chapter1Token.Duplicate)
                            prevChapterToken = mReferenceTokens(i)
                            i += 2
                            maxTokens += 2 'always = count-1

                            'now looking ahead, convert any other chap nums for this book to verse nums
                            Dim j As Integer
                            For j = i + 1 To maxTokens
                                Select Case mReferenceTokens(j).TokenType
                                    Case cReferenceToken.BibleRefTokenType.book
                                        Exit For
                                    Case cReferenceToken.BibleRefTokenType.Chapter
                                        mReferenceTokens(j).TokenType = cReferenceToken.BibleRefTokenType.verse
                                    Case cReferenceToken.BibleRefTokenType.ChapterListDelim
                                        mReferenceTokens(j).TokenType = cReferenceToken.BibleRefTokenType.VerseListDelim
                                    Case cReferenceToken.BibleRefTokenType.ChapterRangeDelim
                                        mReferenceTokens(j).TokenType = cReferenceToken.BibleRefTokenType.VerseRangeDelim
                                    Case cReferenceToken.BibleRefTokenType.RangeDelimiter, cReferenceToken.BibleRefTokenType.ListDelimiter
                                        'do nothing
                                    Case Else
                                        Exit For
                                End Select
                            Next
                        Else
                            addFullChapter = True
                        End If
                    End If
                    If addFullChapter = True Then
                        'add a verse 1 token
                        mReferenceTokens.Insert(i + 1, verse1Token.Duplicate)

                        i = i + 1
                        maxTokens = maxTokens + 1
                    End If

                ElseIf curToken.TokenType = cReferenceToken.BibleRefTokenType.verse Then
                    If prevToken.TokenType <> cReferenceToken.BibleRefTokenType.Chapter Then
                        'add previous chapter
                        mReferenceTokens.Insert(i, prevChapterToken.Duplicate)
                        i = i - 1 'move i back one so I can check the chapter
                        maxTokens = maxTokens + 1
                    End If
                ElseIf curToken.TokenType = cReferenceToken.BibleRefTokenType.BookListDelim Or _
                       curToken.TokenType = cReferenceToken.BibleRefTokenType.ChapterListDelim Or _
                       curToken.TokenType = cReferenceToken.BibleRefTokenType.VerseListDelim Then
                    curToken.TokenType = cReferenceToken.BibleRefTokenType.ListDelimiter
                ElseIf curToken.TokenType = cReferenceToken.BibleRefTokenType.BookRangeDelim Or _
                       curToken.TokenType = cReferenceToken.BibleRefTokenType.ChapterRangeDelim Or _
                       curToken.TokenType = cReferenceToken.BibleRefTokenType.VerseRangeDelim Then
                    curToken.TokenType = cReferenceToken.BibleRefTokenType.RangeDelimiter
                End If
                i = i + 1
            End While

            ParseTokens(findFirst2VersesOnly)
        Catch ex As Exception
            Debug.Print("In cBibleReference.ParseReference: " & ex.Message)
        End Try

    End Sub
    Function NeedToInsertChapter1ForSingleChapterBook(BookName As String, ChapterNumber As Integer) As Boolean
        'many times, single chapter books don't include 1:14, but just 14.
        Dim sn As String = mLibrary.GetBookShortName(BookName)
        Return (sn = "Oba" Or sn = "Phile" Or sn = "2 Jn" Or sn = "3 Jn" Or sn = "Jude") And ChapterNumber > 1
    End Function
    Private Function CheckTokenForSpecialWords(inToken As cReferenceToken) As cReferenceToken
        Dim newToken As cReferenceToken = inToken
        If inToken.TokenType = cReferenceToken.BibleRefTokenType.book Then
            Select Case inToken.ReferencePart
                Case "to"
                    'convert this to a range delimiter
                    newToken = New cReferenceToken
                    newToken.ReferencePart = "-"
                    newToken.TokenType = cReferenceToken.BibleRefTokenType.RangeDelimiter
                Case "vs"
                    'convert this to a verse delimiter
                    newToken = New cReferenceToken
                    newToken.ReferencePart = ":"
                    newToken.TokenType = cReferenceToken.BibleRefTokenType.ChapterVerseDelim
                Case "chapter"
                    'just delete this token
                    newToken = New cReferenceToken
                    newToken.ReferencePart = " "
                    newToken.TokenType = cReferenceToken.BibleRefTokenType.InvalidCharacter
            End Select
        End If

        Return newToken
    End Function
    Private Function PreparseReference(ByVal InValue As String) As String
        Dim ref As String

        'strip extra spaces / lower case
        ref = LCase(Trim(InValue))

        'double spaces
        ref = Replace(ref, "  ", " ")

        Return ref
    End Function
    Private Function GetLastChapterToken(bookToken As cReferenceToken) As cReferenceToken
        Dim ct As New cReferenceToken

        On Error Resume Next
        ct.TokenType = cReferenceToken.BibleRefTokenType.Chapter
        ct.ReferencePart = Me.Library.BibleBooks(Me.Library.GetBookShortName(bookToken.ReferencePart)).Chapters

        Return ct

    End Function
    Private Function GetLastVerseToken(bookToken As cReferenceToken, chapterToken As cReferenceToken) As cReferenceToken
        Dim vt As New cReferenceToken

        On Error Resume Next
        vt.TokenType = cReferenceToken.BibleRefTokenType.verse
        vt.ReferencePart = Me.Library.CurrentVersion.VerseCountOfChapter(Me.Library.GetBookNumber(bookToken.ReferencePart), chapterToken.ReferencePart)

        Return vt
    End Function
    Private Sub ParseTokens(findFirst2VersesOnly As Boolean)
        Dim lastBN As Integer
        Dim bn As Integer
        Dim cn As Integer
        Dim lastCN As Integer
        Dim vn As Integer
        Dim lastVN As Integer
        Dim rt As cReferenceToken
        Dim i As Integer
        Dim foundRange As Boolean
        Dim verse As cVerse
        Dim pVerse As cVerse
        Dim lastVerseIndex As String
        Dim newBCV As String
        Dim done As Boolean

        'convert range of tokens to actual verse indexes
        mBadReferences = ""

        Try
            mVerses = New List(Of cVerse)
            mShortReference = ""
            mLongReference = ""

            For i = 0 To mReferenceTokens.Count - 1
                'Debug.Print(i)

                rt = mReferenceTokens(i)

                Select Case rt.TokenType
                    Case cReferenceToken.BibleRefTokenType.book
                        lastBN = bn
                        bn = mLibrary.GetBookNumber(rt.ReferencePart)
                        If bn = 0 Then
                            mShortReference = mShortReference & rt.ReferencePart & " "
                            mLongReference = mLongReference & rt.ReferencePart & " "
                        ElseIf lastBN <> bn Then
                            If Right(mShortReference, 1) = "," Then
                                'add space after comma when book changes
                                mShortReference = mShortReference & " "
                                mLongReference = mLongReference & " "
                            End If
                            mShortReference = mShortReference & mLibrary.BibleBooks(mLibrary.BibleBooks.Keys(bn - 1)).ShortName & " "
                            mLongReference = mLongReference & mLibrary.BibleBooks(mLibrary.BibleBooks.Keys(bn - 1)).LongName & " "
                        End If
                    Case cReferenceToken.BibleRefTokenType.Chapter
                        lastCN = cn
                        cn = rt.ReferencePart

                        If lastBN = bn And lastCN > cn Then
                            'chapter number decreased. Include book name
                            If Right(mShortReference, 1) = "," Then
                                'add space after comma when book changes
                                mShortReference = mShortReference & " "
                                mLongReference = mLongReference & " "
                            End If
                            mShortReference = mShortReference & mLibrary.BibleBooks(mLibrary.BibleBooks.Keys(bn - 1)).ShortName & " "
                            mLongReference = mLongReference & mLibrary.BibleBooks(mLibrary.BibleBooks.Keys(bn - 1)).LongName & " "
                        End If

                        If lastBN <> bn Or lastCN <> cn Then
                            mShortReference = mShortReference & cn & ":"
                            mLongReference = mLongReference & cn & ":"
                        End If
                    Case cReferenceToken.BibleRefTokenType.verse
                        lastVN = vn
                        vn = rt.ReferencePart

                        'I found a verse. is this part of a range?
                        If foundRange Then
                            'go from start range to current range
                            done = False
                            While Not done  'lastBN <= bn And lastCN <= cn And lastVN <= vn
                                verse = New cVerse
                                verse.Library = mLibrary
                                verse.BookChapterVerseIndex = Library.CurrentVersion.GetNextVerseIndex(lastBN, lastCN, lastVN)
                                If lastBN = bn And (lastCN < cn Or (lastCN = cn And lastVN <= vn)) Then
                                    If verse.IsValid Then
                                        mVerses.Add(verse)
                                        If findFirst2VersesOnly And mVerses.Count = 2 Then Exit For
                                    Else
                                        lastBN = bn
                                        lastCN = cn
                                        lastVN = vn
                                    End If
                                Else
                                    'stop at previous verse
                                    verse.BookChapterVerseIndex = Library.CurrentVersion.GetPreviousVerseIndex(lastBN, lastCN, lastVN)
                                    vn = lastVN
                                    lastVN = vn + 1
                                    done = True
                                End If
                            End While
                            foundRange = False
                        Else
                            'no, just do this verse
                            verse = New cVerse
                            verse.Library = mLibrary
                            verse.BookNumber = bn
                            verse.ChapterNumber = cn
                            verse.VerseNumber = vn
                            If verse.IsValid Then
                                mVerses.Add(verse)
                            End If
                            If findFirst2VersesOnly And mVerses.Count = 2 Then Exit For
                        End If

                        If lastBN <> bn Or lastCN <> cn Or lastVN <> vn Then
                            mShortReference = mShortReference & vn
                            mLongReference = mLongReference & vn
                        ElseIf Right(mShortReference, 1) = "," Then
                            mShortReference = Left(mShortReference, Len(mShortReference) - 1)
                            mLongReference = Left(mLongReference, Len(mLongReference) - 1)
                        End If

                    Case cReferenceToken.BibleRefTokenType.RangeDelimiter
                        foundRange = True
                        mShortReference = mShortReference & "-"
                        mLongReference = mLongReference & "-"
                    Case cReferenceToken.BibleRefTokenType.ListDelimiter
                        mShortReference = mShortReference & ","
                        mLongReference = mLongReference & ","
                End Select
            Next

            'support for paragraph type of bibles. If the bible text = "<p />" then get previous verse.
            For Each verse In mVerses
                If verse.VerseText = "<p />" Then
                    'get the previous verse
                    'Set pVerse = verse
                    lastVerseIndex = 0
                    While verse.VerseText = "<p />" And lastVerseIndex <> verse.BookChapterVerseIndex
                        lastVerseIndex = verse.BookChapterVerseIndex
                        newBCV = verse.Library.CurrentVersion.GetPreviousVerseIndex(verse.BookNumber, verse.ChapterNumber, verse.VerseNumber)
                        verse.BookNumber = Left(newBCV, 2)
                        verse.ChapterNumber = Mid(newBCV, 2, 3)
                        verse.VerseNumber = Right(newBCV, 3)
                    End While
                End If
            Next

            'I might hve duplicates at this point. double check
            pVerse = Nothing
            If mVerses.Count > 1 Then
                For i = mVerses.Count - 1 To 1 Step -1
                    If mVerses(i).BookChapterVerseIndex = mVerses(i - 1).BookChapterVerseIndex Then
                        mVerses.RemoveAt(i)
                    End If
                Next
            End If

        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            Debug.Print("In cBibleReference.ParseTokens: @line:" & st.GetFrame(0).GetFileLineNumber.ToString & " " & ex.Message)
        End Try
    End Sub
    Public Sub ListReferenceTokens()
        Dim rt As cReferenceToken

        For Each rt In mReferenceTokens
            Debug.Print(rt.ReferencePart, rt.TokenType)
        Next
    End Sub

End Class
