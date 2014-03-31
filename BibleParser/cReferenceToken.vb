Public Class cReferenceToken

    Public Enum BibleRefTokenType
        Unknown = 0
        book = 1
        Chapter = 2
        verse = 3
        BookChapterDelim = 4
        ChapterVerseDelim = 5
        BookListDelim = 6
        BookRangeDelim = 7
        VerseListDelim = 8
        VerseRangeDelim = 9
        ChapterListDelim = 10
        ChapterRangeDelim = 11
        ListDelimiter = 12
        RangeDelimiter = 13
        Space = 14
        Number = 15
        RedLetter = 16
        Paragraph = 17
        InvalidCharacter = 18
    End Enum


    Private mReferencePart As String
    Private mTokenType As BibleRefTokenType

    Public Function Duplicate() As cReferenceToken
        Dim nt As New cReferenceToken

        nt.ReferencePart = Me.ReferencePart
        nt.TokenType = Me.TokenType

        Duplicate = nt
    End Function
    Public Property ReferencePart() As String
        Get
            If mTokenType = BibleRefTokenType.Chapter Or mTokenType = BibleRefTokenType.verse Or mTokenType = BibleRefTokenType.Number Then
                If Not IsNumeric(mReferencePart) Then
                    Return 0
                Else
                    Return mReferencePart
                End If
            Else
                Return mReferencePart
            End If
        End Get
        Set(value As String)
            mReferencePart = value
        End Set
    End Property
    Public Property TokenType() As BibleRefTokenType
        Get
            Return mTokenType
        End Get
        Set(value As BibleRefTokenType)
            mTokenType = value
        End Set
    End Property
    Public ReadOnly Property IsANumber() As Boolean
        Get
            Return IsNumeric(mReferencePart)
        End Get
    End Property
End Class
