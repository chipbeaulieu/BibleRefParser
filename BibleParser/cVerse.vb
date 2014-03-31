Imports System.ComponentModel
Public Class cVerse
    Implements INotifyPropertyChanged

    Private mBookChapterVerseIndex As String
    Private mBookNumber As Integer
    Private mChapterNumber As Integer
    Private mVerseNumber As Integer
    Private mVerseText As String
    Private mLibrary As cLibrary
    Private mBookShortName As String
    Private mBookLongName As String

#Region "Public Events"
    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged
#End Region 'Public Events

    ''' <summary>
    '''  
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Property BookChapterVerseIndex() As String
        Get
            Return mBookChapterVerseIndex
        End Get
        Set(ByVal value As String)
            mBookChapterVerseIndex = value.trim
            mVerseText = mLibrary.CurrentVersion.Verses(value)
            BookNumber = value.Substring(0, 2)
            mChapterNumber = value.Substring(2, 3)
            mVerseNumber = value.Substring(Math.Max(0, value.Length - 3))
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("BookChapterVerseIndex"))
        End Set
    End Property
    ''' <summary>
    '''  
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Property BookNumber() As Integer
        Get
            Return mBookNumber
        End Get
        Set(ByVal value As Integer)
            mBookNumber = value
            BuildBookChapterVerseIndex()
            If value > 0 And value <= mLibrary.BibleBooks.Count Then
                mBookShortName = mLibrary.BibleBooks(mLibrary.BibleBooks.Keys(value - 1)).ShortName
                mBookLongName = mLibrary.BibleBooks(mLibrary.BibleBooks.Keys(value - 1)).LongName
            End If
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("BookNumber"))
        End Set
    End Property
    ''' <summary>
    '''  
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Property ChapterNumber() As Integer
        Get
            Return mChapterNumber
        End Get
        Set(ByVal value As Integer)
            mChapterNumber = value
            BuildBookChapterVerseIndex()
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("ChapterNumber"))
        End Set
    End Property
    Public Function IsValid() As Boolean
        Dim result As Boolean = (BookNumber * ChapterNumber * VerseNumber > 0)

        If Library.CurrentVersion IsNot Nothing Then
            result = result And Library.CurrentVersion.IsValid(Library.BookChapterVerseIndex(BookNumber, ChapterNumber, VerseNumber))
        End If
        Return result
    End Function
    ''' <summary>
    '''  
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Property VerseNumber() As Integer
        Get
            Return mVerseNumber
        End Get
        Set(ByVal value As Integer)
            mVerseNumber = value
            BuildBookChapterVerseIndex()
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("VerseNumber"))
        End Set
    End Property
    ''' <summary>
    '''  
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property VerseText() As String
        Get
            Return mVerseText
        End Get
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
    Public Property BookShortName() As String
        Get
            Return mBookShortName
        End Get
        Set(ByVal value As String)
            mBookShortName = value.trim
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("BookShortName"))
        End Set
    End Property
    ''' <summary>
    '''  
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Property BookLongName() As String
        Get
            Return mBookLongName
        End Get
        Set(ByVal value As String)
            mBookLongName = value.trim
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("BookLongName"))
        End Set
    End Property

    Public ReadOnly Property ShortReference As String
        Get
            Return BookShortName & " " & ChapterNumber & ":" & VerseNumber
        End Get
    End Property
    Public ReadOnly Property LongReference As String
        Get
            Return BookLongName & " " & ChapterNumber & ":" & VerseNumber
        End Get
    End Property
    Public ReadOnly Property ShortReferenceWithVersion As String
        Get
            Return ShortReference & " [" & mLibrary.CurrentVersion.Name & "]"
        End Get
    End Property
    Public ReadOnly Property LongReferenceWithVersion() As String
        Get
            Return LongReference & " [" & mLibrary.CurrentVersion.Name & "]"
        End Get
    End Property

    Sub BuildBookChapterVerseIndex()
        On Error Resume Next
        If mBookNumber > 0 And mChapterNumber > 0 And mVerseNumber > 0 Then
            mBookChapterVerseIndex = Format(mBookNumber, "00") & Format(mChapterNumber, "000") & Format(mVerseNumber, "000")

            If mLibrary.CurrentVersion.Verses.ContainsKey(mBookChapterVerseIndex) Then
                mVerseText = mLibrary.CurrentVersion.Verses(mBookChapterVerseIndex)
            Else
                'this is a bad reference. set the verse number to be 0
                mVerseNumber = 0
                mBookChapterVerseIndex = ""
            End If
        End If
    End Sub
End Class
