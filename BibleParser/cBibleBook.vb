Imports System.ComponentModel
Public Class cBibleBook
    Implements INotifyPropertyChanged
    Public Enum enTestament
        OldTestament = 0
        NewTestament = 1
    End Enum

#Region "Public Events"
    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged
#End Region 'Public Events

    Private mChapters As Integer
    Private mAbbreviations As String
    Private mNumber As Integer
    Private mNumberString As String
    Private mLongName As String
    Private mAltLongName As String
    Private mLongNameLower As String
    Private mAltLongNameLower As String
    Private mLongNameNoSpace As String
    Private mAltLongNameNoSpace As String
    Private mShortName As String
    Private mTestament As enTestament

    ''' <summary>    '''      ''' </summary>    ''' <value></value>    ''' <returns>String</returns>    ''' <remarks></remarks>    Public Property Chapters() As Integer
        Get
            Return mChapters
        End Get
        Set(ByVal value As Integer)
            mChapters = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Chapters"))
        End Set
    End Property
    ''' <summary>    '''      ''' </summary>    ''' <value></value>    ''' <returns>String</returns>    ''' <remarks></remarks>    Public Property Abbreviations() As String
        Get
            Return mAbbreviations
        End Get
        Set(ByVal value As String)
            mAbbreviations = value.Trim
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Abbreviations"))
        End Set
    End Property
    ''' <summary>    '''      ''' </summary>    ''' <value></value>    ''' <returns>String</returns>    ''' <remarks></remarks>    Public Property Number() As Integer
        Get
            Return mNumber
        End Get
        Set(ByVal value As Integer)
            mNumber = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Number"))
        End Set
    End Property
    ''' <summary>    '''      ''' </summary>    ''' <value></value>    ''' <returns>String</returns>    ''' <remarks></remarks>    Public Property NumberString() As String
        Get
            Return mNumberString
        End Get
        Set(ByVal value As String)
            mNumberString = value.Trim
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("NumberString"))
        End Set
    End Property
    ''' <summary>    '''      ''' </summary>    ''' <value></value>    ''' <returns>String</returns>    ''' <remarks></remarks>    Public Property LongName() As String
        Get
            Return mLongName
        End Get
        Set(ByVal value As String)
            mLongName = value.Trim
            mLongNameLower = value.ToLower
            'mLongNameNoSpace = Replace(mLongNameLower, " ", "")
            mLongNameNoSpace = mLongNameLower.Replace(" ", "")
            If IsNumeric(value.Substring(0, 1)) Then
                'mAltLongName = String(Left(InValue, 1), "i") & Mid(InValue, 2)
                'mAltLongNameNoSpace = Replace(mAltLongNameLower, " ", "")
                mAltLongName = New String("i", value.Substring(0, 1)) & value.Substring(1)
                mAltLongNameLower = mAltLongName.ToLower
                mAltLongNameNoSpace = mAltLongNameLower.Replace(" ", "")
            Else
                mAltLongName = mLongName
                mAltLongNameLower = mLongNameLower
                mAltLongNameNoSpace = mLongNameNoSpace
            End If
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("LongName"))
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("LongNameLower"))
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("LongNameNoSpace"))
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("AltLongName"))
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("AltLongNameLower"))
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("AltLongNameNoSpace"))
        End Set
    End Property
    ''' <summary>    '''      ''' </summary>    ''' <value></value>    ''' <returns>String</returns>    ''' <remarks></remarks>    Public Property AltLongName() As String
        Get
            Return mAltLongName
        End Get
        Set(ByVal value As String)
            mAltLongName = value.Trim
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("AltLongName"))
        End Set
    End Property
    ''' <summary>    '''      ''' </summary>    ''' <value></value>    ''' <returns>String</returns>    ''' <remarks></remarks>    Public Property LongNameLower() As String
        Get
            Return mLongNameLower
        End Get
        Set(ByVal value As String)
            mLongNameLower = value.Trim
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("LongNameLower"))
        End Set
    End Property
    ''' <summary>    '''      ''' </summary>    ''' <value></value>    ''' <returns>String</returns>    ''' <remarks></remarks>    Public Property AltLongNameLower() As String
        Get
            Return mAltLongNameLower
        End Get
        Set(ByVal value As String)
            mAltLongNameLower = value.Trim
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("AltLongNameLower"))
        End Set
    End Property
    ''' <summary>    '''      ''' </summary>    ''' <value></value>    ''' <returns>String</returns>    ''' <remarks></remarks>    Public Property LongNameNoSpace() As String
        Get
            Return mLongNameNoSpace
        End Get
        Set(ByVal value As String)
            mLongNameNoSpace = value.Trim
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("LongNameNoSpace"))
        End Set
    End Property
    ''' <summary>    '''      ''' </summary>    ''' <value></value>    ''' <returns>String</returns>    ''' <remarks></remarks>    Public Property AltLongNameNoSpace() As String
        Get
            Return mAltLongNameNoSpace
        End Get
        Set(ByVal value As String)
            mAltLongNameNoSpace = value.Trim
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("AltLongNameNoSpace"))
        End Set
    End Property
    ''' <summary>    '''      ''' </summary>    ''' <value></value>    ''' <returns>String</returns>    ''' <remarks></remarks>    Public Property ShortName() As String
        Get
            Return mShortName
        End Get
        Set(ByVal value As String)
            mShortName = value.Trim
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("ShortName"))
        End Set
    End Property
    ''' <summary>    '''      ''' </summary>    ''' <value></value>    ''' <returns>String</returns>    ''' <remarks></remarks>    Public Property Testament() As enTestament
        Get
            Return mTestament
        End Get
        Set(ByVal value As enTestament)
            mTestament = value
        End Set
    End Property

    Public Function HasName(BookName As String) As Boolean
        Dim hn As Boolean
        Dim lbn As String
        Dim lenBN As Integer
        On Error Resume Next

        lbn = BookName.ToLower
        lenBN = BookName.Length
        'hn = InStr(mAbbreviations, "_" & lbn & "_") > -1
        hn = mAbbreviations.IndexOf("_" & lbn & "_") > -1

        If Not hn Then
            'see if it is the start of the name
            If lenBN <= mLongName.Length Then
                'hn = (VBA.Left(mLongNameLower, lenBN) = lbn)
                'hn = mLongNameLower.Substring(0, lenBN) = lbn
                hn = Left(mLongNameLower, lenBN) = lbn
            End If
        End If

        If Not hn Then
            'see if it is the start of the name without spaces
            If lenBN <= mLongName.Length Then
                'hn = (VBA.Left(mLongNameNoSpace, lenBN) = lbn)
                'hn = mLongNameNoSpace.Substring(0, lenBN) = lbn
                hn = Left(mLongNameNoSpace, lenBN) = lbn
            End If
        End If

        If Not hn Then
            'see if it is the start of the name
            If lenBN <= mLongName.Length Then
                'hn = (VBA.Left(mAltLongNameLower, lenBN) = lbn)
                'hn = mAltLongNameLower.Substring(0, lenBN) = lbn
                hn = Left(mAltLongNameLower, lenBN) = lbn
            End If
        End If

        If Not hn Then
            'see if it is the start of the name without spaces
            If lenBN <= mLongName.Length Then
                'special case for 1 samuel and isaiah. If bookname = "isa" assume this is isaiah
                If BookName = "isa" And mAltLongNameNoSpace = "isamuel" Then
                    'nothing to do
                Else
                    'hn = mAltLongNameNoSpace.Substring(0, lenBN) = lbn
                    hn = Left(mAltLongNameNoSpace, lenBN) = lbn
                End If
            End If
        End If

        Return hn
    End Function

End Class
