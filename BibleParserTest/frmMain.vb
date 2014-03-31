Imports BibleParser
Imports System.Text

Public Class frmMain
    Private mLibrary As New cLibrary

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnParse.Click
        Dim sb As New StringBuilder
        mLibrary.BibleReference.ParseReference(txtReference.Text, False)

        If mLibrary.BibleReference.Verses IsNot Nothing Then
            For Each v In mLibrary.BibleReference.Verses
                sb.AppendLine(v.LongReference)
            Next
        End If
        txtResults.Text = sb.ToString
    End Sub

    Private Sub TextBox1_Enter(sender As Object, e As EventArgs) Handles txtReference.Enter
        txtReference.SelectAll()
    End Sub

    Private Sub btnUseTopsFile_Click(sender As Object, e As EventArgs) Handles btnUseTopsFile.Click
        Dim getFile As Boolean = False
        Dim topsPath As String = ""
        Dim topsOutPath As String = ""
        Dim rowcount As Long
        Dim outFile As System.IO.StreamWriter
        Dim inFile As System.IO.StreamReader
        Dim line As String
        Dim i As Long = 0
        If My.Settings.TopsTxtPath = "" Then
            getFile = True
        ElseIf FileIO.FileSystem.FileExists(My.Settings.TopsTxtPath) Then
            topsPath = My.Settings.TopsTxtPath
        Else
            getFile = True
        End If

        If getFile Then
            Dim fd As New OpenFileDialog
            fd.CheckFileExists = True
            fd.CheckPathExists = True
            fd.DefaultExt = "txt"
            fd.FileName = "Tops.txt"
            fd.Title = "Please locate the Tops.txt file."

            fd.ShowDialog()

            topsPath = fd.FileName
        End If

        If Not FileIO.FileSystem.FileExists(topsPath) Then
            Exit Sub
        End If
        My.Settings.TopsTxtPath = topsPath

        topsOutPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(topsPath), "TopsOut.txt")

        'read the file to get row count
        rowcount = System.IO.File.ReadAllLines(topsPath).Length
        pgMain.Minimum = 1
        pgMain.Maximum = 10000 'rowcount
        'pgMain.Width = Me.Width
        pgMain.Visible = True
        pgMain.MarqueeAnimationSpeed = 1000

        inFile = New IO.StreamReader(topsPath)
        outFile = New IO.StreamWriter(topsOutPath)

        Dim inRef As String

        While Not inFile.EndOfStream And i < 10000
            i = i + 1
            'Debug.Print(i)
            pgMain.Value = i
            line = inFile.ReadLine

            If i > 1 Then
                inRef = line.Substring(line.IndexOf(Chr(9)) + 1)
                inRef = inRef.Substring(0, inRef.IndexOf(Chr(9)))

                mLibrary.BibleReference.ParseReference(inRef, False)
                outFile.WriteLine(inRef & Chr(9) & mLibrary.BibleReference.VerseList)
            End If
        End While

        inFile.Close()
        outFile.Close()

        MessageBox.Show("The output file is located at: " & topsOutPath)
        'Process.Start("notepad.exe", topsOutPath)
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        pgMain.Width = Me.Width
    End Sub

    Private Sub frmMain_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        pgMain.Width = Me.Width
    End Sub
End Class
