Imports System.IO
Public Class Form1
    Public unit As Integer 'From 0
    Public whichline As Integer
    Public wordList() As String
    Public english() As String
    Public chinese() As String
    Public mark() As Integer
    Public whichLineInUnit As String
    'Public mistakesIndex() As Integer

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "单词复习器 v0.1"
        Label3.Text = "made by banic ，感谢zjs提供的单词表" & vbCrLf & "All Rights Reserved"
        Label1.Text = ""
        Label2.Text = ""
        Dim f1 As StreamWriter = New StreamWriter("mistakes.txt", True, System.Text.Encoding.Default)
        f1.Close()
        Dim txt As String = ReadTxt("WordList.txt")
        AddVbCrLfAtTheEndIfWithout(txt)
        wordList = OrderIntoList(txt)
        GetEnglishAndChineseFromArray(wordList)
        Dim mark() As Integer = DivideIntoUnites(wordList)
        For i = 0 To mark.LongCount - 1
            ListBox1.Items.Add("第 " & i + 1 & " 单元")
        Next
    End Sub
    Public Sub AddVbCrLfAtTheEndIfWithout(ByRef txt As String)
        If Mid(txt, Len(txt) - 1) <> vbCrLf Then txt = txt & vbCrLf
    End Sub
    Public Function ReadTxt(ByVal address As String)
        Dim reader As String
        reader = File.ReadAllText(address, System.Text.Encoding.Default)
        Return reader
    End Function
    Public Function OrderIntoList(ByVal txt As String)
        Dim strArray() As String
        strArray = Split(txt, vbCrLf)
        Return strArray
    End Function

    Public Sub GetEnglishAndChineseFromArray(ByVal wordList() As String)
        'Dim englishValue() As String
        Dim englishStrs As String = ""
        For i = 0 To wordList.LongCount - 1
            ReDim Preserve english(i)
            Dim strs As String = wordList(i)
            'englishValue(i) = WordListValue(i)
            For ii = 1 To Len(strs)
                Dim str As String = ""
                str = Mid(strs, ii, 1)
                If Asc(str) < 0 Then
                    Exit For
                Else
                    englishStrs = englishStrs & str
                End If
            Next
            If englishStrs <> "" Then
                english(i) = Mid(englishStrs, 1, Len(englishStrs) - 1)
                englishStrs = ""
            End If
        Next

        Dim chineseStrs As String = ""
        For i = 0 To wordList.LongCount - 1
            ReDim Preserve chinese(i)
            Dim strs As String = wordList(i)
            'englishValue(i) = WordListValue(i)
            If strs <> "" Then
                For ii = 1 To Len(strs)
                    Dim str As String = ""
                    str = Mid(strs, ii, 1)
                    If Asc(str) < 0 Then
                        chineseStrs = Mid(strs, ii)
                        Exit For
                    Else

                    End If
                Next
                chinese(i) = chineseStrs
                chineseStrs = ""
            End If
        Next
    End Sub
    Public Function DivideIntoUnites(ByVal strArray() As String)
        Dim ii As Int16 = 0
        For i = 0 To strArray.LongCount - 1
            If strArray(i) = "" Then
                ReDim Preserve mark(ii)
                mark(ii) = i
                ii = ii + 1
            End If
        Next
        Return mark
    End Function

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        unit = ListBox1.SelectedIndex
        If unit <> ListBox1.Items.Count - 1 Then
            whichline = mark(unit) + 1
            Label1.Text = GetWord(whichline)
            TextBox1.Text = ""
            Label2.Text = ""
            RefreshTrackbar1()
        Else
            ListBox1.SelectedIndex = unit - 1
        End If
    End Sub
    Public Function GetWord(ByVal whichline As Integer)
        Dim word As String
        word = chinese(whichline)
        Return word
    End Function

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Dim text As String = TextBox1.Text
        If IsThereAVbCrLf(text) Then
            If IsCorrected(text, whichline) Then
                whichline = whichline + 1
                If Not IsFinished(whichline) Then
                    Label1.Text = GetWord(whichline)
                    TrackBar1.Value = GetWhichLineInUnit(whichline) - 1
                Else

                    unit = unit + 1
                    If unit <> ListBox1.Items.Count - 1 Then
                        'Label1.Text = "The End. 按回车继续下一单元"
                        'RefreshTrackbar1()
                        ListBox1.SelectedIndex = unit
                    Else
                        Label1.Text = "The End"
                        unit = unit - 1
                    End If
                End If
                    TextBox1.Text = ""
                Label2.Text = ""

            Else
                If text <> "" Then
                    TextBox1.Text = Mid(text, 1, Len(text) - 1)
                Else
                    TextBox1.Text = ""
                End If
                TextBox1.SelectAll()
                Label2.Text = english(whichline)
                'RecordMistakes(whichline)
                WriteMistakes(whichline)
            End If
        End If

    End Sub
    Public Function IsThereAVbCrLf(ByVal text As String) As Boolean
        If InStr(text, vbCrLf) <> 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Function IsCorrected(ByVal word As String, ByVal i As Int32) As Boolean
        If word = english(i) & vbCrLf Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Function IsFinished(ByVal whichline As Integer) As Boolean
        Dim a As Boolean = False
        For i = 0 To mark.LongCount - 1
            If mark(i) = whichline Then
                a = True
            End If
        Next
        Return a
    End Function
    'Public Sub RecordMistakes(ByVal whichline As Integer)
    '    'Dim index As Integer
    '    'index = mistakesIndex.LongCount - 1
    '    'ReDim mistakesIndex(index + 1)
    '    'mistakesIndex(index + 1) = whichline
    'End Sub
    Public Sub WriteMistakes(ByVal whichline As Integer)
        Dim f1 As StreamWriter = New StreamWriter("mistakes.txt", True, System.Text.Encoding.Default)
        f1.Write(wordList(whichline) & vbCrLf)
        f1.Close()
    End Sub
    Public Function GetWhichLineInUnit(ByVal whichline As Integer)
        whichLineInUnit = whichline - mark(unit)
        Return whichLineInUnit
    End Function

    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        whichline = TrackBar1.Value + 1 + mark(unit)
        Label1.Text = GetWord(whichline)
        Label2.Text = ""
    End Sub

    Public Sub RefreshTrackbar1()
        TrackBar1.Value = 0
        TrackBar1.Maximum = mark(unit + 1) - mark(unit) - 2
    End Sub
End Class
