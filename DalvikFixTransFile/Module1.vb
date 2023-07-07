Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

Module Module1

    Sub Main()

        Dim radcount As Integer = 0
        Dim radmodcount As Integer = 0
        Dim myWriter As StreamWriter = New StreamWriter("C:\kunder\dalvik\card_transactions februari 2023-g.csv", True, Encoding.GetEncoding(1252))
        Using sr As StreamReader = New StreamReader("C:\kunder\dalvik\card_transactions februari 2023.csv", Encoding.GetEncoding(1252))
            sr.ReadLine()
            Dim CSVParser As Regex = New Regex(",(?=(?:[^""]*""[^""]*"")*(?![^""]*""))")
            'Dim CSVParser As Regex = New Regex("(\S+)")
            While (Not sr.EndOfStream)
                radcount += 1
                Dim line As String = sr.ReadLine
                If line.Length > 0 Then
                    Dim Fields As String() = CSVParser.Split(line)

                    If Fields.Length = 30 Then
                        If Fields(21).Substring(0, 4) = "2003" Then
                            Dim dag As String = Fields(21).Substring(8, 2)
                            radmodcount += 1

                            Fields(21) = Fields(23).Substring(0, 10)
                            If Fields(5) = "" Then
                                Select Case dag
                                    Case "01", "02", "03"
                                        Fields(5) = "23.72"
                                    Case "05", "06"
                                        Fields(5) = "23.19"
                                End Select
                            Else
                                Dim a As String = ""
                            End If
                        End If
                    End If
                    Dim newline As String = ""
                    For Each f As String In Fields
                        newline &= f & ","
                    Next
                    newline = newline.Substring(0, newline.Length - 1)
                    WriteToFile(myWriter, newline)
                End If
            End While
        End Using

    End Sub

    Private Sub WriteToFile(ByVal sw As System.IO.StreamWriter, ByVal text As String)
        sw.WriteLine(text)
    End Sub
End Module
