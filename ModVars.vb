Module ModVars
    Public USERID As Integer = 0
    Public USERNAME As String = ""
    Public USERLEVEL As String = ""
    Public LOCALHOST As String = "192.168.10.248"
    Public PORT As String = "3306"
    Public DBUSER As String = "ivymae"
    Public DBPASSWORD As String = "ojt2022"
    Public DATABASE As String = "jorhsndb"
    'Public curUserlevel As String = ""
    'Public curUsername As String = ""

    Public Function AddSlashes(ByVal InputTxt As String) As String
        ' List of characters handled:
        ' \000 null
        ' \010 backspace
        ' \011 horizontal tab
        ' \012 new line
        ' \015 carriage return
        ' \032 substitute
        ' \042 double quote
        ' \047 single quote
        ' \134 backslash
        ' \140 grave accent

        Dim Result As String = InputTxt

        Try
            Result = System.Text.RegularExpressions.Regex.Replace(InputTxt, "[\000\010\011\012\015\032\042\047\134\140]", "\$0")
        Catch Ex As Exception
            ' handle any exception here
            Console.WriteLine(Ex.Message)
        End Try

        Return Result
    End Function
End Module
