Imports System.Net
Imports System.IO

Partial Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim myWebRequest As HttpWebRequest = Nothing
        Dim myWebResponse As HttpWebResponse = Nothing
        Try
            Dim sURL As String = "http://sample.onewaysms.com.au:xxxx/api.aspx"
            sURL = sURL & "?apiusername=" & HttpUtility.UrlEncode("123")
            sURL = sURL & "&apipassword=" & HttpUtility.UrlEncode("xyz")
            sURL = sURL & "&mobileno=" & HttpUtility.UrlEncode("6141234567")
            sURL = sURL & "&senderid=" & HttpUtility.UrlEncode("onewaysms")
            sURL = sURL & "&languagetype=" & "1"
            sURL = sURL & "&message=" & HttpUtility.UrlEncode("testing sms from api")
            myWebRequest = System.Net.WebRequest.Create(sURL)
            myWebResponse = myWebRequest.GetResponse()
            If myWebResponse.StatusCode = HttpStatusCode.OK Then
                Dim oStream As Stream = myWebResponse.GetResponseStream
                Dim oReader As StreamReader = New StreamReader(oStream)
                Dim sResult As String = oReader.ReadToEnd
                If Long.Parse(sResult) > 0 Then
                    Response.Write("success - MT ID :" & sResult)
                Else
                    Response.Write("fail - Error code :" & sResult)
                End If
            End If
        Catch ex As Exception
            Response.Write("Some issue happen")
        Finally
            If Not myWebResponse Is Nothing Then
                myWebResponse.Close()
            End If
        End Try
    End Sub
End Class
