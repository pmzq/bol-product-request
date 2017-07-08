Imports System.IO
Imports System.Net
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim Url As String
        Dim API_key As String

        API_key = "66744DFB331F4D9E838E38D48A2052D6"
        Url = txtUrl1.Text

        Call get_product(API_key, Url)


    End Sub

    Sub get_product(API_key, Product_id)

        Dim request As HttpWebRequest
        Dim response As HttpWebResponse = Nothing
        Dim reader As StreamReader

        request = DirectCast(WebRequest.Create("https://api.bol.com/catalog/v4/products/" & Product_id & "?apikey=" & API_key & "&includeAttributes=true&format=json"), HttpWebRequest)

        response = DirectCast(request.GetResponse(), HttpWebResponse)
        reader = New StreamReader(response.GetResponseStream())

        Dim rawresp As String
        rawresp = reader.ReadToEnd()
        Dim tokenJson = JsonConvert.SerializeObject(rawresp)

        'Dim jsonResult = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(rawresp)
        'Dim firstItem = jsonResult.Item("products").Item(0).ToString()
        'MsgBox(firstItem)

        Dim jResults As Object = JObject.Parse(rawresp)
        txtOutput.Text = If(jResults("products") Is Nothing, "", jResults("products").ToString())
        MsgBox(jResults("products")(0)("id"))


    End Sub

End Class




