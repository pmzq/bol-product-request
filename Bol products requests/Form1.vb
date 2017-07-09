Imports System.IO
Imports System.Net
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class Form1

    Public API_key, Bol_ID, Product_id As String
    Dim request As HttpWebRequest
    Dim response As HttpWebResponse = Nothing
    Dim reader As StreamReader
    Dim Title, Description, key, Prijs As String
    Dim Rating As Long
    Dim Image_url As String = ""
    Dim Image_urls As Object
    Dim Output As String = ""
    Dim Url, Alt_tag As String
    Dim Bol_tag, Bol_subid As String

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click



        Dim Counter As Integer

        Dim Url_boxes As New List(Of TextBox)() From {
            txtUrl1,
            txtUrl2,
            txtUrl3,
            txtUrl4,
            txtUrl5,
            txtUrl6,
            txtUrl7,
            txtUrl7,
            txtUrl8,
            txtUrl9,
            txtUrl10
        }

        Dim Alt_boxes As New List(Of TextBox)() From {
            txtAlt1,
            txtAlt2,
            txtAlt3,
            txtAlt4,
            txtAlt5,
            txtAlt6,
            txtAlt7,
            txtAlt7,
            txtAlt8,
            txtAlt9,
            txtAlt10
        }

        Dim Subid_boxes As New List(Of TextBox)() From {
            txtSubid1,
            txtSubid2,
            txtSubid3,
            txtSubid4,
            txtSubid5,
            txtSubid6,
            txtSubid7,
            txtSubid7,
            txtSubid8,
            txtSubid9,
            txtSubid10
        }

        API_key = "66744DFB331F4D9E838E38D48A2052D6"
        Bol_ID = "13464"
        Bol_tag = txtTag.Text

        Counter = 1

        For Each tb As TextBox In Url_boxes
            Url = tb.Text
            If Url <> "" Then
                Alt_tag = Alt_boxes(Counter).Text
                Bol_subid = Subid_boxes(Counter).Text

                Dim parts As String() = Url.Split(New Char() {"/"c})

                ' Loop through result strings with For Each.
                Title = parts(5)
                Product_id = parts(6)

                Call Get_product()
                Call Compose_html()
            Else
            End If
        Next





    End Sub

    Sub Get_product()



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
        'txtOutput.Text = If(jResults("products") Is Nothing, "", jResults("products").ToString())
        Product_id = jResults("products")(0)("id")
        'Title = jResults("products")(0)("title")
        Description = jResults("products")(0)("longDescription")
        Prijs = jResults("products")(0)("price")
        Rating = jResults("products")(0)("rating")
        Image_urls = jResults("products")(0)("images")

        For Each image In Image_urls

            key = image("key")
            If key = "XL" Then
                Image_url = image("url")
            ElseIf key = "L" Then
                Image_url = image("url")
            End If

        Next


        'Url = jResults("urls")("DESKTOP")

    End Sub

    Sub Compose_html()

        Output = Output & "&nbsp;&nbsp;<span style=""font-size: 14pt;""><strong><a href=""https://partnerprogramma.bol.com/click/click?p=1&t=url&s=" & Bol_ID & "&f=TXL&url=https%3A%2F%2Fwww.bol.com%2Fnl%2Fp%2F" & Title & "%2F" & Product_id & "&name=" & Bol_tag & "&subid=" & Bol_subid & """>" & Title & "</a></strong></span>
                
                <a href=""" & Image_url & """><img Class=""size-thumbnail wp-image-276 alignleft"" src=""" & Image_url & """ alt=""" & Alt_tag & """ width=""150"" height=""150"" /></a>" & Description & "

                <a href=""https://partnerprogramma.bol.com/click/click?p=1&t=url&s=" & Bol_ID & "&f=TXL&url=https%3A%2F%2Fwww.bol.com%2Fnl%2Fp%2F" & Title & "%2F" & Product_id & "%2F%3FsuggestionType%3Dbrowse%23product_reviews&name=" & Bol_tag & "&subid=Reviews"">Reviews Bol.com</a> [usr " & Rating / 10 & "]"

        txtOutput.Text = Output

    End Sub

End Class




