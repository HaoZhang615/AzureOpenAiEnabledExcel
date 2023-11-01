# AzureOpenAiEnabledExcel
This repo is about adding a custom function in Excel file to enable API call to Azure OpenAI Service. 

## prerequrisites:
1. Access to an Azure OpenAI Service instance as well as its API endpoint and API key
2. Authorization in your organization to enable developer capabilities in Excel.

Code for your VBA module creation:
########### copy from here to the end #########################
Function AOAIGPT(userPrompt As String) As String

Dim apiKey As String
Dim apiEndpoint As String
Dim modelDeploymentName As String
Dim systemMessage As String
Dim temperature As Single
Dim maxTokens As Integer
Dim URL As String

' Set the API parameters from the "parameters" sheet
apiEndpoint = Worksheets("Parameters").Cells(1, "B").Value
apiKey = Worksheets("Parameters").Cells(2, "B").Value
modelDeploymentName = Worksheets("Parameters").Cells(3, "B").Value
systemMessage = Worksheets("Parameters").Cells(4, "B").Value
temperature = Worksheets("Parameters").Cells(5, "B").Value
maxTokens = Worksheets("Parameters").Cells(6, "B").Value

Set req = New MSXML2.ServerXMLHTTP60

' Create the API request
URL = apiEndpoint & "openai/deployments/" & modelDeploymentName & "/chat/completions?api-version=2023-07-01-preview"

req.Open "POST", URL, False
req.setRequestHeader "Content-Type", "application/json"
req.setRequestHeader "api-key", apiKey
    
' Create the JSON request payload with user prompt and other parameters
Dim jsonRequest As String
jsonRequest = "{""messages"": [{""role"": ""system"",""content"": """ & systemMessage & """},{""role"": ""user"",""content"": """ & userPrompt & """}],""temperature"": " & temperature & ",""top_p"": 0.95,""frequency_penalty"": 0, ""presence_penalty"": 0,""max_tokens"":" & maxTokens & ",""stop"": null}"

req.send jsonRequest

' Check for a successful response (you should add more error handling)
If req.Status = 200 Then
    ' Parse the JSON response and extract the completion text
    Dim jsonResponse As Dictionary
    Set jsonResponse = JsonConverter.ParseJson(req.responseText)
    ' Extract the "content" from the first choice
    AOAIGPT = jsonResponse("choices")(1)("message")("content")
Else
    ' Handle the error (e.g., return an error message)
    AOAIGPT = "Error: " & req.Status & " - " & req.statusText
End If

End Function
