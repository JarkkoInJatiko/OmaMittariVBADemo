'***********************************************************************
'* OmaMittari service platform VBA demo - OmaMittariVBADemo.xlsm
'* Copyright (c) 2017, Jatiko Oy, email: info@jatiko.fi  Date:4.10.2017
'* You have to obey OmaMittari Terms of Usage
'***********************************************************************

Option Explicit
Option Compare Text

'*** The uri of the RESTful OmaMittari consumer API
Const Customer_URI_Start As String = "https://apigateway.omamittari.fi/sahkoasiakas/api/v1.1"

Dim clsJSON As New JSON
Dim objJSON As Object
Dim strJSON As String

'*** Get individual customer by customer id
Sub Customer_GetCustomer()
Dim strCustomerId As String

On Error GoTo ErrorHandler
strCustomerId = Sheets("Consumer API v1.1").Cells(3, 2).Value
strJSON = JsonDataFetch(Customer_URI_Start, "asiakas/" & strCustomerId)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    Debug.Print objJSON("Asiakastunnus")
    Debug.Print objJSON("Jakeluosoite")("Katuosoite")
End If

Exit Sub
ErrorHandler:
MsgBox "An error happened, debug the code for more information", vbCritical, "Error"
End Sub

'*** Customer list selection program
Sub Customer_SelectCustomerListRequest()
Dim Choice As Variant

Choice = InputBox("Select request: " & vbNewLine & "1 = by id" & vbNewLine & "2 = by name or" & vbNewLine & "3 = all customers", "Request selection", 1)
Select Case Choice
    Case 1
        Customer_GetCustomerList
    Case 2
        Customer_GetCustomerByName
    Case 3
        Customer_GetAllCustomers
    Case ""
    
    Case Else
        MsgBox "Wrong choice!", vbCritical, "Error"
End Select
End Sub

'*** Get list of all customers
Sub Customer_GetAllCustomers()
Dim i As Integer

On Error GoTo ErrorHandler
strJSON = JsonDataFetch(Customer_URI_Start, "asiakkaat")

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON.Count
        Debug.Print objJSON(i)("Asiakastunnus")
        Debug.Print objJSON(i)("Jakeluosoite")("Katuosoite")
    Next i
End If

Exit Sub
ErrorHandler:
MsgBox "An error happened, debug the code for more information", vbCritical, "Error"
End Sub

'*** Get list of customers
Sub Customer_GetCustomerList()
Dim i As Integer
Dim strCustomerIdList As String

On Error GoTo ErrorHandler
strCustomerIdList = Sheets("Consumer API v1.1").Cells(3, 5).Value & "," & Sheets("Consumer API v1.1").Cells(4, 5).Value
strJSON = JsonDataFetch(Customer_URI_Start, "asiakkaat?lista=" & strCustomerIdList)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON.Count
        Debug.Print objJSON(i)("Asiakastunnus")
        Debug.Print objJSON(i)("Jakeluosoite")("Katuosoite")
    Next i
End If

Exit Sub
ErrorHandler:
MsgBox "An error happened, debug the code for more information", vbCritical, "Error"
End Sub

'*** Get customer by name
Sub Customer_GetCustomerByName()
Dim strName As String
Dim i As Integer

On Error GoTo ErrorHandler
strName = Sheets("Consumer API v1.1").Cells(7, 5).Value
strJSON = JsonDataFetch(Customer_URI_Start, "asiakkaat?nimi=" & URLEncode(strName))

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON.Count
        Debug.Print objJSON(i)("Asiakastunnus")
        Debug.Print objJSON(i)("Nimi")
    Next i
End If

Exit Sub
ErrorHandler:
MsgBox "An error happened, debug the code for more information", vbCritical, "Error"
End Sub

'*** Get individual consumption place by consumption place id
Sub Customer_GetConsumptionPlace()
Dim strConsumptionPlaceId As String

On Error GoTo ErrorHandler
strConsumptionPlaceId = Sheets("Consumer API v1.1").Cells(3, 8).Value
strJSON = JsonDataFetch(Customer_URI_Start, "kayttopaikka/" & strConsumptionPlaceId)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    Debug.Print objJSON("Käyttöpaikkatunnus")
    Debug.Print objJSON("Osoite")("Katuosoite")
End If

Exit Sub
ErrorHandler:
MsgBox "An error happened, debug the code for more information", vbCritical, "Error"
End Sub

'*** Consumption place selection program
Sub Customer_SelectConsumptionPlaceRequest()
Dim Choice As Variant
Choice = InputBox("Select request: " & vbNewLine & "1 = by id" & vbNewLine & "2 = by customer id or" & vbNewLine & "3 = by address", "Request selection", 1)
Select Case Choice
    Case 1
        Customer_GetConsumptionPlaceListByIdList
    Case 2
        Customer_GetConsumptionPlaceListByCustomerId
    Case 3
        Customer_GetConsumptionPlaceListByAddress
    Case ""
    
    Case Else
        MsgBox "Wrong choice!", vbCritical, "Error"
End Select
End Sub

'*** Get list of consumption places by consumption place id list
Sub Customer_GetConsumptionPlaceListByIdList()
Dim i As Integer
Dim strConsumptionPlaceIdList As String

On Error GoTo ErrorHandler
strConsumptionPlaceIdList = Sheets("Consumer API v1.1").Cells(3, 11).Value & "," & Sheets("Consumer API v1.1").Cells(4, 11).Value
strJSON = JsonDataFetch(Customer_URI_Start, "kayttopaikat?lista=" & strConsumptionPlaceIdList)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON.Count
        Debug.Print objJSON(i)("Käyttöpaikkatunnus")
        Debug.Print objJSON(i)("Nimi")
    Next i
End If

Exit Sub
ErrorHandler:
MsgBox "An error happened, debug the code for more information", vbCritical, "Error"
End Sub

'*** Get list of consumption places by customer id
Sub Customer_GetConsumptionPlaceListByCustomerId()
Dim i As Integer
Dim strCustomerId As String

On Error GoTo ErrorHandler
strCustomerId = Sheets("Consumer API v1.1").Cells(7, 11).Value
strJSON = JsonDataFetch(Customer_URI_Start, "kayttopaikat?asiakas=" & strCustomerId)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON.Count
        Debug.Print objJSON(i)("Käyttöpaikkatunnus")
        Debug.Print objJSON(i)("Osoite")("Katuosoite")
    Next i
End If

Exit Sub
ErrorHandler:
MsgBox "An error happened, debug the code for more information", vbCritical, "Error"
End Sub

'*** Get list of consumption places by address
Sub Customer_GetConsumptionPlaceListByAddress()
Dim i As Integer
Dim strStreetName As String

On Error GoTo ErrorHandler
strStreetName = Sheets("Consumer API v1.1").Cells(9, 11).Value
strJSON = JsonDataFetch(Customer_URI_Start, "kayttopaikat?osoite=" & URLEncode(strStreetName))

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON.Count
        Debug.Print objJSON(i)("Käyttöpaikkatunnus")
        Debug.Print objJSON(i)("Osoite")("Katuosoite")
    Next i
End If

Exit Sub
ErrorHandler:
MsgBox "An error happened, debug the code for more information", vbCritical, "Error"
End Sub

'*** Measurement set selection program
Sub Customer_SelectMeasurementSetRequest()
Dim Choice As Variant
Choice = InputBox("Select request: " & vbNewLine & "1 = accurate report" & vbNewLine & "2 = day report " & vbNewLine & "3 = week report" & vbNewLine & "4 = month report" & vbNewLine & "5 = year report", "Request selection", 1)
Select Case Choice
    Case 1
        Customer_GetMeasurementSetAccurateReport
    Case 2
        Customer_GetMeasurementSetDayReport
    Case 3
        Customer_GetMeasurementSetWeekReport
    Case 4
        Customer_GetMeasurementSetMonthReport
    Case 5
        Customer_GetMeasurementSetYearReport
    Case ""
    
    Case Else
        MsgBox "Wrong choice!", vbCritical, "Error"
End Select
End Sub

'*** Get measurement set accurate report
Sub Customer_GetMeasurementSetAccurateReport()
Dim i As Integer
Dim strStartDate As String, strEndDate As String, Target As String, Id As String, URI As String

On Error GoTo ErrorHandler
Target = Sheets("Consumer API v1.1").Cells(3, 14).Value
Id = Sheets("Consumer API v1.1").Cells(5, 14).Value
strStartDate = Sheets("Consumer API v1.1").Cells(7, 14).Value
strEndDate = Sheets("Consumer API v1.1").Cells(8, 14).Value
URI = "mittaussarja/" & Target & "/" & Id & "?alku=" & strStartDate & "&loppu=" & strEndDate
strJSON = JsonDataFetch(Customer_URI_Start, URI)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON("Mittausjaksot").Count
        Debug.Print objJSON("Mittausjaksot")(i)("aika")
        Debug.Print objJSON("Mittausjaksot")(i)("sähkömittaus")("Pätöteho")
    Next i
End If

Exit Sub
ErrorHandler:
MsgBox "An error happened, debug the code for more information", vbCritical, "Error"
End Sub

'*** Get measurement set day report
Sub Customer_GetMeasurementSetDayReport()
Dim i As Integer
Dim strDate As String, Target As String, Id As String, URI As String

On Error GoTo ErrorHandler
Target = Sheets("Consumer API v1.1").Cells(3, 14).Value
Id = Sheets("Consumer API v1.1").Cells(5, 14).Value
strDate = Sheets("Consumer API v1.1").Cells(10, 14).Value
URI = "mittaussarja/" & Target & "/" & Id & "?pvm=" & strDate
strJSON = JsonDataFetch(Customer_URI_Start, URI)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON("Mittausjaksot").Count
        Debug.Print objJSON("Mittausjaksot")(i)("aika")
        Debug.Print objJSON("Mittausjaksot")(i)("sähkömittaus")("Pätöteho")
    Next i
End If

Exit Sub
ErrorHandler:
MsgBox "An error happened, debug the code for more information", vbCritical, "Error"
End Sub

'*** Get measurement set week report
Sub Customer_GetMeasurementSetWeekReport()
Dim i As Integer
Dim strWeek As String, strYear As String, Target As String, Id As String, URI As String

On Error GoTo ErrorHandler
Target = Sheets("Consumer API v1.1").Cells(3, 14).Value
Id = Sheets("Consumer API v1.1").Cells(5, 14).Value
strWeek = Sheets("Consumer API v1.1").Cells(12, 14).Value
strYear = Sheets("Consumer API v1.1").Cells(16, 14).Value
URI = "mittaussarja/" & Target & "/" & Id & "?viikko=" & strWeek & "&vuosi=" & strYear
strJSON = JsonDataFetch(Customer_URI_Start, URI)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON("Mittausjaksot").Count
        Debug.Print objJSON("Mittausjaksot")(i)("aika")
        Debug.Print objJSON("Mittausjaksot")(i)("sähkömittaus")("Pätöteho")
    Next i
End If

Exit Sub
ErrorHandler:
MsgBox "An error happened, debug the code for more information", vbCritical, "Error"
End Sub

'*** Get measurement set month report
Sub Customer_GetMeasurementSetMonthReport()
Dim i As Integer
Dim strMonth As String, strYear As String, Target As String, Id As String, URI As String

On Error GoTo ErrorHandler
Target = Sheets("Consumer API v1.1").Cells(3, 14).Value
Id = Sheets("Consumer API v1.1").Cells(5, 14).Value
strMonth = Sheets("Consumer API v1.1").Cells(14, 14).Value
strYear = Sheets("Consumer API v1.1").Cells(16, 14).Value
URI = "mittaussarja/" & Target & "/" & Id & "?kuukausi=" & strMonth & "&vuosi=" & strYear
strJSON = JsonDataFetch(Customer_URI_Start, URI)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON("Mittausjaksot").Count
        Debug.Print objJSON("Mittausjaksot")(i)("aika")
        Debug.Print objJSON("Mittausjaksot")(i)("sähkömittaus")("Pätöteho")
    Next i
End If

Exit Sub
ErrorHandler:
MsgBox "An error happened, debug the code for more information", vbCritical, "Error"
End Sub

'*** Get measurement set year report
Sub Customer_GetMeasurementSetYearReport()
Dim i As Integer
Dim strYear As String, Target As String, Id As String, URI As String

On Error GoTo ErrorHandler
Target = Sheets("Consumer API v1.1").Cells(3, 14).Value
Id = Sheets("Consumer API v1.1").Cells(5, 14).Value
strYear = Sheets("Consumer API v1.1").Cells(16, 14).Value
URI = "mittaussarja/" & Target & "/" & Id & "?vuosi=" & strYear
strJSON = JsonDataFetch(Customer_URI_Start, URI)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON("Mittausjaksot").Count
        Debug.Print objJSON("Mittausjaksot")(i)("aika")
        Debug.Print objJSON("Mittausjaksot")(i)("sähkömittaus")("Pätöteho")
    Next i
End If

Exit Sub
ErrorHandler:
MsgBox "An error happened, debug the code for more information", vbCritical, "Error"
End Sub

'*** Measurement set sum selection program
Sub Customer_SelectMeasurementSetSumRequest()
Dim Choice As Variant
Choice = InputBox("Select request: " & vbNewLine & "1 = accurate report" & vbNewLine & "2 = day report " & vbNewLine & "3 = week report" & vbNewLine & "4 = month report" & vbNewLine & "5 = year report", "Request selection", 1)
Select Case Choice
    Case 1
        Customer_GetMeasurementSetSumAccurateReport
    Case 2
        Customer_GetMeasurementSetSumDayReport
    Case 3
        Customer_GetMeasurementSetSumWeekReport
    Case 4
        Customer_GetMeasurementSetSumMonthReport
    Case 5
        Customer_GetMeasurementSetSumYearReport
    Case ""
    
    Case Else
        MsgBox "Wrong choice!", vbCritical, "Error"
End Select
End Sub

'*** Get measurement set sum accurate report
Sub Customer_GetMeasurementSetSumAccurateReport()
Dim i As Integer
Dim strStartDate As String, strEndDate As String, Target As String, List As String, URI As String

On Error GoTo ErrorHandler
Target = Sheets("Consumer API v1.1").Cells(3, 17).Value
List = Sheets("Consumer API v1.1").Cells(5, 17).Value
strStartDate = Sheets("Consumer API v1.1").Cells(7, 17).Value
strEndDate = Sheets("Consumer API v1.1").Cells(8, 17).Value
URI = "mittaussarja/" & Target & "?alku=" & strStartDate & "&lista=" & List & "&loppu=" & strEndDate
strJSON = JsonDataFetch(Customer_URI_Start, URI)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON("Mittausjaksot").Count
        Debug.Print objJSON("Mittausjaksot")(i)("aika")
        Debug.Print objJSON("Mittausjaksot")(i)("sähkömittaus")("Pätöteho")
    Next i
End If

Exit Sub
ErrorHandler:
MsgBox "An error happened, debug the code for more information", vbCritical, "Error"
End Sub

'*** Get measurement set sum day report
Sub Customer_GetMeasurementSetSumDayReport()
Dim i As Integer
Dim strDate As String, Target As String, List As String, URI As String

On Error GoTo ErrorHandler
Target = Sheets("Consumer API v1.1").Cells(3, 17).Value
List = Sheets("Consumer API v1.1").Cells(5, 17).Value
strDate = Sheets("Consumer API v1.1").Cells(10, 17).Value
URI = "mittaussarja/" & Target & "?lista=" & List & "&pvm=" & strDate
strJSON = JsonDataFetch(Customer_URI_Start, URI)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON("Mittausjaksot").Count
        Debug.Print objJSON("Mittausjaksot")(i)("aika")
        Debug.Print objJSON("Mittausjaksot")(i)("sähkömittaus")("Pätöteho")
    Next i
End If

Exit Sub
ErrorHandler:
MsgBox "An error happened, debug the code for more information", vbCritical, "Error"
End Sub

'*** Get measurement set sum week report
Sub Customer_GetMeasurementSetSumWeekReport()
Dim i As Integer
Dim strWeek As String, strYear As String, Target As String, List As String, URI As String

On Error GoTo ErrorHandler
Target = Sheets("Consumer API v1.1").Cells(3, 17).Value
List = Sheets("Consumer API v1.1").Cells(5, 17).Value
strWeek = Sheets("Consumer API v1.1").Cells(12, 17).Value
strYear = Sheets("Consumer API v1.1").Cells(16, 17).Value
URI = "mittaussarja/" & Target & "?lista=" & List & "&viikko=" & strWeek & "&vuosi=" & strYear
strJSON = JsonDataFetch(Customer_URI_Start, URI)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON("Mittausjaksot").Count
        Debug.Print objJSON("Mittausjaksot")(i)("aika")
        Debug.Print objJSON("Mittausjaksot")(i)("sähkömittaus")("Pätöteho")
    Next i
End If

Exit Sub
ErrorHandler:
MsgBox "An error happened, debug the code for more information", vbCritical, "Error"
End Sub

'*** Get measurement set sum month report
Sub Customer_GetMeasurementSetSumMonthReport()
Dim i As Integer
Dim strMonth As String, strYear As String, Target As String, List As String, URI As String

On Error GoTo ErrorHandler
Target = Sheets("Consumer API v1.1").Cells(3, 17).Value
List = Sheets("Consumer API v1.1").Cells(5, 17).Value
strMonth = Sheets("Consumer API v1.1").Cells(14, 17).Value
strYear = Sheets("Consumer API v1.1").Cells(16, 17).Value
URI = "mittaussarja/" & Target & "?kuukausi=" & strMonth & "&lista=" & List & "&vuosi=" & strYear
strJSON = JsonDataFetch(Customer_URI_Start, URI)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON("Mittausjaksot").Count
        Debug.Print objJSON("Mittausjaksot")(i)("aika")
        Debug.Print objJSON("Mittausjaksot")(i)("sähkömittaus")("Pätöteho")
    Next i
End If

Exit Sub
ErrorHandler:
MsgBox "An error happened, debug the code for more information", vbCritical, "Error"
End Sub

'*** Get measurement set sum year report
Sub Customer_GetMeasurementSetSumYearReport()
Dim i As Integer
Dim strYear As String, Target As String, List As String, URI As String

On Error GoTo ErrorHandler
Target = Sheets("Consumer API v1.1").Cells(3, 17).Value
List = Sheets("Consumer API v1.1").Cells(5, 17).Value
strYear = Sheets("Consumer API v1.1").Cells(16, 17).Value
URI = "mittaussarja/" & Target & "?lista=" & List & "&vuosi=" & strYear
strJSON = JsonDataFetch(Customer_URI_Start, URI)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON("Mittausjaksot").Count
        Debug.Print objJSON("Mittausjaksot")(i)("aika")
        Debug.Print objJSON("Mittausjaksot")(i)("sähkömittaus")("Pätöteho")
    Next i
End If

Exit Sub
ErrorHandler:
MsgBox "An error happened, debug the code for more information", vbCritical, "Error"
End Sub
