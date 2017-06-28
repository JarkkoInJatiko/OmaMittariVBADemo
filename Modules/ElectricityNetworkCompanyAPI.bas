Attribute VB_Name = "ElectricityNetworkCompanyAPI"
'***********************************************************************
'* OmaMittari service platform demo - OmaMittariVBADemo.xlsm
'* Copyright (c) 2017, Jatiko Oy, email: info@jatiko.fi              Date:21.6.2017
'* You have to obey OmaMittari Terms of usage
'***********************************************************************

Option Explicit
Option Compare Text

'*** The uri of the RESTful OmaMittari consumer API
Const Network_URI_Start As String = "https://apigateway.omamittari.fi/verkkoyhtio/api/v1.1"

Dim clsJSON As New JSON
Dim objJSON As Object
Dim strJSON As String

'*** Get individual customer by customer id
Sub Network_GetCustomer()
Dim strCustomerId As String
strCustomerId = Sheets("ElectricityNetwork API v1.1").Cells(3, 2).Value
strJSON = JsonDataFetch(Network_URI_Start, "asiakas/" & strCustomerId)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    Debug.Print objJSON("Asiakastunnus")
    Debug.Print objJSON("Jakeluosoite")("Katuosoite")
End If
End Sub

'*** Customer list selection program
Sub Network_SelectCustomerListRequest()
Dim Choice As Variant
Choice = InputBox("Select request: " & vbNewLine & "1 = by id" & vbNewLine & "2 = by name or" & vbNewLine & "3 = all customers", "Request selection", 1)
Select Case Choice
    Case 1
        Network_GetCustomerList
    Case 2
        Network_GetCustomerByName
    Case 3
        Network_GetAllCustomers
    Case ""
    
    Case Else
        MsgBox "Wrong choice!", vbCritical, "Error"
End Select
End Sub

'*** Get list of all customers
Sub Network_GetAllCustomers()
Dim i As Integer
strJSON = JsonDataFetch(Network_URI_Start, "asiakkaat")

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
End Sub

'*** Get list of customers
Sub Network_GetCustomerList()
Dim i As Integer
Dim strCustomerIdList As String
strCustomerIdList = Sheets("ElectricityNetwork API v1.1").Cells(3, 5).Value & "," & Sheets("ElectricityNetwork API v1.1").Cells(4, 5).Value & "," & Sheets("ElectricityNetwork API v1.1").Cells(5, 5).Value
strJSON = JsonDataFetch(Network_URI_Start, "asiakkaat?lista=" & strCustomerIdList)

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
End Sub

'*** Get customer by name
Sub Network_GetCustomerByName()
Dim strName As String
Dim i As Integer
strName = Sheets("ElectricityNetwork API v1.1").Cells(7, 5).Value
strJSON = JsonDataFetch(Network_URI_Start, "asiakkaat?nimi=" & strName)

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
End Sub

'*** Get individual consumption place by consumption place id
Sub Network_GetConsumptionPlace()
Dim strConsumptionPlaceId As String
strConsumptionPlaceId = Sheets("ElectricityNetwork API v1.1").Cells(3, 8).Value
strJSON = JsonDataFetch(Network_URI_Start, "kayttopaikka/" & strConsumptionPlaceId)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    Debug.Print objJSON("Käyttöpaikkatunnus")
    Debug.Print objJSON("Osoite")("Katuosoite")
End If
End Sub

'*** Consumption place selection program
Sub Network_SelectConsumptionPlaceRequest()
Dim Choice As Variant
Choice = InputBox("Select request: " & vbNewLine & "1 = by id" & vbNewLine & "2 = by customer id or" & vbNewLine & "3 = by address", "Request selection", 1)
Select Case Choice
    Case 1
        Network_GetConsumptionPlaceListByIdList
    Case 2
        Network_GetConsumptionPlaceListByCustomerId
    Case 3
        Network_GetConsumptionPlaceListByAddress
    Case ""
    
    Case Else
        MsgBox "Wrong choice!", vbCritical, "Error"
End Select
End Sub

'*** Get list of consumption places by consumption place id list
Sub Network_GetConsumptionPlaceListByIdList()
Dim i As Integer
Dim strConsumptionPlaceIdList As String
strConsumptionPlaceIdList = Sheets("ElectricityNetwork API v1.1").Cells(3, 11).Value & "," & Sheets("ElectricityNetwork API v1.1").Cells(4, 11).Value & "," & Sheets("ElectricityNetwork API v1.1").Cells(5, 11).Value
strJSON = JsonDataFetch(Network_URI_Start, "kayttopaikat?lista=" & strConsumptionPlaceIdList)

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
End Sub

'*** Get list of consumption places by customer id
Sub Network_GetConsumptionPlaceListByCustomerId()
Dim i As Integer
Dim strCustomerId As String
strCustomerId = Sheets("ElectricityNetwork API v1.1").Cells(7, 11).Value
strJSON = JsonDataFetch(Network_URI_Start, "kayttopaikat?asiakas=" & strCustomerId)

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
End Sub

'*** Get list of consumption places by address
Sub Network_GetConsumptionPlaceListByAddress()
Dim i As Integer
Dim strStreetName As String
strStreetName = Sheets("ElectricityNetwork API v1.1").Cells(9, 11).Value
strJSON = JsonDataFetch(Network_URI_Start, "kayttopaikat?osoite=" & URLEncode(strStreetName))

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
End Sub

'*** Get individual distribution transformer by distribution transformer id
Sub Network_GetDistributionTransformer()
Dim strDistributionTransformerId As String
strDistributionTransformerId = Sheets("ElectricityNetwork API v1.1").Cells(3, 14).Value
strJSON = JsonDataFetch(Network_URI_Start, "jakelumuuntaja/" & strDistributionTransformerId)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    Debug.Print objJSON("Id")
    Debug.Print objJSON("Nimi")
End If
End Sub

'*** Distribution transformer selection program
Sub Network_SelectDistributionTransformerRequest()
Dim Choice As Variant
Choice = InputBox("Select request: " & vbNewLine & "1 = by id list" & vbNewLine & "2 = by transformer id" & vbNewLine & "3 = by medium voltage output id" & vbNewLine & "4 = by substation id or" & vbNewLine & "5 = all distribution transformers", "Request selection", 1)
Select Case Choice
    Case 1
        Network_GetDistributionTransformerListByIdList
    Case 2
        Network_GetDistributionTransformerListByTransformerId
    Case 3
        Network_GetDistributionTransformerListByMediumVoltageOutputId
    Case 4
        Network_GetDistributionTransformerListBySubstationId
    Case 5
        Network_GetAllDistributionTransformers
    Case ""
    Case Else
        MsgBox "Wrong choice!", vbCritical, "Error"
End Select
End Sub

'*** Get list of distribution transformers by distribution transformers id list
Sub Network_GetDistributionTransformerListByIdList()
Dim i As Integer
Dim strDistributionTransformerIdList As String
strDistributionTransformerIdList = Sheets("ElectricityNetwork API v1.1").Cells(3, 17).Value & "," & Sheets("ElectricityNetwork API v1.1").Cells(4, 17).Value
strJSON = JsonDataFetch(Network_URI_Start, "jakelumuuntajat?lista=" & strDistributionTransformerIdList)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON.Count
        Debug.Print objJSON(i)("Id")
        Debug.Print objJSON(i)("Nimi")
    Next i
End If
End Sub

'*** Get list of distribution transformers by transformer id
Sub Network_GetDistributionTransformerListByTransformerId()
Dim i As Integer
Dim strTransformerId As String
strTransformerId = Sheets("ElectricityNetwork API v1.1").Cells(6, 17).Value
strJSON = JsonDataFetch(Network_URI_Start, "jakelumuuntajat?muuntopiiri=" & strTransformerId)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON.Count
        Debug.Print objJSON(i)("Id")
        Debug.Print objJSON(i)("Nimi")
    Next i
End If
End Sub

'*** Get list of distribution transformers by medium voltage output id
Sub Network_GetDistributionTransformerListByMediumVoltageOutputId()
Dim i As Integer
Dim MediumVoltageOutputId As String
MediumVoltageOutputId = Sheets("ElectricityNetwork API v1.1").Cells(8, 17).Value
strJSON = JsonDataFetch(Network_URI_Start, "jakelumuuntajat?kjlahto=" & MediumVoltageOutputId)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON.Count
        Debug.Print objJSON(i)("Id")
        Debug.Print objJSON(i)("Nimi")
    Next i
End If
End Sub

'*** Get list of distribution transformers by substation id
Sub Network_GetDistributionTransformerListBySubstationId()
Dim i As Integer
Dim SubstationId As String
SubstationId = Sheets("ElectricityNetwork API v1.1").Cells(10, 17).Value
strJSON = JsonDataFetch(Network_URI_Start, "jakelumuuntajat?sahkoasema=" & SubstationId)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON.Count
        Debug.Print objJSON(i)("Id")
        Debug.Print objJSON(i)("Nimi")
    Next i
End If
End Sub

'*** Get list of all distribution transformers
Sub Network_GetAllDistributionTransformers()
Dim i As Integer

strJSON = JsonDataFetch(Network_URI_Start, "jakelumuuntajat")

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON.Count
        Debug.Print objJSON(i)("Id")
        Debug.Print objJSON(i)("Nimi")
    Next i
End If
End Sub

'*** Get individual connection point by connection point id
Sub Network_GetConnectionPoint()
Dim strConnectionPointId As String
strConnectionPointId = Sheets("ElectricityNetwork API v1.1").Cells(3, 20).Value
strJSON = JsonDataFetch(Network_URI_Start, "liittyma/" & strConnectionPointId)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    Debug.Print objJSON("Id")
    Debug.Print objJSON("Nimi")
End If
End Sub

'*** Connection point selection program
Sub Network_SelectConnectionPointRequest()
Dim Choice As Variant
Choice = InputBox("Select request: " & vbNewLine & "1 = by id list" & vbNewLine & "2 = by distribution transformer id" & vbNewLine & "3 = transformer id" & vbNewLine & "4 = by medium voltage output id" & vbNewLine & "5 = by substation id or" & vbNewLine & "6 = by coordinates" & vbNewLine & "7 = all distribution transformers", "Request selection", 1)
Select Case Choice
    Case 1
        Network_GetConnectionPointListByIdList
    Case 2
        Network_GetConnectionPointListByDistributionTransformerId
    Case 3
        Network_GetConnectionPointListByTransformerId
    Case 4
        Network_GetConnectionPointListByMediumVoltageOutputId
    Case 5
        Network_GetConnectionPointListBySubstationId
    Case 6
        Network_GetConnectionPointListByCoordinates
    Case 7
        Network_GetAllConnectionPoints
    Case ""
    Case Else
        MsgBox "Wrong choice!", vbCritical, "Error"
End Select
End Sub

'*** Get list of connection points by distribution transformers id list
Sub Network_GetConnectionPointListByIdList()
Dim i As Integer
Dim strConnectionPointIdList As String
strConnectionPointIdList = Sheets("ElectricityNetwork API v1.1").Cells(3, 23).Value & "," & Sheets("ElectricityNetwork API v1.1").Cells(4, 23).Value & "," & Sheets("ElectricityNetwork API v1.1").Cells(5, 23).Value
strJSON = JsonDataFetch(Network_URI_Start, "liittymat?lista=" & strConnectionPointIdList)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON.Count
        Debug.Print objJSON(i)("Liittymätunnus")
        Debug.Print objJSON(i)("Pääsulake")
    Next i
End If
End Sub

'*** Get list of connection points by distribution transformer id
Sub Network_GetConnectionPointListByDistributionTransformerId()
Dim i As Integer
Dim strDistributionTransformerId As String
strDistributionTransformerId = Sheets("ElectricityNetwork API v1.1").Cells(7, 23).Value
strJSON = JsonDataFetch(Network_URI_Start, "liittymat?jakelumuuntaja=" & strDistributionTransformerId)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON.Count
        Debug.Print objJSON(i)("Liittymätunnus")
        Debug.Print objJSON(i)("Pääsulake")
    Next i
End If
End Sub

'*** Get list of connection points by transformer id
Sub Network_GetConnectionPointListByTransformerId()
Dim i As Integer
Dim strTransformerId As String
strTransformerId = Sheets("ElectricityNetwork API v1.1").Cells(9, 23).Value
strJSON = JsonDataFetch(Network_URI_Start, "liittymat?muuntopiiri=" & strTransformerId)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON.Count
        Debug.Print objJSON(i)("Liittymätunnus")
        Debug.Print objJSON(i)("Pääsulake")
    Next i
End If
End Sub

'*** Get list of connection points by medium voltage output id
Sub Network_GetConnectionPointListByMediumVoltageOutputId()
Dim i As Integer
Dim MediumVoltageOutputId As String
MediumVoltageOutputId = Sheets("ElectricityNetwork API v1.1").Cells(11, 23).Value
strJSON = JsonDataFetch(Network_URI_Start, "liittymat?kjlahto=" & MediumVoltageOutputId)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON.Count
        Debug.Print objJSON(i)("Liittymätunnus")
        Debug.Print objJSON(i)("Pääsulake")
    Next i
End If
End Sub

'*** Get list of connection points by substation id
Sub Network_GetConnectionPointListBySubstationId()
Dim i As Integer
Dim SubstationId As String
SubstationId = Sheets("ElectricityNetwork API v1.1").Cells(13, 23).Value
strJSON = JsonDataFetch(Network_URI_Start, "liittymat?sahkoasema=" & SubstationId)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON.Count
        Debug.Print objJSON(i)("Liittymätunnus")
        Debug.Print objJSON(i)("Pääsulake")
    Next i
End If
End Sub

'*** Get list of connection points by coordinates
Sub Network_GetConnectionPointListByCoordinates()
Dim i As Integer
Dim CoordinateList As String
CoordinateList = Sheets("ElectricityNetwork API v1.1").Cells(15, 23).Value
strJSON = JsonDataFetch(Network_URI_Start, "liittymat?aluehaku=" & CoordinateList)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON.Count
        Debug.Print objJSON(i)("Liittymätunnus")
        Debug.Print objJSON(i)("Pääsulake")
    Next i
End If
End Sub

'*** Get list of all connection points
Sub Network_GetAllConnectionPoints()
Dim i As Integer

strJSON = JsonDataFetch(Network_URI_Start, "liittymat")

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To objJSON.Count
        Debug.Print objJSON(i)("Liittymätunnus")
        Debug.Print objJSON(i)("Pääsulake")
    Next i
End If
End Sub

'*** Measurement set selection program
Sub Network_SelectMeasurementSetRequest()
Dim Choice As Variant
Choice = InputBox("Select request: " & vbNewLine & "1 = accurate report" & vbNewLine & "2 = day report " & vbNewLine & "3 = week report" & vbNewLine & "4 = month report" & vbNewLine & "5 = year report", "Request selection", 1)
Select Case Choice
    Case 1
        Network_GetMeasurementSetAccurateReport
    Case 2
        Network_GetMeasurementSetDayReport
    Case 3
        Network_GetMeasurementSetWeekReport
    Case 4
        Network_GetMeasurementSetMonthReport
    Case 5
        Network_GetMeasurementSetYearReport
    Case ""
    
    Case Else
        MsgBox "Wrong choice!", vbCritical, "Error"
End Select
End Sub

'*** Get measurement set accurate report
Sub Network_GetMeasurementSetAccurateReport()
Dim i As Integer
Dim strStartDate As String, strEndDate As String, Target As String, Id As String, URI As String
Target = Sheets("ElectricityNetwork API v1.1").Cells(3, 26).Value
Id = Sheets("ElectricityNetwork API v1.1").Cells(5, 26).Value
strStartDate = Sheets("ElectricityNetwork API v1.1").Cells(7, 26).Value
strEndDate = Sheets("ElectricityNetwork API v1.1").Cells(8, 26).Value
URI = "mittaussarja/" & Target & "/" & Id & "?alku=" & strStartDate & "&loppu=" & strEndDate
strJSON = JsonDataFetch(Network_URI_Start, URI)

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
End Sub

'*** Get measurement set day report
Sub Network_GetMeasurementSetDayReport()
Dim i As Integer
Dim strDate As String, Target As String, Id As String, URI As String
Target = Sheets("ElectricityNetwork API v1.1").Cells(3, 26).Value
Id = Sheets("ElectricityNetwork API v1.1").Cells(5, 26).Value
strDate = Sheets("ElectricityNetwork API v1.1").Cells(10, 26).Value
URI = "mittaussarja/" & Target & "/" & Id & "?pvm=" & strDate
strJSON = JsonDataFetch(Network_URI_Start, URI)

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
End Sub

'*** Get measurement set week report
Sub Network_GetMeasurementSetWeekReport()
Dim i As Integer
Dim strWeek As String, strYear As String, Target As String, Id As String, URI As String
Target = Sheets("ElectricityNetwork API v1.1").Cells(3, 26).Value
Id = Sheets("ElectricityNetwork API v1.1").Cells(5, 26).Value
strWeek = Sheets("ElectricityNetwork API v1.1").Cells(12, 26).Value
strYear = Sheets("ElectricityNetwork API v1.1").Cells(16, 26).Value
URI = "mittaussarja/" & Target & "/" & Id & "?viikko=" & strWeek & "&vuosi=" & strYear
strJSON = JsonDataFetch(Network_URI_Start, URI)

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
End Sub

'*** Get measurement set month report
Sub Network_GetMeasurementSetMonthReport()
Dim i As Integer
Dim strMonth As String, strYear As String, Target As String, Id As String, URI As String
Target = Sheets("ElectricityNetwork API v1.1").Cells(3, 26).Value
Id = Sheets("ElectricityNetwork API v1.1").Cells(5, 26).Value
strMonth = Sheets("ElectricityNetwork API v1.1").Cells(14, 26).Value
strYear = Sheets("ElectricityNetwork API v1.1").Cells(16, 26).Value
URI = "mittaussarja/" & Target & "/" & Id & "?kuukausi=" & strMonth & "&vuosi=" & strYear
strJSON = JsonDataFetch(Network_URI_Start, URI)

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
End Sub

'*** Get measurement set year report
Sub Network_GetMeasurementSetYearReport()
Dim i As Integer
Dim strYear As String, Target As String, Id As String, URI As String
Target = Sheets("ElectricityNetwork API v1.1").Cells(3, 26).Value
Id = Sheets("ElectricityNetwork API v1.1").Cells(5, 26).Value
strYear = Sheets("ElectricityNetwork API v1.1").Cells(16, 26).Value
URI = "mittaussarja/" & Target & "/" & Id & "?vuosi=" & strYear
strJSON = JsonDataFetch(Network_URI_Start, URI)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To 10 'objJSON("Mittausjaksot").Count
        Debug.Print objJSON("Mittausjaksot")(i)("aika")
        Debug.Print objJSON("Mittausjaksot")(i)("sähkömittaus")("Pätöteho")
    Next i
End If
End Sub

'*** Measurement set sum selection program
Sub Network_SelectMeasurementSetSumRequest()
Dim Choice As Variant
Choice = InputBox("Select request: " & vbNewLine & "1 = accurate report" & vbNewLine & "2 = day report " & vbNewLine & "3 = week report" & vbNewLine & "4 = month report" & vbNewLine & "5 = year report", "Request selection", 1)
Select Case Choice
    Case 1
        Network_GetMeasurementSetSumAccurateReport
    Case 2
        Network_GetMeasurementSetSumDayReport
    Case 3
        Network_GetMeasurementSetSumWeekReport
    Case 4
        Network_GetMeasurementSetSumMonthReport
    Case 5
        Network_GetMeasurementSetSumYearReport
    Case ""
    
    Case Else
        MsgBox "Wrong choice!", vbCritical, "Error"
End Select
End Sub

'*** Get measurement set sum accurate report
Sub Network_GetMeasurementSetSumAccurateReport()
Dim i As Integer
Dim strStartDate As String, strEndDate As String, Target As String, List As String, URI As String
Target = Sheets("ElectricityNetwork API v1.1").Cells(3, 29).Value
List = Sheets("ElectricityNetwork API v1.1").Cells(5, 29).Value
strStartDate = Sheets("ElectricityNetwork API v1.1").Cells(7, 29).Value
strEndDate = Sheets("ElectricityNetwork API v1.1").Cells(8, 29).Value
URI = "mittaussarja/" & Target & "?alku=" & strStartDate & "&lista=" & List & "&loppu=" & strEndDate
strJSON = JsonDataFetch(Network_URI_Start, URI)

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
End Sub

'*** Get measurement set sum day report
Sub Network_GetMeasurementSetSumDayReport()
Dim i As Integer
Dim strDate As String, Target As String, List As String, URI As String
Target = Sheets("ElectricityNetwork API v1.1").Cells(3, 29).Value
List = Sheets("ElectricityNetwork API v1.1").Cells(5, 29).Value
strDate = Sheets("ElectricityNetwork API v1.1").Cells(10, 29).Value
URI = "mittaussarja/" & Target & "?lista=" & List & "&pvm=" & strDate
strJSON = JsonDataFetch(Network_URI_Start, URI)

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
End Sub

'*** Get measurement set sum week report
Sub Network_GetMeasurementSetSumWeekReport()
Dim i As Integer
Dim strWeek As String, strYear As String, Target As String, List As String, URI As String
Target = Sheets("ElectricityNetwork API v1.1").Cells(3, 29).Value
List = Sheets("ElectricityNetwork API v1.1").Cells(5, 29).Value
strWeek = Sheets("ElectricityNetwork API v1.1").Cells(12, 29).Value
strYear = Sheets("ElectricityNetwork API v1.1").Cells(16, 29).Value
URI = "mittaussarja/" & Target & "?lista=" & List & "&viikko=" & strWeek & "&vuosi=" & strYear
strJSON = JsonDataFetch(Network_URI_Start, URI)

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
End Sub

'*** Get measurement set sum month report
Sub Network_GetMeasurementSetSumMonthReport()
Dim i As Integer
Dim strMonth As String, strYear As String, Target As String, List As String, URI As String
Target = Sheets("ElectricityNetwork API v1.1").Cells(3, 29).Value
List = Sheets("ElectricityNetwork API v1.1").Cells(5, 29).Value
strMonth = Sheets("ElectricityNetwork API v1.1").Cells(14, 29).Value
strYear = Sheets("ElectricityNetwork API v1.1").Cells(16, 29).Value
URI = "mittaussarja/" & Target & "?kuukausi=" & strMonth & "&lista=" & List & "&vuosi=" & strYear
strJSON = JsonDataFetch(Network_URI_Start, URI)

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
End Sub

'*** Get measurement set sum year report
Sub Network_GetMeasurementSetSumYearReport()
Dim i As Integer
Dim strYear As String, Target As String, List As String, URI As String
Target = Sheets("ElectricityNetwork API v1.1").Cells(3, 29).Value
List = Sheets("ElectricityNetwork API v1.1").Cells(5, 29).Value
strYear = Sheets("ElectricityNetwork API v1.1").Cells(16, 29).Value
URI = "mittaussarja/" & Target & "?lista=" & List & "&vuosi=" & strYear
strJSON = JsonDataFetch(Network_URI_Start, URI)

'Echo json-string
MsgBox strJSON, vbInformation, "strJSON"

'Convert json string to object and Debug Print some values
Set objJSON = clsJSON.parse(strJSON)
If Not objJSON Is Nothing Then
    For i = 1 To 10 'objJSON("Mittausjaksot").Count
        Debug.Print objJSON("Mittausjaksot")(i)("aika")
        Debug.Print objJSON("Mittausjaksot")(i)("sähkömittaus")("Pätöteho")
    Next i
End If
End Sub

