Option Explicit

Public FirstDate, LastDate As Date
Public WeeklyRecords() As New PeriodRecord
Public MonthlyRecords() As New PeriodRecord
Public ServiceRecords() As New ProcRecord
Public ServiceRecordCount As Long
Public NumWeeks, NumMonths, NumProviders, NumLocations, NumServiceLines, NumServiceTypes As Integer



Sub Main()
    Import
    Parse
    SumUpWeekMonths
    'Dashboard
    'FindUniques
    'Cleanup

End Sub

Sub Import()
    'Sort
    Sheets("Raw").Activate
    Sheets("Raw").Columns("A:XFD").Sort key1:=Range("C:C"), order1:=xlAscending, Header:=xlYes
End Sub

Sub Parse()
    Dim i As Long
    
    FirstDate = Date
    
    ' Make headers for new columns
    Sheets("Raw").Cells(1, 57) = "Revenue"
    Sheets("Raw").Cells(1, 58) = "Service Type"
    Sheets("Raw").Cells(1, 59) = "Service Line"
    ServiceRecordCount = Sheets("Raw").Cells.SpecialCells(xlCellTypeLastCell).Row
    
    ReDim ServiceRecords(ServiceRecordCount - 2)
    
    For i = 0 To ServiceRecordCount - 2
        'Track first and last dates in data set
        If (Sheets("Raw").Cells(i + 2, 3) < FirstDate) Then
            FirstDate = Sheets("Raw").Cells(i + 2, 3)
        ElseIf (Sheets("Raw").Cells(i + 2, 3) > LastDate) Then
            LastDate = Sheets("Raw").Cells(i + 2, 3)
        End If
        
        'Add revenue and categorization to line items
        If (Sheets("Raw").Cells(i + 2, 15) = 0) Then
            Sheets("Raw").Cells(i + 2, 57) = Cells(i + 2, 11) + Cells(i + 2, 12)
        Else
            Sheets("Raw").Cells(i + 2, 57) = Cells(i + 2, 10) * Application.WorksheetFunction.VLookup(Cells(i + 2, 4), Sheets("Service Key").Range("A:I"), 8, False)
        End If
        Sheets("Raw").Cells(i + 2, 58) = Application.WorksheetFunction.VLookup(Cells(i + 2, 4), Sheets("Service Key").Range("A:I"), 3, False)
        Sheets("Raw").Cells(i + 2, 59) = Application.WorksheetFunction.VLookup(Cells(i + 2, 4), Sheets("Service Key").Range("A:I"), 4, False)
        
        'Scrape pertinent info
        ServiceRecords(i).ServiceDate = Sheets("Raw").Cells(i + 2, 3)
        ServiceRecords(i).AttendingProvider = Sheets("Raw").Cells(i + 2, 46)
        ServiceRecords(i).location = Sheets("Raw").Cells(i + 2, 8)
        ServiceRecords(i).Revenue = Sheets("Raw").Cells(i + 2, 57)
        ServiceRecords(i).serviceType = Sheets("Raw").Cells(i + 2, 58)
        ServiceRecords(i).serviceLine = Sheets("Raw").Cells(i + 2, 59)
    Next i
    
End Sub

Sub SumUpWeekMonths()
    
    Dim i, week, mnth, location, provider, servLine, servType As Long
    Dim recRevenue As Currency
    
    ' Up front date calculations
    While (WorksheetFunction.Weekday(FirstDate) <> vbSaturday)
        FirstDate = FirstDate - 1
    Wend
    
    NumMonths = DateDiff("m", FirstDate, LastDate, vbSaturday)
    NumWeeks = DateDiff("ww", FirstDate, LastDate, vbSaturday)
    
    ReDim WeeklyRecords(NumWeeks + 1)
    ReDim MonthlyRecords(NumMonths)
    
    For i = 0 To NumMonths
        MonthlyRecords(i).BeginDate = CDate(Month(DateAdd("m", i, FirstDate)) & "/1/" & Year(DateAdd("m", i, FirstDate)))
        MonthlyRecords(i).EndDate = CDate(DateAdd("d", -1, DateAdd("m", 1, MonthlyRecords(i).BeginDate)))
    Next
    
    For i = 0 To NumWeeks
        WeeklyRecords(i).BeginDate = CDate(DateAdd("w", i, FirstDate))
        WeeklyRecords(i).EndDate = CDate(DateAdd("d", -1, DateAdd("w", 1, WeeklyRecords(i).BeginDate)))
    Next
    
    For i = 2 To ServiceRecordCount
        ' Calculate indices
        week = DateDiff("ww", FirstDate, ServiceRecords(i).ServiceDate, vbSaturday)
        mnth = DateDiff("m", FirstDate, ServiceRecords(i).ServiceDate, vbSaturday)
        recRevenue = ServiceRecords(i).Revenue
        location = LocationIndex(ServiceRecords(i).location)
        provider = ProviderIndex(ServiceRecords(i).AttendingProvider)
        servLine = ServiceLineIndex(ServiceRecords(i).serviceLine)
        servType = ServiceTypeIndex(ServiceRecords(i).serviceType)
        
        ' Calculate weekly numbers
        WeeklyRecords(week).TotalRevenue = WeeklyRecords(week).TotalRevenue + recRevenue
        WeeklyRecords(week).Locations(location).LocationRevenue = WeeklyRecords(week).Locations(location).LocationRevenue + recRevenue
        WeeklyRecords(week).ProviderRevenue(provider) = WeeklyRecords(week).ProviderRevenue(provider) + recRevenue
        WeeklyRecords(week).ServiceLineRevenue(servLine) = WeeklyRecords(week).ServiceLineRevenue(servLine) + recRevenue
        WeeklyRecords(week).ServiceTypeRevenue(servType) = WeeklyRecords(week).ServiceTypeRevenue(servType) + recRevenue
        
        ' Calculate monthly numbers
        MonthlyRecords(mnth).TotalRevenue = MonthlyRecords(mnth).TotalRevenue + recRevenue
        MonthlyRecords(mnth).LocationRevenue(location) = MonthlyRecords(mnth).LocationRevenue(location) + recRevenue
        MonthlyRecords(mnth).ProviderRevenue(provider) = MonthlyRecords(mnth).ProviderRevenue(provider) + recRevenue
        MonthlyRecords(mnth).ServiceLineRevenue(servLine) = MonthlyRecords(mnth).ServiceLineRevenue(servLine) + recRevenue
        MonthlyRecords(mnth).ServiceTypeRevenue(servType) = MonthlyRecords(mnth).ServiceTypeRevenue(servType) + recRevenue
    Next i
End Sub

Sub Dashboard()
    Dim i, j, k As Integer
    Dim firstMonth, lastMonth, lastWeekStart As Date
    Dim firstMonthCell As Range
    Dim testDate1, testDate2 As Date
    
    Dim test1, test2, test3, test4 As Integer
        
    ' Find First/Last Week/Month
    firstMonth = CDate(Month(FirstDate) & "/1/" & Year(FirstDate))
    lastMonth = CDate(Month(LastDate) & "/1/" & Year(LastDate))
    lastWeekStart = DateAdd("ww", NumWeeks, FirstDate)
    
    NumLocations = 13
    
    ' Index goes to 100 b/c unsure of # of months in data set
    Sheets("DashboardMonthly").Activate
    For i = 0 To 130
        If (IsDate(Worksheets("DashboardMonthly").Cells(4, 4 + i))) Then
            testDate1 = CDate(Worksheets("DashboardMonthly").Cells(4, 4 + i).Value)
            For j = 0 To NumMonths
                testDate2 = CDate(MonthlyRecords(j).BeginDate)
                If (testDate1 = testDate2) Then
                    For k = 0 To NumLocations - 1
                        test2 = LocationIndex(Worksheets("DashboardMonthly").Cells(6 + k, 2))
                        test1 = MonthlyRecords(j).LocationRevenue(test2)
                        Worksheets("DashboardMonthly").Cells(6 + k, 4 + i).Value = test1
                    Next
                End If
            Next
        End If
    Next
    
    ' Weekly Dashboard
    Sheets("DashboardWeekly").Activate
    For i = 0 To 130 * 52
        If (IsDate(Worksheets("DashboardWeekly").Cells(4, 4 + i))) Then
            testDate1 = CDate(Worksheets("DashboardWeekly").Cells(4, 4 + i).Value)
            For j = 0 To NumWeeks
                testDate2 = CDate(WeeklyRecords(j).BeginDate)
                If (testDate1 = testDate2) Then
                    For k = 0 To NumLocations - 1
                        test2 = LocationIndex(Worksheets("DashboardWeekly").Cells(6 + k, 2))
                        test1 = WeeklyRecords(j).LocationRevenue(test2)
                        Worksheets("DashboardWeekly").Cells(6 + k, 4 + i).Value = test1
                    Next
                End If
            Next
        End If
    Next
    
    
    Calculate

End Sub

Sub FindUniques()
    Dim UniqueProviders, UniqueLocations, UniqueTypes, UniqueLines As Collection
    Dim source As Range
'    Set source = Sheets("Raw").Range(Cells(2, 46).Value, Cells(2, ServiceRecordCount).Value)
    
    'Unique Providers
    Set source = Sheets("Raw").Range("AT2:AT" & ServiceRecordCount)
    Set UniqueProviders = GetUniqueValues(source.Value)
     
    'Unique Locations
    Set source = Sheets("Raw").Range("H2:H" & ServiceRecordCount)
    Set UniqueLocations = GetUniqueValues(source.Value)
    
    'Unique ServiceTypes
    Set source = Sheets("Raw").Range("BF2:BF" & ServiceRecordCount)
    Set UniqueTypes = GetUniqueValues(source.Value)
     
    'Unique ServiceLines
    Set source = Sheets("Raw").Range("BG2:BG" & ServiceRecordCount)
    Set UniqueLines = GetUniqueValues(source.Value)
     
    'Print Unique Lists
    Dim it
    Debug.Print "\n"
    Debug.Print "\n"
    Debug.Print "****THESE ARE PROVIDERS****"
    Debug.Print "\n"
    Debug.Print "\n"
    For Each it In UniqueProviders
        Debug.Print it
    Next
    Debug.Print "\n"
    Debug.Print "\n"
    Debug.Print "****THESE ARE LOCATIONS****"
    Debug.Print "\n"
    Debug.Print "\n"
    For Each it In UniqueLocations
        Debug.Print it
    Next
    Debug.Print "\n"
    Debug.Print "\n"
    Debug.Print "****THESE ARE SERVICE TYPES****"
    Debug.Print "\n"
    Debug.Print "\n"
    For Each it In UniqueTypes
        Debug.Print it
    Next
    Debug.Print "\n"
    Debug.Print "\n"
    Debug.Print "****THESE ARE SERVICE LINES****"
    Debug.Print "\n"
    Debug.Print "\n"
    For Each it In UniqueLines
        Debug.Print it
    Next
    Debug.Print "\n"
    Debug.Print "\n"
    Debug.Print "****END OF SERVICE LINES****"
    Debug.Print "\n"
    Debug.Print "\n"
End Sub

Private Sub CommandButton1_Click()
    Main
End Sub

Public Function GetUniqueValues(ByVal values As Variant) As Collection
    Dim result As Collection
    Dim cellValue As Variant
    Dim cellValueTrimmed As String

    Set result = New Collection
    Set GetUniqueValues = result

    On Error Resume Next

    For Each cellValue In values
        cellValueTrimmed = Trim(cellValue)
        If cellValueTrimmed = "" Then GoTo NextValue
        result.Add cellValueTrimmed, cellValueTrimmed
NextValue:
    Next cellValue

    On Error GoTo 0
End Function

Public Function LocationIndex(location As String) As Integer
    Select Case location
        ' AICR
        Case "AICR"
            LocationIndex = 0
        ' Bee Cave
        Case "BCAVE"
            LocationIndex = 1
        ' CentralAustin
        Case "CNTRL"
            LocationIndex = 2
        ' DrippingSprings
        Case "DRPSP"
            LocationIndex = 3
        ' NorthAustin
        Case "N AUS"
            LocationIndex = 4
        ' Pflugerville
        Case "PFLUG"
            LocationIndex = 5
        ' SteinerRanch
        Case "STEIN"
            LocationIndex = 6
        ' SanAntonio
        Case "SanAntonio"
            LocationIndex = 7
        ' BatonRouge
        Case "BATR "
            LocationIndex = 8
        Case "BATR"
            LocationIndex = 8
        ' BossierCity
        Case "BOSSR"
            LocationIndex = 9
        ' Lafayette
        Case "LAFYT"
            LocationIndex = 10
        ' OldMetarie
        Case "OMET"
            LocationIndex = 11
        Case "OMET "
            LocationIndex = 11
        ' Shreveport
        Case "SHVPT"
            LocationIndex = 12
        Case ""
            LocationIndex = 13
        ' Default
        Case Else
            LocationIndex = 99
    End Select
End Function

Public Function ProviderIndex(provider As String) As Integer
    Select Case provider
        Case "REED, KELLIE"
            ProviderIndex = 0
        Case "Morrison, Kayla"
            ProviderIndex = 1
        Case "BOWEN, AMY"
            ProviderIndex = 2
        Case "SADEGHIAN, AZEEN"
            ProviderIndex = 3
        Case "MAMELAK, ADAM"
            ProviderIndex = 4
        Case "CARRASCO, DANIEL"
            ProviderIndex = 5
        Case "EVANS, ASHLEY"
            ProviderIndex = 6
        Case "North, Nurse"
            ProviderIndex = 7
        Case "JOHNSTON, EMILY"
            ProviderIndex = 8
        Case "GEWIRTZMAN, ARON"
            ProviderIndex = 9
        Case "LAIN, EDWARD"
            ProviderIndex = 10
        Case "VICKERS, JENNIFER"
            ProviderIndex = 11
        Case "SANTOS, SELINA"
            ProviderIndex = 12
        Case "PRATHER, CHAD"
            ProviderIndex = 13
        Case "JORDAN, JENNIFER"
            ProviderIndex = 14
        Case "GRAY, MISTY"
            ProviderIndex = 15
        Case "Balke, Jenna"
            ProviderIndex = 16
        Case "Winkenwerder, Shelley"
            ProviderIndex = 17
        Case "VIAL, DONNA"
            ProviderIndex = 18
        Case "HANSON, MIRIAM"
            ProviderIndex = 19
        Case "Pennington, Angela"
            ProviderIndex = 20
        Case "MCBURNEY, ELIZABETH"
            ProviderIndex = 21
        Case "TUCKER, LYNN"
            ProviderIndex = 22
        Case "ADA, NURSE"
            ProviderIndex = 23
        Case "WAGUESPACK-LABICHE, JENNIFER"
            ProviderIndex = 24
        Case "Lain, Nurse"
            ProviderIndex = 25
        Case "Rodrigues, Emily"
            ProviderIndex = 26
        Case "CARRINGTON, PATRICK"
            ProviderIndex = 27
        Case "DONNES, GRETCHEN"
            ProviderIndex = 28
        Case "Bee Cave, Nurse"
            ProviderIndex = 29
        Case "SAENZ, GILBERTO"
            ProviderIndex = 30
        Case "FARRIS, PATRICIA"
            ProviderIndex = 31
        Case "raymond, Bridget"
            ProviderIndex = 32
        Case "TUREGANO, MAMINA"
            ProviderIndex = 33
        Case "MANSOURI, BOBBAK"
            ProviderIndex = 34
        Case "PROVIDER NOT SELECTED"
            ProviderIndex = 35
        Case ""
            ProviderIndex = 36
        Case Else
            ProviderIndex = 99
    End Select
End Function

Public Function ServiceTypeIndex(serviceType As String) As Integer
    Select Case serviceType
        ' Cosmetic - Other
        Case "Cosmetic - Other"
            ServiceTypeIndex = 0
        ' Medical
        Case "Medical"
            ServiceTypeIndex = 1
        ' Product
        Case "Product"
            ServiceTypeIndex = 2
        ' Cosmetic - Neurotoxin
        Case "Cosmetic - Neurotoxin"
            ServiceTypeIndex = 3
        ' ?
        Case "?"
            ServiceTypeIndex = 4
        Case ""
            ServiceTypeIndex = 5
        ' Default
        Case Else
            ServiceTypeIndex = 99
    End Select
End Function

Public Function ServiceLineIndex(serviceLine As String) As Integer
    Select Case serviceLine
        ' Cosmetic -Other
        Case "Cosmetic - Other"
            ServiceLineIndex = 0
        ' Consults
        Case "Consults"
            ServiceLineIndex = 1
        ' General Derm
        Case "General Derm"
            ServiceLineIndex = 2
        ' HCPC
        Case "HCPC"
            ServiceLineIndex = 3
        ' Product
        Case "Product"
            ServiceLineIndex = 4
        ' Botox
        Case "Botox"
            ServiceLineIndex = 5
        ' Resurfacing
        Case "Resurfacing"
            ServiceLineIndex = 6
        ' Dysport
        Case "Dysport"
            ServiceLineIndex = 7
        ' Filler
        Case "Filler"
            ServiceLineIndex = 8
        ' Lab
        Case "Lab"
            ServiceLineIndex = 9
        Case "lab"
            ServiceLineIndex = 9
        ' MOHS
        Case "MOHS"
            ServiceLineIndex = 10
        ' Surgical Excisions
        Case "Surgical Excisions"
            ServiceLineIndex = 11
        ' Cool Sculpting
        Case "Cool Sculpting"
            ServiceLineIndex = 12
        ' LHR
        Case "LHR"
            ServiceLineIndex = 13
        ' Micro -Needling
        Case "Micro-Needling"
            ServiceLineIndex = 14
        ' IPL
        Case "IPL"
            ServiceLineIndex = 15
        ' Hydrofacial
        Case "Hydrofacial"
            ServiceLineIndex = 16
        ' ?
        Case "?"
            ServiceLineIndex = 17
        ' Peel
        Case "Peel"
            ServiceLineIndex = 18
        ' Dermaplaning
        Case "Dermaplaning"
            ServiceLineIndex = 19
        ' Contouring
        Case "Contouring"
            ServiceLineIndex = 20
        ' VBEAM
        Case "VBEAM"
            ServiceLineIndex = 21
        ' Excel V
        Case "Excel V"
            ServiceLineIndex = 22
        ' PRP
        Case "PRP"
            ServiceLineIndex = 23
        ' Radiation
        Case "Radiation"
            ServiceLineIndex = 24
        ' Microderm Abrasion
        Case "Microderm Abrasion"
            ServiceLineIndex = 25
        ' Skin Tightening
        Case "Skin Tightening"
            ServiceLineIndex = 26
        ' Miradry
        Case "Miradry"
            ServiceLineIndex = 27
        ' Sclerotherapy
        Case "Sclerotherapy"
            ServiceLineIndex = 28
        ' Tattoo Removal
        Case "Tattoo Removal"
            ServiceLineIndex = 29
        ' Microneedling
        Case "Microneedling"
            ServiceLineIndex = 30
        ' CO2 Laser
        Case "CO2 Laser"
            ServiceLineIndex = 31
        ' Xeomin
        Case "Xeomin"
            ServiceLineIndex = 32
        ' General Derm
        Case "General Derm"
            ServiceLineIndex = 33
        ' No Show
        Case "No Show"
            ServiceLineIndex = 34
        ' Default
        Case ""
            ServiceLineIndex = 35
        Case Else
            ServiceLineIndex = 99
    End Select
End Function


