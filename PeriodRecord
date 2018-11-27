Option Explicit

Private Const NUM_LOCATIONS = 12

Public BeginDate As Date
Public EndDate As Date
Public TotalRevenue As Currency
Private m_Locations() As New LocationRecord

Public Property Get Locations(index As Single)
    Locations = m_Locations(index)
End Property

Public Property Set Locations(index As Single, LocationSet As LocationRecord)
    m_Locations(index) = LocationSet
End Property

Private Sub Class_Initialize()
    BeginDate = 0
    EndDate = 0
    ReDim m_Locations(NUM_LOCATIONS)
    Dim loc_index As Single
    For loc_index = 0 To NUM_LOCATIONS
        m_Locations(loc_index) = 0
    Next
End Sub

'Public m_BeginDate As Date
'Public m_EndDate As Date
'Public m_TotalRevenue As Currency

'Public Locations

'Public Property Get BeginDate() As Date
'    BeginDate = m_BeginDate
'End Property

'Public Property Get EndDate() As Date
'    EndDate = m_EndDate
'End Property

'Public Type PeriodRecord
'    BeginDate As Date
'    EndDate As Date
'    TotalRevenue As Currency
'    Locations(12) As LocationRecord
'End Type


