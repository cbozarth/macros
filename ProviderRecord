Option Explicit

Private Const NUM_SERVICES = 6

Public ServiceRevenue As Currency
Private m_ServiceTypes() As New ServiceTypeRecord

Public Property Get Services(index As Single)
    Services = m_Services(index)
End Property

Public Property Set Services(index As Single, ServiceSet As ServiceTypeRecord)
    m_Services(index) = ServiceSet
End Property

Private Sub Class_Initialize()
    ServiceRevenue = 0
    
    ReDim m_ServiceTypes(NUM_SERVICES)
    Dim serv_index As Single
    For serv_index = 0 To NUM_SERVICES
        m_ServiceTypes(serv_index) = 0
    Next
End Sub


'Public Type ProviderRecord
'    ProviderRevenue As Single
'    ServiceTypes(6) As ServiceTypeRecord
'End Type

