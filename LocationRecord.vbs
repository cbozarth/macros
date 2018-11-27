Option Explicit

Private Const NUM_PROVIDERS = 37

Public LocationRevenue As Currency
Private m_Providers(37) As New ProviderRecord

Public Property Get Providers(index As Single)
    Providers = m_Providers(index)
End Property

Public Property Set Providers(index As Single, ProviderSet As ProviderRecord)
    m_Providers(index) = ProviderSet
End Property

Private Sub Class_Initialize()
    LocationRevenue = 0
    
    ReDim m_Providers(NUM_PROVIDERS)
    Dim prov_index As Single
    For prov_index = 0 To NUM_PROVIDERS
        m_Providers(prov_index) = 0
    Next
End Sub

'Public Type LocationRecord
'    LocationRevenue As Single
'    Providers(37) As ProviderRecord
'End Type



'Public Property Get LocationRevenue()
'    LocationRevenue As Single
'    Providers(37) As ProviderRecord
'End Type

