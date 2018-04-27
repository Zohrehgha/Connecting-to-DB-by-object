Public Class rateClass
    'Class Variables 
    Private CcountryID As Integer
    Private CcountryName As String
    Private Crate As Integer
    'step 2 Creating a constructor 
    'Public Sub New()
    '    CcountryID = 0
    '    CcountryName = ""
    '    Crate = 0
    'End Sub
    Public Property countryID As Integer
        Get
            'when you want the outside world to know about the value of
            'this property 
            countryID = CcountryID
        End Get
        Set(value As Integer)
            'when you want to put a value in this property from the world outside 
            CcountryID = value
        End Set
    End Property
    Public Property countryName As String
        Get
            'when you want the outside world to know about the value of
            'this property 
            countryName = CcountryName
        End Get
        Set(value As String)
            'when you want to put a value in this property from the world outside 
            CcountryName = value
        End Set
    End Property
    Public Property rate As Integer
        Get
            'when you want the outside world to know about the value of
            'this property 
            rate = Crate
        End Get
        Set(value As Integer)
            'when you want to put a value in this property from the world outside 
            Crate = value
        End Set
    End Property
End Class
