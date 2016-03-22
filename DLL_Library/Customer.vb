
Imports CSLib
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Runtime.Serialization.Formatters.Soap


Namespace IOTS
    <Serializable()>
    Public Class Customer
        Inherits Address
        Implements IValidator
        Private custId As Long
        Private firstName As String
        Private lastName As String
        Private streetAddress As String
        Private city As String
        Private province As String
        Private postalCode As String
        Private creditLimit As Double
        Private email As String
        Private count As Int16 = 0

        Public ReadOnly Property _custId() As Long
            Get
                Return custId
            End Get
        End Property
        Public Property _firstName() As String
            Get
                ' Return the value stored in the local variable. 
                Return firstName
            End Get
            Set(ByVal Value As String)
                ' Store the value in a local variable.
                Dim regName As String = "^[A-Za-z]+$"
                If IsValid(Value, regName) Then
                    firstName = Value
                    count = count + 1
                Else
                    Throw New OrderSystemExceptions("An exception has occurred, name should contain only upper case and lower case alphabets.")
                End If
            End Set
        End Property
        Public Property _lastName() As String
            Get
                ' Return the value stored in the local variable. 
                Return lastName
            End Get
            Set(ByVal Value As String)
                ' Store the value in a local variable.
                Dim regName As String = "^[A-Za-z]+$"
                If IsValid(Value, regName) Then
                    lastName = Value
                    count = count + 1
                Else
                    Throw New OrderSystemExceptions("An exception has occurred, name should contain only upper case and lower case alphabets.")
                End If
            End Set
        End Property
        Public Property _streetAddress() As String
            Get
                ' Return the value stored in the local variable. 
                Return streetAddress
            End Get
            Set(ByVal Value As String)
                ' Store the value in a local variable.
                Dim regStreet As String = "^[0-9].*$"
                If IsValid(Value, regStreet) Then
                    streetAddress = Value
                    count = count + 1
                Else
                    Throw New OrderSystemExceptions("An exception has occurred, Street address should start with a digit.")
                End If
            End Set
        End Property
        Public Property _city() As String
            Get
                ' Return the value stored in the local variable. 
                Return city
            End Get
            Set(ByVal Value As String)
                ' Store the value in a local variable.
                Dim regName As String = "^[a-z A-Z]*$"
                If IsValid(Value, regName) Then
                    city = Value
                    count = count + 1
                Else
                    Throw New OrderSystemExceptions("An exception has occurred, city name should contain only upper case and lower case alphabets.")
                End If
            End Set
        End Property
        Public Property _province() As String
            Get
                ' Return the value stored in the local variable. 
                Return province
            End Get
            Set(ByVal Value As String)
                ' Store the value in a local variable.
                Dim regProv As String = "^[A-Za-z]{2}$"
                If IsValid(Value, regProv) Then
                    province = Value
                    count = count + 1
                Else
                    Throw New OrderSystemExceptions("An exception has occurred, Province should be 2 alphabets.")
                End If
            End Set
        End Property
        Public Property _postalCode() As String
            Get
                ' Return the value stored in the local variable. 
                Return postalCode
            End Get
            Set(ByVal Value As String)
                ' Store the value in a local variable.
                Dim regPCode As String = "^([A-Za-z][0-9]){3}$"
                If IsValid(Value, regPCode) Then
                    postalCode = Value
                    count = count + 1
                Else
                    Throw New OrderSystemExceptions("An exception has occurred, Postal code should be of the form X1X1X1.")
                End If
            End Set
        End Property

        Public Property _email() As String
            Get
                ' Return the value stored in the local variable. 
                Return email
            End Get
            Set(ByVal Value As String)
                ' Store the value in a local variable.
                Dim regEmail1 As String = "^.+@.+\..+$"
                If IsValid(Value, regEmail1) Then
                    Dim subString As String = Right(Value, 4)
                    Dim ending As String = subString.Split(".")(1)
                    Dim regEnding As String = "com|ca|org|net|biz"
                    If IsValid(ending, regEnding) Then
                        email = Value
                        count = count + 1
                    Else
                        Throw New OrderSystemExceptions("An exception has occurred, email domains may only end in com/.ca/.org/.net/.biz/.")
                    End If

                Else
                    Throw New OrderSystemExceptions("An exception has occurred, email should have @ and .")
                End If
            End Set
        End Property

        Public Property _creditLimit() As Double
            Get
                ' Return the value stored in the local variable. 
                Return creditLimit
            End Get
            Set(ByVal Value As Double)
                ' Store the value in a local variable.

                If (Value > 0) Then
                    creditLimit = Value
                Else
                    Throw New OrderSystemExceptions("An exception has occurred, Credit can not be negative.")
                End If
            End Set
        End Property

        Sub New(ByVal custId, ByVal firstname, ByVal lastname, ByVal street, ByVal city, ByVal province, ByVal postalcode, ByVal credit, ByVal email)
            MyClass.custId = custId
            _firstName = firstname
            _lastName = lastname
            _streetAddress = street
            _city = city
            _province = province
            _postalCode = postalcode
            _creditLimit = credit
            _email = email
        End Sub


        'imlements the IsValid from IValidator
        Public Function IsValid(ByVal testData As String, ByVal regex As String) As Boolean Implements IValidator.IsValid
            Dim r As Regex = New Regex(regex, RegexOptions.IgnoreCase)
            Dim m As Match = r.Match(testData)
            Return m.Success
        End Function

        Public Overrides Function ToString() As String
            Return String.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}", custId, firstName, lastName, streetAddress, city, province, postalCode, creditLimit, email)
        End Function

        Public Function GetDetails() As String
            Dim result As String = Nothing
            'result = result + "Customer Infomation:" + vbCrLf
            result = result + "Customer Id: " & CStr(_custId) & vbCrLf
            result = result + "Customer Name: " + _firstName + " " + _lastName & vbCrLf
            result = result + "Address: " + _streetAddress + "," + _city + "," + _province + " " + _postalCode + vbCrLf
            result = result + "Email: " + _email + vbTab + "Credit Limit:" + CStr(_creditLimit)
            Return result
        End Function
    End Class

    <Serializable()>
    Public Class CustomerHolder
        Inherits ObjectHolder
        Public customerList As ArrayList
        Sub New(customers As List(Of Customer))
            customerList = New ArrayList()
            For i As Integer = 0 To customers.Count - 1
                customerList.Add(customers(i))
            Next
        End Sub
        Sub New()

        End Sub

        Public Overrides Sub loadFromFile(ByVal filename As String)

            Dim line As String = Nothing
            Dim lineData() As String
            Dim oneCust As Customer
            'empty the old list
            customerList = New ArrayList()
            Try


                'read every line and split
                Dim sr As StreamReader = New StreamReader(filename)
                'the first line is header
                line = sr.ReadLine()
                Do
                    Try
                        line = sr.ReadLine()
                        lineData = Split(line, ",")
                        If lineData.Length >= 7 Then
                            Console.WriteLine("Read record:" + line)
                            'save to array
                            oneCust = New DLL_Library.IOTS.Customer(lineData(0), lineData(1), lineData(2), lineData(3), lineData(4), lineData(5), lineData(6), lineData(7), lineData(8))
                            customerList.Add(oneCust)
                        End If
                    Catch ex As Exception
                        Console.WriteLine(ex.Message)
                    End Try
                Loop Until line Is Nothing
                sr.Close()
            Catch ex As Exception
                Console.WriteLine("Exception in loading customers from file:" & ex.Message)
            End Try
        End Sub
        Public Overrides Sub loadFromSOAP(ByVal filename As String)
            Dim fs As New FileStream(filename, FileMode.Open)
            Dim formatter As New SoapFormatter
            Dim custList As DLL_Library.IOTS.CustomerHolder = Nothing
            custList = DirectCast(formatter.Deserialize(fs), CustomerHolder)
            Me.customerList = custList.customerList
        End Sub

        Public Overrides Sub SyncToFile(ByVal fileName As String)
            Dim SR As StreamReader = New StreamReader(fileName)
            Dim header As String = SR.ReadLine()
            SR.Close()
            Dim sw As StreamWriter = New StreamWriter(fileName)
            sw.WriteLine(header)
            For i As Integer = 0 To customerList.Count - 1
                sw.WriteLine(customerList(i))
            Next
            sw.Close()
        End Sub

        Public Overrides Sub SyncToSOAP(fileName As String)
            Dim sformatter As New SoapFormatter
            Dim soapStream As FileStream = New FileStream(fileName, FileMode.Create)
            sformatter.Serialize(soapStream, Me)
            soapStream.Close()
        End Sub

        Public Function SearchByName(ByVal pid As String) As List(Of Customer)
            Dim result As New List(Of Customer)()
            For Each cust As Customer In customerList
                If cust._firstName.ToLower.Contains(pid.ToLower) Or
                    cust._lastName.ToLower.Contains(pid.ToLower) Then
                    result.Add(cust)
                End If
            Next
            Return result
        End Function

        Public Sub AddCustomer(ByVal newCust As Customer)
            For Each cust As Customer In customerList
                If newCust._custId = cust._custId Then
                    Throw New OrderSystemExceptions("The customer id can not be duplicated.")
                End If
            Next
            customerList.Add(newCust)
        End Sub

        Public Sub deleteCustomerById(ByVal custId As Long)
            Dim found As Boolean = False
            For i As Integer = customerList.Count - 1 To 0 Step -1
                Dim cust As Customer = customerList(i)
                If cust._custId = custId Then
                    customerList.RemoveAt(i)
                    found = True
                    Exit For
                End If
            Next
            If Not found Then
                Throw New OrderSystemExceptions("Can't find customer with Id of" & custId & ". No Customer removed.")
            End If
        End Sub
    End Class
End Namespace

