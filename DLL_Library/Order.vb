Imports CSLib
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Runtime.Serialization.Formatters.Soap
Namespace IOTS
    <Serializable()>
    Public Class Order
        Implements IValidator
        Private orderNumber As Long
        Private orderDate As String
        Private ShipDate As String
        Private numberOrdered As Integer
        Private productId As String
        Private discount As Double
        Private custId As Long

        'the order number cannot be changed!
        Public ReadOnly Property _orderNumber() As Long
            Get
                Return orderNumber
            End Get
        End Property


        Public Property _productId() As String
            Get
                ' Return the value stored in the local variable. 
                Return productId
            End Get
            Set(ByVal Value As String)
                ' Store the value in a local variable.
                Dim regName As String = "^[A-Z][A-Z][0-9][0-9]$"
                If IsValid(Value, regName) Then
                    productId = Value
                Else
                    Throw New OrderSystemExceptions("An exception has occurred, productId should in XX00 format.")
                End If
            End Set
        End Property
        Public Property _numberOrdered() As Integer
            Get
                ' Return the value stored in the local variable. 
                Return numberOrdered
            End Get
            Set(ByVal Value As Integer)
                If Value > 0 Then
                    numberOrdered = Value
                Else
                    Throw New OrderSystemExceptions("An exception has occurred, number ordered should more than 0.")
                End If
            End Set
        End Property
        Public Property _orderDate() As String
            Get
                Return orderDate
            End Get
            Set(ByVal Value As String)
                If Value = Nothing Then
                    Throw New OrderSystemExceptions("An exception has occurred, orderDate not valid.")
                Else
                    Dim dateIn As Date = Date.ParseExact(Value, "dd/MM/yyyy",
            System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    'Order Date cannot be future
                    Dim result As Integer = DateTime.Compare(dateIn, Date.Now())
                    If result <= 0 Then
                        orderDate = dateIn
                    Else
                        Throw New OrderSystemExceptions("orderDate has to be earlier or same to current date.")
                    End If
                End If
            End Set
        End Property

        Public Property _shipDate() As String
            Get
                Return ShipDate
            End Get
            Set(ByVal Value As String)
                If Value = Nothing Then
                    Throw New OrderSystemExceptions("An exception has occurred, shipDate not valid.")
                Else
                    Dim dateIn As Date = Date.ParseExact(Value, "dd/MM/yyyy",
            System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    Dim result1 As Integer = DateTime.Compare(orderDate, dateIn)
                    Dim result2 As Integer = DateTime.Compare(dateIn, Date.Now())
                    'If result1 <= 0 And result2 >= 0 Then
                    If result1 <= 0 And result1 >= -180 Then
                        ShipDate = dateIn
                    Else
                        Throw New OrderSystemExceptions("shipDate has to be later than or same as orderDate as well as current date.")
                    End If
                End If
            End Set
        End Property

        Public Property _custId() As Long
            Get
                Return custId
            End Get
            Set(ByVal Value As Long)
                If Value > 0 Then
                    custId = Value
                Else
                    Throw New OrderSystemExceptions("custId has to be positive integer.")
                End If

            End Set
        End Property

        Public Property _discount() As Double
            Get
                Return discount
            End Get
            Set(value As Double)
                Dim regName As String = "^\d+([.]\d{1,2})?$"
                If IsValid(value.ToString, regName) And (value >= 0) Then
                    discount = value
                Else
                    Throw New OrderSystemExceptions("Discount must be numbers with maximum of 2 decimals")
                End If
            End Set
        End Property
        Sub New(ByVal orderNumber As Long, ByVal orderDate As String,
                ByVal ShipDate As String, ByVal numberOrdered As Integer,
                ByVal productId As String, ByVal discount As Double,
                ByVal custid As Long)
            MyClass.orderNumber = orderNumber
            _orderDate() = orderDate
            _shipDate() = ShipDate
            _numberOrdered = numberOrdered
            _productId = productId

            _discount = discount

                MyClass.custId = custid

        End Sub

        Public Function IsValid(ByVal testData As String, ByVal regex As String) As Boolean Implements IValidator.IsValid
            Dim r As Regex = New Regex(regex, RegexOptions.IgnoreCase)
            Dim m As Match = r.Match(testData)
            Return m.Success
        End Function

        Public Overrides Function ToString() As String
            Return String.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}", orderNumber, orderDate, ShipDate, numberOrdered, productId, discount, custId)
        End Function


    End Class

    <Serializable()>
    Public Class OrderHolder
        Inherits ObjectHolder
        Public orderList As New ArrayList
        Public orderDetails As List(Of OrderDetails)
        Sub New(orders As List(Of Order))
            orderList = New ArrayList()
            For i As Integer = 0 To orders.Count - 1
                orderList.Add(orders(i))
            Next
        End Sub
        Sub New()

        End Sub
        Public Overrides Sub loadFromFile(ByVal filename As String)
            Dim sr As StreamReader = New StreamReader(filename)
            Dim line As String = Nothing
            Dim lineData() As String
            Dim order As Order
            Try
                'empty the old list
                If Not (orderList Is Nothing) Then

                    orderList.Clear()
                End If
                'read every line and split

                'the first line is header
                line = sr.ReadLine()
                Do
                    Try
                        line = sr.ReadLine()
                        lineData = Split(line, ",")
                        If lineData.Length >= 7 Then
                            Console.WriteLine("Reading record:" + line)
                            'save to array
                            order = New DLL_Library.IOTS.Order(lineData(0), lineData(1), lineData(2), lineData(3), lineData(4), lineData(5), lineData(6))
                            orderList.Add(order)
                        End If
                    Catch ex As Exception
                        Console.WriteLine(ex.Message)
                    End Try
                Loop Until line Is Nothing
                sr.Close()
            Catch ex As Exception
                Console.WriteLine("Exception in loading orders from file:" & ex.Message)
            End Try
        End Sub
        Public Overrides Sub loadFromSOAP(ByVal filename As String)
            Dim fs As New FileStream(filename, FileMode.Open)
            Dim formatter As New SoapFormatter
            Dim newList As DLL_Library.IOTS.OrderHolder = Nothing
            newList = DirectCast(formatter.Deserialize(fs), OrderHolder)
            Me.orderList = newList.orderList
        End Sub

        Public Overrides Sub SyncToFile(ByVal fileName As String)
            Dim SR As StreamReader = New StreamReader(fileName)
            Dim header As String = SR.ReadLine()
            SR.Close()
            Dim sw As StreamWriter = New StreamWriter(fileName)
            sw.WriteLine(header)
            For i As Integer = 0 To orderList.Count - 1
                sw.WriteLine(orderList(i))
            Next
            sw.Close()
        End Sub

        Public Overrides Sub SyncToSOAP(fileName As String)
            Dim sformatter As New SoapFormatter
            Dim soapStream As FileStream = New FileStream(fileName, FileMode.Create)
            sformatter.Serialize(soapStream, Me)
            soapStream.Close()
        End Sub

        'check if the order is valid and return a list of valid orders
        Public Sub Validate(customers As CustomerHolder, products As ProductHolder) ' As List(Of OrderDetails)
            Dim validOrderDetails As New List(Of OrderDetails)()
            Dim hasCust As Boolean = False
            Dim hasProd As Boolean = False
            Dim validOrder As OrderDetails = Nothing
            Dim rightCust As Customer = Nothing
            Dim rightProd As Product = Nothing
            For i As Integer = orderList.Count - 1 To 0 Step -1
                Dim order As Order = orderList(i)
                For Each product As Product In products.productList
                    If product._productId.Equals(order._productId) Then
                        rightProd = product
                        hasProd = True
                        Exit For
                    End If
                Next

                For Each cust As Customer In customers.customerList
                    If cust._custId = order._custId Then
                        rightCust = cust
                        hasCust = True
                        Exit For
                    End If
                Next
                If (hasCust And hasProd) Then
                    validOrder = New OrderDetails(order, rightCust, rightProd)
                    validOrderDetails.Add(validOrder)
                Else
                    If Not hasProd Then
                        Console.WriteLine("No such product in product list!")

                    End If
                    If Not hasCust Then
                        Console.WriteLine("No such customer in customer list!")

                    End If
                    'invalid order will be removed
                    orderList.Remove(order)
                End If
                hasCust = False
                hasProd = False
            Next
            orderDetails = validOrderDetails
            'Return validOrderDetails
        End Sub
        Public Function SearchById(ByVal oid As Long) As Order
            Dim result As Order = Nothing
            For Each ord As Order In orderList
                If ord._orderNumber = oid Then
                    result = ord
                    Exit For
                End If
            Next
            Return result
        End Function

        Public Function GetOrderDetails(ByVal ord As Order) As OrderDetails
            Dim result As OrderDetails = Nothing
            For Each orderDetail As OrderDetails In orderDetails
                If orderDetail._order._orderNumber = ord._orderNumber Then
                    result = orderDetail
                End If
            Next
            Return result
        End Function
        Public Function validateOrder(ByVal order As Order,
                                      customers As CustomerHolder,
                                      products As ProductHolder) As String
            Dim result As Boolean = False
            Dim hasCust As Boolean = False
            Dim hasProd As Boolean = False
            Dim rightProd As Product = Nothing
            Dim rightCust As Customer = Nothing
            For Each product As Product In products.productList
                If product._productId.Equals(order._productId) Then
                    Dim diff As Double = Convert.ToDouble(product._Price) - order._discount
                    'check if the discount is valid 
                    If (diff >= 0) Then
                        rightProd = product
                        hasProd = True
                    Else Throw New OrderSystemExceptions("Discount can not be more than the product price!!")
                    End If
                    Exit For
                End If
            Next

            For Each cust As Customer In customers.customerList
                If cust._custId = order._custId Then
                    rightCust = cust
                    hasCust = True
                    Exit For
                End If
            Next
            If (hasCust And hasProd) Then
                If rightCust._creditLimit >= order._numberOrdered * (Convert.ToDouble(rightProd._Price) - order._discount) Then
                    orderList.Add(order)
                    orderDetails.Add(New OrderDetails(order, rightCust, rightProd))
                    result = True
                Else
                    Throw New OrderSystemExceptions("The customer hasn't enough credit for the order!!")
                End If
                'Check if customer has enough credit

            Else
                If Not hasProd Then
                    Throw New OrderSystemExceptions("No such product in product list!")

                End If
                If Not hasCust Then
                    Throw New OrderSystemExceptions("No such customer in customer list!")

                End If
                'invalid order will be removed

            End If
            Return result

        End Function
        Public Sub deleteOrderById(ByVal orderId As Long)
            Dim found1 As Boolean = False
            Dim found2 As Boolean = False
            For i As Integer = orderList.Count - 1 To 0 Step -1
                Dim order As Order = orderList(i)
                If order._orderNumber = orderId Then
                    orderList.RemoveAt(i)
                    found1 = True
                    Exit For
                End If
            Next
            For i As Integer = orderDetails.Count - 1 To 0 Step -1
                Dim orderDetail As OrderDetails = orderDetails(i)
                If orderDetail._order._orderNumber = orderId Then
                    orderDetails.RemoveAt(i)
                    found2 = True
                    Exit For
                End If
            Next
            If Not (found1) Then
                Throw New OrderSystemExceptions("Can't find order with number of" & orderId & ". No product removed")
            End If
        End Sub
        'Check if the new order is valid then add it if yes
        Public Function addOrder(ByVal order As Order, customers As CustomerHolder, products As ProductHolder) As Boolean
            Dim result As Boolean = False
            Dim hasCust As Boolean = False
            Dim hasProd As Boolean = False
            Dim rightProd As Product = Nothing
            Dim rightCust As Customer = Nothing
            For Each product As Product In products.productList
                If product._productId.Equals(order._productId) Then
                    Dim diff As Double = Convert.ToDouble(product._Price) - order._discount
                    'check if the discount is valid 
                    If (diff >= 0) Then
                        rightProd = product
                        hasProd = True
                    Else Throw New OrderSystemExceptions("Discount is more than the product price!!")
                    End If
                    Exit For
                End If
            Next

            For Each cust As Customer In customers.customerList
                If cust._custId = order._custId Then
                    rightCust = cust
                    hasCust = True
                    Exit For
                End If
            Next
            If (hasCust And hasProd) Then
                If rightCust._creditLimit >= order._numberOrdered * (Convert.ToDouble(rightProd._Price) - order._discount) Then
                    orderList.Add(order)
                    orderDetails.Add(New OrderDetails(order, rightCust, rightProd))
                    result = True
                Else
                    Throw New OrderSystemExceptions("The customer hasn't enough credit for the order!!")
                End If
                'Check if customer has enough credit

            Else
                If Not hasProd Then
                    Throw New OrderSystemExceptions("No such product in product list!")

                End If
                If Not hasCust Then
                    Throw New OrderSystemExceptions("No such customer in customer list!")

                End If
                'invalid order will be removed
                orderList.Remove(order)
            End If
            Return result
        End Function
    End Class

End Namespace