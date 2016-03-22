'OrderDetails class is used to display order details 
Namespace IOTS
    Public Class OrderDetails
        Private order As Order
        Private customer As Customer
        Private Product As Product

        Public Sub New(order As Order, cust As Customer, prod As Product)
            Me.order = order
            Me.customer = cust
            Me.Product = prod
        End Sub
        Public ReadOnly Property _order As Order
            Get
                Return order
            End Get
        End Property
        Public Overrides Function ToString() As String
            Dim one As String

            one = one + "Order Number: " + CStr(order._orderNumber) & vbCrLf
            one = one + "Order Date: " + order._orderDate
            one = one + "  Ship Date: " + order._shipDate & vbCrLf
            one = one + Product.GetDetails()
            one = one & vbCrLf
            one = one + "Quantity: " + CStr(order._numberOrdered) & vbCrLf
            one = one + customer.GetDetails()
            one = one & vbCrLf
            If (order._discount > 0) Then
                one = one + "Discount Amount: $" + order._discount.ToString()
            End If

            Return one
        End Function
    End Class
End Namespace