Imports CSLib
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Runtime.Serialization.Formatters.Soap
Namespace IOTS
    <Serializable()>
    Public Class Product
        Implements IValidator
        Private ProductId As String
        Private Description As String
        Private QtyOnHand As Integer
        Private Price As Double

        Public Property _Price() As String
            Get
                ' Return the value stored in the local variable. 
                Return Price
            End Get
            Set(ByVal Value As String)
                ' Store the value in a local variable.
                'price can has 2 decimals max.
                Dim regName As String = "^\d+([.]\d{1,2})?$"
                If IsValid(Value.Trim, regName) Then
                    Price = Value
                Else
                    Throw New OrderSystemExceptions("An exception has occurred, price should have two decimal places.")
                End If
            End Set
        End Property
        Public Property _description() As String
            Get
                Return Description
            End Get
            Set(value As String)
                Description = value
            End Set
        End Property

        Public Property _QtyOnHand() As Integer
            Get
                ' Return the value stored in the local variable. 
                Return QtyOnHand
            End Get
            Set(ByVal Value As Integer)
                If Value >= 0 Then
                    QtyOnHand = Value
                Else
                    Throw New OrderSystemExceptions("An exception has occurred, quantity on hand cannot less than 0.")
                End If
            End Set
        End Property
        Public Property _productId() As String
            Get
                ' Return the value stored in the local variable. 
                Return ProductId
            End Get
            Set(ByVal Value As String)
                ' Store the value in a local variable.
                Dim regName As String = "^[A-Z][A-Z][0-9][0-9]$"
                If IsValid(Value, regName) Then
                    ProductId = Value
                Else
                    Throw New OrderSystemExceptions("An exception has occurred, productId should in XX00 format.")
                End If
            End Set
        End Property
        Sub New(ByVal ProductId, ByVal Description, ByVal QtyOnHand, ByVal Price)
            _productId = ProductId
            MyClass.Description = Description
            _QtyOnHand = QtyOnHand
            _Price = Price
        End Sub

        Public Function IsValid(ByVal testData As String, ByVal regex As String) As Boolean Implements IValidator.IsValid
            Dim r As Regex = New Regex(regex, RegexOptions.IgnoreCase)
            Dim m As Match = r.Match(testData)
            Return m.Success
        End Function

        Public Overrides Function ToString() As String
            Return String.Format("{0}, {1}, {2}, {3}", ProductId, Description, QtyOnHand, Price)
        End Function

        Public Function GetDetails() As String
            Dim result As String = Nothing
            result = result + "Product Id:" + _productId + vbCrLf
            result = result + "Product Description: " + _description + vbCrLf
            result = result + "QtyOnHand: " + CStr(_QtyOnHand) + vbTab
            result = result + "Price: $" + CStr(_Price)

            Return result
        End Function
    End Class

    <Serializable()>
    Public Class ProductHolder
        Inherits ObjectHolder
        Public productList As New ArrayList
        Sub New(products As List(Of Product))
            productList = New ArrayList()
            For i As Integer = 0 To products.Count - 1
                productList.Add(products(i))
            Next
        End Sub
        Sub New()

        End Sub
        Public Overrides Sub loadFromFile(ByVal filename As String)

            Dim line As String = Nothing
            Dim lineData() As String
            Dim product As Product
            'empty the old list
            productList = New ArrayList()
            Try
                'read every line and split
                Dim sr As StreamReader = New StreamReader(filename)
                'the first line is header
                line = sr.ReadLine()
                Do
                    Try
                        line = sr.ReadLine()
                        lineData = Split(line, ",")
                        If lineData.Length >= 3 Then
                            Console.WriteLine("Reading record:" + line)
                            'save to array
                            product = New Product(lineData(0), lineData(1), lineData(2), lineData(3))
                            productList.Add(product)
                        End If
                    Catch ex As Exception
                        Console.WriteLine(ex.Message)
                    End Try
                Loop Until line Is Nothing
                sr.Close()
            Catch ex As Exception
                Console.WriteLine("Exception in loading products:" & ex.Message)
            End Try
        End Sub
        Public Overrides Sub loadFromSOAP(ByVal filename As String)
            Dim fs As New FileStream(filename, FileMode.Open)
            Dim formatter As New SoapFormatter
            Dim newList As DLL_Library.IOTS.ProductHolder = Nothing
            newList = DirectCast(formatter.Deserialize(fs), ProductHolder)
            Me.productList = newList.productList
        End Sub

        Public Overrides Sub SyncToFile(ByVal fileName As String)
            Dim SR As StreamReader = New StreamReader(fileName)
            Dim header As String = SR.ReadLine()
            SR.Close()
            Dim sw As StreamWriter = New StreamWriter(fileName)
            sw.WriteLine(header)
            For i As Integer = 0 To productList.Count - 1
                sw.WriteLine(productList(i))
            Next
            sw.Close()
        End Sub

        Public Overrides Sub SyncToSOAP(fileName As String)
            Dim sformatter As New SoapFormatter
            Dim soapStream As FileStream = New FileStream(fileName, FileMode.Create)
            sformatter.Serialize(soapStream, Me)
            soapStream.Close()
        End Sub

        Public Function SearchById(ByVal pid As String) As Product
            Dim result As Product = Nothing
            For Each prod As Product In productList
                If prod._productId.ToLower.Equals(pid.Trim.ToLower) Then
                    result = prod
                    Exit For
                End If
            Next
            Return result
        End Function

        Public Sub addProduct(ByVal newPord As Product)
            For Each prod As Product In productList
                If prod._productId.ToLower.Equals(newPord._productId.ToLower) Then
                    Throw New OrderSystemExceptions("Product ID is duplicated, please change it to another id.")
                End If
            Next
            productList.Add(newPord)
        End Sub

        Public Sub deleteProductById(ByVal productId As String)
            Dim found As Boolean = False
            For i As Integer = productList.Count - 1 To 0 Step -1
                Dim prod As Product = productList(i)
                If prod._productId.ToLower.Equals(productId.ToLower) Then
                    productList.RemoveAt(i)
                    found = True
                    Exit For
                End If
            Next
            If Not found Then
                Throw New OrderSystemExceptions("Can't find product with Id of" & productId & ". No product removed")
            End If
        End Sub
    End Class
End Namespace