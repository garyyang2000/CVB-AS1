Imports System.IO
Imports System.DateTime

Module Module1

    Sub Main()
        Dim orderBook As New DLL_Library.IOTS.OrderHolder()
        Dim productBook As New DLL_Library.IOTS.ProductHolder()
        Dim custBook As New DLL_Library.IOTS.CustomerHolder()
        'Dim orderDetails As New List(Of DLL_Library.IOTS.OrderDetails)()
        Try
            Console.WriteLine("Reading data from file......")
            Dim fileName As String = "c:\temp\cvb\order.csv"
            orderBook.loadFromFile(fileName)
            fileName = "c:\temp\cvb\product.csv"
            productBook.loadFromFile(fileName)
            fileName = "c:\temp\cvb\customer.csv"
            custBook.loadFromFile(fileName)
            orderBook.Validate(custBook, productBook)

            Dim choice As Integer
            'Console.SetBufferSize(Console.WindowWidth, Console.WindowHeight)

            Do
                DisplayHeader()
                GetMenuSelection(choice)
                Select Case choice
                    Case 1
                        ClientSearch(custBook)
                    Case 2
                        OrderSearch(orderBook, productBook, custBook)
                    Case 3
                        ProductSearch(productBook)
                    Case 4
                        AddClient(custBook)
                    Case 5
                        AddOrder(orderBook, productBook, custBook)
                    Case 6
                        AddProduct(productBook)
                    Case 7
                        'Quit
                End Select
            Loop Until (choice = 7)
        Catch ex As Exception
            Console.WriteLine("Operation Failed:{0}", ex.Message)
        End Try
        Console.ReadLine()
    End Sub



    Private Sub AddOrder(ByVal orderBook As DLL_Library.IOTS.OrderHolder,
                         ByVal productBook As DLL_Library.IOTS.ProductHolder,
                         ByVal custBook As DLL_Library.IOTS.CustomerHolder)
        ' Adds new music items to the database.

        ' Local variables.

        Dim productId As String = ""
        Dim numberOrdered As String = ""
        Dim discount As Integer
        Dim custId As Long
        Dim shipDate As String
        Dim orderDate As String
        vpLocate(25, 1)                                     ' Print message on status line.
        vpPrint("Enter Order data.  Type END for product ID to quit...")
        vpViewPrint(5, 23)                                  ' Enable viewport (lines 5 - 23).
        vpPrintLine()                                       ' Prompt for data.
        vpPrintLine("Enter new order information (without commas)")
        vpPrintLine()

        ' Get records for file until user enters END for title.

        While (Strings.UCase(productId) <> "END")
            Dim oid As Long = orderBook.orderList.Count
            oid = CInt(Int((10000000 * Rnd()) + oid))
            vpPrintLine("Order Number:  " & oid.ToString)
            vpPrint("ProductId:  ")
            productId = Console.ReadLine()                      ' Get item title.

            If (Strings.UCase(productId) <> "END") Then
                Try
                    vpPrint("Quantity for order:  ")
                    numberOrdered = Convert.ToInt64(Console.ReadLine())                ' ... and other info.

                    vpPrint("Customer ID:  ")
                    custId = Integer.Parse(Console.ReadLine())                 ' ... and other info.

                    vpPrint("Shipping Date(dd/mm/yyyy):  ")
                    shipDate = Console.ReadLine()

                    vpPrint("Discount:  ")
                    discount = Double.Parse(Console.ReadLine())

                    vpPrintLine()
                    ' Write record to database file.
                    Dim today As Date = Date.Now()
                    orderDate = today.ToString("dd\/MM\/yyyy")

                    Dim newOrder As New DLL_Library.IOTS.Order(oid, orderDate, shipDate, numberOrdered, productId, discount, custId)
                    If orderBook.addOrder(newOrder, custBook, productBook) Then
                        vpPrintLine("New order added!")
                        'vpPrintLine(orderBook.GetOrderDetails(newOrder).ToString)
                    Else
                        Throw New DLL_Library.OrderSystemExceptions("Not a valid order")
                    End If
                Catch ex As Exception
                    vpViewPrint(5, 23)                                  ' Enable viewport (lines 5 - 23).  
                    vpCls(2)
                    vpPrintLine("Failed to add this order." & vbCrLf & "Possible reason: " & ex.Message)
                    vpPrintLine()
                    vpPrintLine("Enter new product information (without commas)")
                    vpPrintLine()
                End Try
            End If
        End While
    End Sub

    Private Sub AddProduct(ByVal productBook As DLL_Library.IOTS.ProductHolder)
        ' Adds new music items to the database.

        ' Local variables.

        Dim productId As String = ""
        Dim description As String = ""
        Dim qtyOnHand As Integer
        Dim price As Double

        vpLocate(25, 1)                                     ' Print message on status line.
        vpPrint("Enter Product data.  Type END for product ID to quit...")
        vpViewPrint(5, 23)                                  ' Enable viewport (lines 5 - 23).
        vpPrintLine()                                       ' Prompt for data.
        vpPrintLine("Enter new product information (without commas)")
        vpPrintLine()

        ' Get records for file until user enters END for title.

        While (Strings.UCase(productId) <> "END")
            vpPrint("ProductId:  ")
            productId = Console.ReadLine()                      ' Get item title.

            If (Strings.UCase(productId) <> "END") Then
                Try
                    vpPrint("Description:  ")
                    description = Console.ReadLine()                 ' ... and other info.

                    vpPrint("Quantity on hand:  ")
                    qtyOnHand = Integer.Parse(Console.ReadLine())
                    vpPrint("Price:  ")
                    price = Double.Parse(Console.ReadLine())
                    Dim newProd As New DLL_Library.IOTS.Product(productId, description, qtyOnHand, price)
                    productBook.addProduct(newProd)
                    vpPrintLine()
                    ' Write record to database file.
                Catch ex As Exception
                    vpViewPrint(5, 23)                                  ' Enable viewport (lines 5 - 23).  
                    vpCls(2)
                    vpPrintLine("Failed to add this product." & vbCrLf & "Possible reason:" + ex.Message)
                    vpPrintLine()
                    vpPrintLine("Enter new product information (without commas)")
                    vpPrintLine()
                End Try
            End If
        End While
        vpViewPrint()                                               ' Disable viewport and update status line.
        vpLocate(25, 1)
        vpPrint("Press Enter to return to main menu...")
        Console.ReadKey()
    End Sub

    Private Sub AddClient(ByVal custBook As DLL_Library.IOTS.CustomerHolder)
        ' Adds new music items to the database.

        ' Local variables.

        Dim first_name As String = ""
        Dim last_name As String = ""
        Dim address As String = ""
        Dim city As String = ""
        Dim province As String = ""
        Dim postalcode As String = ""
        Dim credit_limit As Double = 0
        Dim email As String = ""


        vpLocate(25, 1)                                     ' Print message on status line.
        vpPrint("Enter customer data.  Type END for first name to quit...")
        vpViewPrint(5, 23)                                  ' Enable viewport (lines 5 - 23).
        vpPrintLine()                                       ' Prompt for data.
        vpPrintLine("Enter new customer information (without commas)")
        vpPrintLine()

        ' Get records for file until user enters END for title.

        While (Strings.UCase(first_name) <> "END")
            vpPrint("First name:  ")
            first_name = Console.ReadLine()                      ' Get item title.

            If (Strings.UCase(first_name) <> "END") Then
                Try
                    vpPrint("Last name:  ")
                    last_name = Console.ReadLine()                 ' ... and other info.
                    vpPrint("Address:  ")
                    address = Console.ReadLine()
                    vpPrint("City:  ")
                    city = Console.ReadLine()
                    vpPrint("Province:  ")
                    province = Console.ReadLine()
                    vpPrint("Postal Code:  ")
                    postalcode = Console.ReadLine()
                    vpPrint("Credit Limit(with max of 2 decimals):  ")
                    credit_limit = Double.Parse(Console.ReadLine())
                    vpPrint("Email:  ")
                    email = Console.ReadLine()

                    vpPrintLine()
                    ' Write record to database

                    Dim custId As Long = custBook.customerList.Count
                    'auto generate a integer as cust id
                    custId = CInt(Int((10 * Rnd()) + custId))
                    Dim newCust As New DLL_Library.IOTS.Customer(custId,
                             first_name, last_name, address, city, province, postalcode, credit_limit, email)
                    custBook.AddCustomer(newCust)
                Catch ex As Exception

                    vpViewPrint(5, 23)                                  ' Enable viewport (lines 5 - 23).  
                    vpCls(2)
                    vpPrintLine("Failed to create customer." & vbCrLf & "Possible reason:" + ex.Message)
                    vpPrintLine()
                    vpPrintLine("Enter new customer information (without commas)")
                    vpPrintLine()
                End Try
            End If
        End While

        vpViewPrint()                                               ' Disable viewport and update status line.
        vpLocate(25, 1)
        vpPrint("Press Enter to return to main menu...")
        Console.ReadKey()
    End Sub

    Private Sub ProductSearch(ByVal productList As DLL_Library.IOTS.ProductHolder)
        Dim searchStr As String
        Dim found As Boolean = False                                ' Initialize "record found" flag.
        Dim num As Integer = 0
        vpViewPrint(5, 23)
        vpPrintLine()                                               ' Get search string.
        vpPrintLine("Product Search")
        vpPrint("Enter product id to be searched for:  ")
        searchStr = Console.ReadLine()
        Dim foundProd As DLL_Library.IOTS.Product = productList.SearchById(searchStr)
        If foundProd IsNot Nothing Then
            vpPrintLine("Search results:")
            ' Display search results.
            vpPrintLine(foundProd.GetDetails)
            found = True
        Else
            vpPrintLine()
        End If
        'Search and display client

        If (Not found) Then
            vpColor(ConsoleColor.DarkGreen)                         ' "not found" message.
            vpPrint(searchStr)
            vpColor(ConsoleColor.White)
            vpPrintLine(" not found in database")
        Else
            vpPrintLine()
            vpPrintLine("Select a operation:")                    ' Prompt for search topic.
            vpPrintLine()
            vpPrintLine("  1) EDIT Product")
            vpPrintLine("  2) DELETE Product")
            vpPrintLine("  3) QUIT to main menu")
            vpPrintLine()

            Do While (num < 1) Or (num > 3)                             ' Get number associated with
                vpPrint("Operation (1-3):  ")                            ' search topic.
                num = CInt(Console.ReadLine())
            Loop

            Select Case num
                Case 1

                    EditProduct(productList, foundProd)
                Case 2
                    vpPrint("Are you sure to delete this product?(Y/N):  ")
                    Dim choice As String = Console.ReadLine()
                    ' call delete product
                    If choice.ToUpper.Equals("Y") Then
                        productList.productList.Remove(foundProd)
                    End If

            End Select
        End If

        vpViewPrint()                                               ' Disable viewport and update status line.
        vpLocate(25, 1)
        vpPrint("Press Enter to return to main menu...")
        Console.ReadKey()
    End Sub


    Private Sub OrderSearch(ByVal orderList As DLL_Library.IOTS.OrderHolder,
                            ByVal productBook As DLL_Library.IOTS.ProductHolder,
                         ByVal custBook As DLL_Library.IOTS.CustomerHolder)
        Dim searchStr As String
        Dim found As Boolean = False                                ' Initialize "record found" flag.

        vpViewPrint(3, 23)
        vpPrintLine()                                               ' Get search string.
        vpPrintLine("Order Search")
        vpPrint("Enter order number(must be integer) to be searched for:  ")
        searchStr = Console.ReadLine()
        Try
            Dim orderNumber As Long = Convert.ToInt64(searchStr)
            Dim foundOrder As DLL_Library.IOTS.Order = orderList.SearchById(orderNumber)
            If foundOrder IsNot Nothing Then
                Dim foundOrderDetail As DLL_Library.IOTS.OrderDetails = orderList.GetOrderDetails(foundOrder)

                vpPrintLine("Search results:")
                ' Display search results.
                vpPrintLine(foundOrderDetail.ToString())
                found = True
            Else
                vpPrintLine()
            End If
            'Search and display client

            If (Not found) Then
                vpColor(ConsoleColor.DarkGreen)                         ' "not found" message.
                vpPrint(searchStr)
                vpColor(ConsoleColor.White)
                vpPrintLine(" not found in database")
            Else
                vpPrintLine()
                vpPrintLine("Select a operation:")                    ' Prompt for search topic.
                vpPrintLine()
                vpPrintLine("  1) EDIT Order")
                vpPrintLine("  2) DELETE Order")
                vpPrintLine("  3) QUIT to main menu")

                vpPrintLine()
                Dim num As Integer = 0
                Do While (num < 1) Or (num > 3)                             ' Get number associated with
                    vpPrint("Operation (1-3):  ")                            ' search topic.
                    num = CInt(Console.ReadLine())
                Loop

                Dim orderId As Integer
                Select Case num
                    Case 1
                        'vpPrint("Enter order id to edit:  ")
                        'orderId = Integer.Parse(Console.ReadLine())
                        EditOrder(orderList, foundOrder, productBook, custBook)
                    Case 2
                        vpPrint("Enter order id to delete:  ")
                        orderId = Integer.Parse(Console.ReadLine())
                        vpPrint("Are you sure to delete this order?(Y/N):  ")
                        Dim choice As String = Console.ReadLine()
                        ' call delete customer
                        If choice.ToUpper.Equals("Y") Then
                            orderList.deleteOrderById(orderId)
                            vpPrint("order deleted.")
                        End If
                        ' call delete order


                End Select
            End If
        Catch ex As Exception
            vpPrintLine("Operation failed.")
            vpPrintLine("Possible reason: " & ex.Message)
        End Try
        vpViewPrint()                                               ' Disable viewport and update status line.
        vpLocate(25, 1)
        vpPrint("Press Enter to return to main menu...")
        Console.ReadKey()
    End Sub

    Private Sub ClientSearch(ByVal custHolder As DLL_Library.IOTS.CustomerHolder)
        Dim searchStr As String
        Dim found As Boolean = False                                ' Initialize "record found" flag.
        Dim num As Integer = 0
        vpViewPrint(5, 23)
        vpPrintLine()                                               ' Get search string.
        vpPrintLine("Customer Search")
        vpPrint("Enter customer name to be searched for:  ")
        searchStr = Console.ReadLine()
        Dim results As List(Of DLL_Library.IOTS.Customer) = custHolder.SearchByName(searchStr)
        vpPrintLine("Search results:")                              ' Display search results.
        If (results.Count > 0) Then
            searchStr = ""
            For Each cust As DLL_Library.IOTS.Customer In results
                searchStr = searchStr + cust.GetDetails + vbCrLf
            Next
            vpPrintLine(searchStr)
            found = True
        Else vpPrintLine()
        End If
        'Search and display Customer

        If (Not found) Then
            vpColor(ConsoleColor.DarkGreen)                         ' "not found" message.
            vpPrint(searchStr)
            vpColor(ConsoleColor.White)
            vpPrintLine(" not found in database")
        Else
            vpPrintLine()
            vpPrintLine("Select a operation:")                    ' Prompt for search topic.
            vpPrintLine()
            vpPrintLine("  1) EDIT Customer")
            vpPrintLine("  2) DELETE Customer")
            vpPrintLine("  3) QUIT to main menu")

            vpPrintLine()
            Try
                Do While (num < 1) Or (num > 3)                             ' Get number associated with
                    vpPrint("Operation (1-3):  ")                            ' search topic.
                    num = CInt(Console.ReadLine())
                Loop

                Dim custId As Integer
                Select Case num
                    Case 1
                        found = False
                        vpPrint("Enter customer id to edit:  ")
                        custId = Convert.ToInt64(Console.ReadLine())
                        For i As Integer = 0 To results.Count - 1
                            If custId = results(i)._custId Then
                                EditClient(custHolder, results(i))
                                found = True
                                Exit For
                            End If

                        Next
                        If Not found Then
                            Throw New DLL_Library.OrderSystemExceptions("Customer id is not in the result list")
                        End If
                    Case 2
                        vpPrint("Enter customer id to delete:  ")
                        custId = Convert.ToInt64(Console.ReadLine())
                        vpPrint("Are you sure to delete this customer?(Y/N):  ")
                        Dim choice As String = Console.ReadLine()
                        ' call delete customer
                        If choice.ToUpper.Equals("Y") Then
                            custHolder.deleteCustomerById(custId)
                            vpPrint("Customer deleted.")
                        End If
                        ' call delete customer
                End Select
            Catch ex As Exception
                vpPrintLine("Operation failed. Reason:" & ex.Message)
            End Try
        End If

        vpViewPrint()                                               ' Disable viewport and update status line.
        vpLocate(25, 1)
        vpPrint("Press Enter to return to main menu...")
        Console.ReadKey()
    End Sub

    Public Sub EditClient(ByRef custBook As DLL_Library.IOTS.CustomerHolder,
                          ByVal cust As DLL_Library.IOTS.Customer)
        Try
            vpViewPrint(5, 23)                                                                   ' Enable viewport (lines 5-23).
            vpCls(2)                                            ' Clear viewport for choice prompts.
            'display customer information
            vpPrintLine("Customer:")
            vpPrintLine(cust.GetDetails)                                       ' Prompt for data.

            Dim first_name As String = ""
            Dim last_name As String = ""
            Dim address As String = ""
            Dim city As String = ""
            Dim province As String = ""
            Dim postalcode As String = ""
            Dim credit_limit As Double = 0
            Dim email As String = ""

            vpPrint("Enter customer data.")

            vpPrintLine()                                       ' Prompt for data.
            vpPrintLine("Enter new customer information (without commas)")
            vpPrintLine()

            vpPrint("First name:  ")
            first_name = Console.ReadLine()                      ' Get item title.


            vpPrint("Last name:  ")
            last_name = Console.ReadLine()                 ' ... and other info.
            vpPrint("Address:  ")
            address = Console.ReadLine()
            vpPrint("City:  ")
            city = Console.ReadLine()
            vpPrint("Province:  ")
            province = Console.ReadLine()
            vpPrint("Postal Code:  ")
            postalcode = Console.ReadLine()
            vpPrint("Credit Limit:  ")
            credit_limit = Double.Parse(Console.ReadLine())
            vpPrint("Email:  ")
            email = Console.ReadLine()
            vpPrintLine()
            ' Write record to database file.

            Dim newCust As New DLL_Library.IOTS.Customer(cust._custId,
                             first_name, last_name, address, city, province, postalcode, credit_limit, email)
            custBook.customerList.Remove(cust)
            custBook.AddCustomer(newCust)
            vpPrintLine()
            vpPrintLine("Customer modified!")
        Catch ex As Exception
            vpPrintLine("Failed to modify this customer.")
            vpPrintLine("Possible reason: " & ex.Message)
        End Try
        vpViewPrint()                                               ' Disable viewport and update status line.
        vpLocate(25, 1)
        vpPrint("Press Enter to return to main menu...")
        Console.ReadKey()

    End Sub

    Private Sub EditOrder(ByRef orderBook As DLL_Library.IOTS.OrderHolder,
                          ByVal order As DLL_Library.IOTS.Order,
                          ByVal productBook As DLL_Library.IOTS.ProductHolder,
                         ByVal custBook As DLL_Library.IOTS.CustomerHolder)
        Dim orderId As Long = order._orderNumber

        vpViewPrint(3, 33)                                  ' Enable viewport (lines 5-23).
        vpCls(2)                                            ' Clear viewport for choice prompts.
        'display order information
        vpPrintLine("Order:") '  cust.ToString()
        vpPrintLine(orderBook.GetOrderDetails(order).ToString)                                       ' Prompt for data.

        Try
            Dim productId As String = ""
            Dim numberOrdered As String = ""
            Dim discount As Integer
            Dim custId As Long

            ' Print message on status line.
            ' Prompt for data.
            vpPrintLine("Enter new order information (without commas)")
            vpPrintLine()

            vpPrint("Order Date(dd/mm/yyyy):  ")
            Dim orderDate As String = Console.ReadLine()
            vpPrint("Shipping Date(dd/mm/yyyy):  ")
            Dim shipDate As String = Console.ReadLine()

            vpPrint("ProductId:  ")
            productId = Console.ReadLine()                      ' Get item title.

            vpPrint("Quantity for order:  ")
            numberOrdered = Console.ReadLine()                 ' ... and other info.

            vpPrint("Customer ID:  ")
            custId = Integer.Parse(Console.ReadLine())                 ' ... and other info.

            vpPrint("Discount:  ")
            discount = Double.Parse(Console.ReadLine())

            vpPrintLine()
            ' Write record to database file.
            Dim newOrder As New DLL_Library.IOTS.Order(order._orderNumber, orderDate, shipDate, numberOrdered, productId, discount, custId)
            If orderBook.validateOrder(newOrder, custBook, productBook) Then
                orderBook.deleteOrderById(order._orderNumber)
                orderBook.addOrder(newOrder, custBook, productBook)
            End If
            vpPrintLine()
            vpPrintLine("Order modified!")
        Catch ex As Exception
            vpPrintLine("Failed to modify this order.")
            vpPrintLine("Possible reason: " & ex.Message)
        End Try


        vpViewPrint()                                               ' Disable viewport and update status line.
        vpLocate(25, 1)
        vpPrint("Press Enter to return to main menu...")
        Console.ReadKey()

    End Sub

    Public Sub EditProduct(ByRef productBook As DLL_Library.IOTS.ProductHolder,
                           ByVal product As DLL_Library.IOTS.Product)
        Try

            vpViewPrint(5, 23)        ' Enable viewport (lines 5-23).
            vpCls(2)                                            ' Clear viewport for choice prompts.
            'display product information
            vpPrintLine("Product:") '  product.ToString()
            vpPrintLine(product.GetDetails)                        ' Prompt for data.

            Dim productId As String = ""
            Dim description As String = ""
            Dim qtyOnHand As Integer
            Dim price As Double


            vpPrint("Enter Product data.  Type END for product ID to quit...")

            vpPrintLine()                                       ' Prompt for data.
            vpPrintLine("Enter new product information (without commas)")
            vpPrintLine()

            vpPrint("ProductId:  ")
            productId = Console.ReadLine()                      ' Get item title.
            vpPrint("Description:  ")
            description = Console.ReadLine()                 ' ... and other info.

            vpPrint("Quantity on hand:  ")
            qtyOnHand = Integer.Parse(Console.ReadLine())
            vpPrint("Price:  ")
            price = Double.Parse(Console.ReadLine())
            vpPrintLine()
            Dim newProd As New DLL_Library.IOTS.Product(productId, description, qtyOnHand, price)
            ' Write record to database file.

            productBook.productList.Remove(product)
            productBook.addProduct(newProd)


            vpPrintLine()
            vpPrintLine("Product modified!")
        Catch ex As Exception
            vpPrintLine("Failed to modify this product.")
            vpPrintLine("Possible reason:" & ex.Message)

        End Try
        vpViewPrint()                                               ' Disable viewport and update status line.
        vpLocate(25, 1)
        vpPrint("Press Enter to return to main menu...")
        Console.ReadKey()
    End Sub
    Private Sub GetMenuSelection(ByRef choice As Integer)
        ' Gets a menu choice from the user and returns it to the main
        ' program in the choice variable. The vpViewPrint calls are
        ' used to enable and disable the viewport area (lines 5-23).
        ' The information displayed here does not disturb the data in
        ' lines 1 through 4 and 24 through 25.

        choice = 0                                          ' Initialize choice to zero.

        vpLocate(25, 1)                                     ' Reset to line 25.
        vpPrint("Type a number between 1 and 8 and press Enter...")
        vpViewPrint(5, 23)                                  ' Enable viewport (lines 5-23).
        vpCls(2)                                            ' Clear viewport for choice prompts.

        vpPrintLine()                                       ' Prompt user for choice.
        vpPrintLine("SELECT an option:")
        vpPrintLine()
        vpPrintLine("  1) Customer Search")
        vpPrintLine("  2) Order Search")
        vpPrintLine("  3) Product Search")
        vpPrintLine("  4) Add Customer")
        vpPrintLine("  5) Add Order")
        vpPrintLine("  6) Add Product")
        vpPrintLine("  7) QUIT order tracking system")
        vpPrintLine()

        Do While (choice < 1) Or (choice > 8)               ' Choice must be integer between 1 and 8.
            vpPrint("Choice (1-7):  ")
            choice = CInt(Console.ReadLine())
        Loop

        vpCls(2)                                            ' Clear viewport for upcoming choice.
        vpViewPrint()                                       ' Disable viewport to clear status line.
        vpLocate(25, 1)
        vpPrint(Strings.StrDup(80, " "))                    ' Print a blank line.
    End Sub



    Private Sub DisplayHeader()
        ' Displays the status information on the first three lines of
        ' the screen and the two divinding lines that set off program
        ' information window.

        vpCls()                                         ' Clear the screen.


        vpPrintLine("                           SENECA ORDER TRACKING SYSTEM")
        vpPrintLine()

        vpPrint(Strings.StrDup(80, "-"))                ' Print the dividing lines ...

        vpColor(ConsoleColor.White)                     ' Set color back to default white.
    End Sub
End Module
