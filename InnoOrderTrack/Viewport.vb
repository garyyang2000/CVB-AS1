Module Viewport

    ' As the Console class/object from .Net does not implements
    ' VIEW PRINT (clipped area) functionality, this module tries
    ' to simulate similar behaviour.

    ' Properties to store the viewport lines.
    Private topLine As Integer
    Private bottomLine As Integer = 25

    Public Sub vpInitConsole()
        ' Special procedure to set buffer to window size.

        Console.SetBufferSize(Console.WindowWidth, Console.WindowHeight)
    End Sub

    Public Sub vpCls(Optional ByVal type As Integer = 1)
        ' This procedure clears the viewport, screen.

        Select Case type
            Case 1
                Console.Clear()                                 ' Clear the full screen.
                Console.SetCursorPosition(0, 0)                 ' Also reset the cursor position.
            Case 2
                Dim tmpBgColor = Console.BackgroundColor        ' Temporarily store the console colors.
                Dim tmpFgColor = Console.ForegroundColor

                Console.ResetColor()                            ' Reset to the default colors.
                Console.SetCursorPosition(0, topLine)           ' Set the cursor to the topleft corner.

                For cnt As Integer = topLine To bottomLine      ' Clear only the viewport itself.
                    Console.Write(Strings.StrDup(80, " "))
                Next

                Console.BackgroundColor = tmpBgColor            ' Restore original colors.
                Console.ForegroundColor = tmpFgColor

                ' Reset the cursor to the topleft corner of the viewport again.
                Console.SetCursorPosition(0, topLine)
        End Select

    End Sub

    Public Sub vpColor(ByVal color As ConsoleColor)
        ' Sets the foreground color based on the provided parameter.

        Console.ForegroundColor = color
    End Sub

    Public Sub vpPrintLine(Optional ByVal text As String = "")
        ' Writes the provided string and a newline character to the console.

        ' First check whether the printing action is within the boundaries of
        ' the viewport.
        If Console.CursorTop > bottomLine Then              ' It is below the viewport.
            ' Move a copy of the viewport output, one line up, overwriting the top line.
            Console.MoveBufferArea(0, (topLine + 1), (Console.WindowWidth - 1), (bottomLine - topLine), 0, topLine)

            ' Set the cursor at one line higher.
            Console.CursorTop = Console.CursorTop - 1
        End If

        Console.WriteLine(text)
    End Sub

    Public Sub vpPrint(Optional ByVal text As String = "")
        ' Writes the provided string to the console.

        ' When a string is printed on the last line of the window,
        ' it could happen that the buffer will be scrolled down one line.
        If Console.CursorTop = (Console.WindowHeight - 1) Then                      ' Are we printing on the last line?
            If (Console.CursorLeft + text.Length) >= Console.WindowWidth Then        ' Yes.
                ' Calculate the clipsize.
                ' Take the window width, subtract the start point of the string
                ' and subtract the bottomright character.
                Dim clipSize As Integer = Console.WindowWidth - Console.CursorLeft - 1

                Console.Write(Strings.Left(text, clipSize))                         ' Print the clipped line.
            Else
                ' First check whether the printing action is within the boundaries of
                ' the viewport.
                If Console.CursorTop > bottomLine Then              ' It is below the viewport.
                    ' Move a copy of the viewport output, one line up, overwriting the top line.
                    Console.MoveBufferArea(0, topLine + 1, Console.WindowWidth, bottomLine, 0, topLine)
                End If

                Console.Write(text)                                                 ' No need for clipping, just print the text.
            End If
        Else
            ' First check whether the printing action is within the boundaries of
            ' the viewport.
            If Console.CursorTop > bottomLine Then              ' It is below the viewport.
                ' Move a copy of the viewport output, one line up, overwriting the top line.
                Console.MoveBufferArea(0, (topLine + 1), (Console.WindowWidth - 1), (bottomLine - topLine), 0, topLine)

                ' Set the cursor at one line higher.
                Console.CursorTop = Console.CursorTop - 1
            End If

            Console.Write(text)                                                     ' No, just print the text.
        End If

    End Sub

    Public Sub vpLocate(ByVal row As Integer, ByVal column As Integer)
        ' Sets the cursor position to the column and row given in the parameters.
        ' Both row and column start counting from 1. The procedure will
        ' adapt the values to the correct settings for the console.

        Console.SetCursorPosition((column - 1), (row - 1))
    End Sub

    Public Sub vpViewPrint(Optional ByVal top As Integer = 1, Optional ByVal bottom As Integer = 25)
        ' Sets the definition of the viewport.
        ' The top and bottom line are already corrected for translation to the console.

        topLine = top - 1
        bottomLine = bottom - 1

        Console.SetCursorPosition(0, topLine)
    End Sub

End Module
