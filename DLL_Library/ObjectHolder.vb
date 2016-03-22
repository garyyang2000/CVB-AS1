Namespace IOTS
    Public MustInherit Class ObjectHolder

        Public MustOverride Sub loadFromFile(ByVal filename As String)
        Public MustOverride Sub loadFromSOAP(ByVal filename As String)
        Public MustOverride Sub SyncToFile(ByVal fileName As String)
        Public MustOverride Sub SyncToSOAP(ByVal fileName As String)


    End Class
End Namespace