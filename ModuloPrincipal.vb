Module ModuloPrincipal
    Public lst As PersonasCollection
    Sub main()
        lst = New PersonasCollection
        'clase q levanta en la pantalla
        Application.Run(PersonasGrid)
    End Sub
End Module
