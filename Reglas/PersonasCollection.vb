Imports System.ComponentModel
Imports System.Text

Public Class PersonasCollection

    Inherits BindingList(Of PersonasClass)
    Protected Overrides Sub OnAddingNew(ByVal e As System.ComponentModel.AddingNewEventArgs)
        MyBase.OnAddingNew(e)
        e.NewObject = New PersonasClass


    End Sub


    Protected Overrides ReadOnly Property SupportsSearchingCore() As Boolean
        Get
            Return True
        End Get
    End Property

    Protected Overrides Function FindCore(ByVal prop As PropertyDescriptor, ByVal key As Object) As Integer
        For Each carrera In Me
            If prop.GetValue(carrera).Equals(key) Then
                Return Me.IndexOf(carrera)
            End If
        Next

        Return -1
    End Function
    Public Sub New()
        Me.TraerPersonas()

    End Sub

    Public Function TraerPersonas() As PersonasCollection
        'crea la intancia de base de datos
        Dim ObjBaseDatos As New BaseDatosClass
        Dim MiDataTable As New DataTable
        Dim MiPersona As PersonasClass

        ObjBaseDatos.objTabla = "Personas"
        'devuelve los datos de la base de dato
        MiDataTable = ObjBaseDatos.Consultar
        'por cada dr (fila)llena cada persona 

        For Each dr As DataRow In MiDataTable.Rows
            'crea la instancia por cada campo
            MiPersona = New PersonasClass

            MiPersona.Id = CInt(dr("Id"))

            MiPersona.Nombre = dr("Nombre")

            MiPersona.Direccion = (dr("Direccion"))

            MiPersona.CodPostal = CInt(dr("CodPostal"))

            MiPersona.IdProvincia = CInt(dr("IdProvincia"))

            MiPersona.TipoDocumento = (dr("TipoDocumento"))

            MiPersona.NumDocumento = CInt(dr("NumDocumento"))

            Me.Add(MiPersona)

        Next

        Return Me

    End Function
    Public Sub InsertarPersona(ByVal MiPersona As PersonasClass)

        Dim ObjBasedeDato As New BaseDatosClass
        'busca la tabla personas 
        ObjBasedeDato.objTabla = "Personas"

        Dim vsql As New StringBuilder
        ' vsql.Append("(Id")
        vsql.Append("(Nombre")
        vsql.Append(", Direccion")
        vsql.Append(", CodPostal")
        vsql.Append(", IdProvincia")
        vsql.Append(", TipoDocumento")
        vsql.Append(", NumDocumento)")

        vsql.Append(" VALUES ")

        ' vsql.Append("('" & MiPersona.id & "'")
        vsql.Append("('" & MiPersona.Nombre & "'")
        vsql.Append(", '" & MiPersona.Direccion & "'")
        vsql.Append(", '" & MiPersona.CodPostal & "'")
        vsql.Append(", '" & MiPersona.IdProvincia & "'")
        vsql.Append(", '" & MiPersona.TipoDocumento & "'")
        vsql.Append(", '" & MiPersona.NumDocumento & "')")

        MiPersona.id = ObjBasedeDato.Insertar(vsql.ToString)
        Me.Add(MiPersona)


    End Sub

   


    Public Sub EliminarPersona(ByVal MiPersona As PersonasClass)

        'Instancio el el Objeto BaseDatosClass para acceder al la base personas.
        Dim objBaseDatos As New BaseDatosClass
        objBaseDatos.objTabla = "Personas"

        'Ejecuta el método base eliminar.
        Dim resultado As Boolean
        resultado = objBaseDatos.Eliminar(MiPersona.Id)

        If resultado Then

            'Creates a new collection and assign it the properties for modulo.
            Dim properties As PropertyDescriptorCollection = TypeDescriptor.GetProperties(MiPersona)

            'Sets an PropertyDescriptor to the specific property Id.
            Dim myProperty As PropertyDescriptor = properties.Find("Id", False)

            Me.RemoveAt(Me.FindCore(myProperty, MiPersona.Id))
        Else
            MessageBox.Show("No fue posible eliminar el registro.")
        End If

    End Sub

    Public Sub ActualizarPersona(ByVal MiPersona As PersonasClass)

        'Instancio el el Objeto BaseDatosClass para acceder al la base personas.
        Dim objBaseDatos As New BaseDatosClass
        objBaseDatos.objTabla = "Personas"

        Dim vSQL As New StringBuilder
        Dim vResultado As Boolean = False

        'vSQL.Append("Id='" & MiPersona.Id.ToString & "'")
        vSQL.Append("TipoDocumento='" & MiPersona.TipoDocumento & "'")
        vSQL.Append(",NumDocumento='" & MiPersona.NumDocumento.ToString & "'")
        vSQL.Append(",Nombre='" & MiPersona.nombre & "'")
        vSQL.Append(",Direccion='" & MiPersona.direccion & "'")
        vSQL.Append(",IdProvincia='" & MiPersona.idProvincia.ToString & "'")
        vSQL.Append(",CodPostal='" & MiPersona.codPostal.ToString & "'")

        'Actualizo la tabla personas con el Id.
        Dim resultado As Boolean
        'El método actualizar es una función que devuelve True o False
        'Según como haya resultado la operación sobre la tabla SQL.
        resultado = objBaseDatos.Actualizar(vSQL.ToString, MiPersona.Id)

        If resultado Then
            'El código a continuación sirve para localizar el ítem en la lista
            'en este caso un persona.
            ' Creates a new collection and assign it the properties for modulo.
            Dim properties As PropertyDescriptorCollection = TypeDescriptor.GetProperties(MiPersona)

            'Sets an PropertyDescriptor to the specific property Id.
            Dim myProperty As PropertyDescriptor = properties.Find("Id", False)
            Me.Items.Item(Me.FindCore(myProperty, MiPersona.Id)) = MiPersona
        Else
            MessageBox.Show("No fue posible modificar el registro.")
        End If

    End Sub

End Class
