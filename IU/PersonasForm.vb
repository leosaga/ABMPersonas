Public Class PersonasForm


    Dim MiPersona As New PersonasClass
    Dim operacion_ As String

    Public Property operacion() As String
        Get
            Return operacion_
        End Get
        Set(ByVal value As String)
            operacion_ = value
        End Set
    End Property
    Dim indice_ As Byte
    Public Property indice() As Byte
        Get
            Return indice_
        End Get
        Set(ByVal value As Byte)
            indice_ = value
        End Set
    End Property

   
    
    Private Sub PersonasForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ComboBox1.DataSource = MiPersona.provincias
        ComboBox2.DataSource = MiPersona.documentos
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
    End Sub


    Private Sub Cancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancelar.Click
        Me.Close()

    End Sub

    Private Sub Aceptar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Aceptar.Click
        'para validar los texbox que no queden en blanco primer caso
        ' no deja pasar sin poner nada

        MiPersona.Id = TextBox5.Text
        MiPersona.Nombre = TextBox1.Text
        MiPersona.CodPostal = TextBox3.Text
        MiPersona.NumDocumento = TextBox4.Text
        MiPersona.Direccion = TextBox2.Text
        MiPersona.IdProvincia = ComboBox1.SelectedIndex
        MiPersona.TipoDocumento = ComboBox2.SelectedIndex


        PersonasGrid.DataGridView1.Refresh()

        If ComboBox1.SelectedIndex = -1 Then
            MsgBox("elija una provincia")
            Exit Sub
        End If
        ' no deja pasar sin poner nada
        If ComboBox2.SelectedIndex = -1 Then
            MsgBox("elija un tipo de documento")
            Exit Sub
        End If
        ' no deja pasar sin poner nada
        If TextBox1.Text = "" Then
            MsgBox("ingrese un nombre")
            Exit Sub
        End If
        ' no deja pasar sin poner nada
        If TextBox2.Text.Trim = "" Then
            MsgBox("ingrese una direccion ")
            Exit Sub
        End If
        ' no deja pasar sin poner nada
        If TextBox3.Text = "" Then
            MsgBox("ingrese un codigo postal")
            Exit Sub
        End If
        ' no deja pasar sin poner nada
        If TextBox4.Text = "" Then
            MsgBox("ingrese un documento")
            Exit Sub
        End If



        'esta funcion es para ver si es un nuevo objeto, eliminar o moficar
        Select Case operacion_
            Case "nuevo"
                lst.InsertarPersona(MiPersona)
            Case "elimina"
                lst.RemoveAt(indice_)
            Case "modifica"
                lst.Item(indice_) = MiPersona
                lst.ActualizarPersona(MiPersona)
        End Select
        'en la operacion elimina o modifica toma al objeto por el id
        If operacion_ = "elimina" Or operacion_ = "modifica" Then
            MiPersona.Id = TextBox5.Text
        End If





        Me.Close()

    End Sub

    
   
   
    
End Class