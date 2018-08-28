Public Class PersonasGrid

    
    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        'boton agrega objeto
        PersonasForm.operacion = "nuevo"
        PersonasForm.Text = "Alta de Personas"
        PersonasForm.Show()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        'boton modificar
        PersonasForm.operacion = "modifica"
        PersonasForm.Text = "Modificar Persona"
        PersonasForm.indice = DataGridView1.CurrentRow.Index
        llenarform()
        PersonasForm.Show()
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        'boton eliminar 
        PersonasForm.operacion = "elimina"
        PersonasForm.Text = "Elimina Persona"
        PersonasForm.indice = DataGridView1.CurrentRow.Index
        llenarForm()
        PersonasForm.Show()
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        'boton salir
        Me.Close()
    End Sub
    Private Sub llenarform()
        'funcion que llena la grilla con los datos de base de datos
        PersonasForm.TextBox1.Text = DataGridView1.CurrentRow.Cells(1).Value.ToString
        PersonasForm.TextBox2.Text = DataGridView1.CurrentRow.Cells(4).Value.ToString
        PersonasForm.TextBox3.Text = DataGridView1.CurrentRow.Cells(5).Value.ToString
        PersonasForm.TextBox4.Text = DataGridView1.CurrentRow.Cells(2).Value.ToString
        PersonasForm.TextBox5.Text = DataGridView1.CurrentRow.Cells(0).Value.ToString

        PersonasForm.ComboBox2.Text = DataGridView1.CurrentRow.Cells(6).Value.ToString
        PersonasForm.ComboBox1.Text = DataGridView1.CurrentRow.Cells(3).Value.ToString

    End Sub
    Private Sub personasGrid_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'lst o lista lo refleja en la dataGridView1
        DataGridView1.DataSource = lst
    End Sub

   
End Class
