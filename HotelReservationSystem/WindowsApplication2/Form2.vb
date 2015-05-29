Public Class Form2
    Dim mainform As New Form1
    

    Private Sub Form2_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load



        For i As Integer = 0 To Int(Label5.Text)
            ComboBox1.Items.Add(i)
        Next
        For i As Integer = 0 To Int(Label6.Text)
            ComboBox2.Items.Add(i)
        Next
        For i As Integer = 0 To Int(Label7.Text)
            ComboBox3.Items.Add(i)
        Next
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        ComboBox3.SelectedIndex = 0

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
     

    

        Label8.Text = " Single rooms for "
        Label8.Text &= ComboBox1.SelectedItem * 2
        Label8.Text &= " People"

        Label11.Text = Convert.ToInt32(Label12.Text) - (ComboBox1.SelectedItem * 2) - ComboBox2.SelectedItem * 4 - ComboBox3.SelectedItem * 8
        If Int(Label11.Text) < 0 Then
            MsgBox("Over selection")
            ComboBox1.SelectedIndex = 0
            
            Return
        End If
        Label11.Text &= " people left"


    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Label9.Text = " Double rooms for "
        Label9.Text &= ComboBox2.SelectedItem * 4
        Label9.Text &= " People" '
        Label11.Text = Convert.ToInt32(Label12.Text) - Int(ComboBox1.SelectedItem * 2) - Int(ComboBox2.SelectedItem * 4) - Int(ComboBox3.SelectedItem * 8)
        If Int(Label11.Text) < 0 Then
            MsgBox("Over selection")

            ComboBox2.SelectedIndex = 0

            Return
        End If
        Label11.Text &= " people left"
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        Label10.Text = " Double rooms for "
        Label10.Text &= ComboBox3.SelectedItem * 8
        Label10.Text &= " People"
        Label11.Text = Convert.ToInt32(Label12.Text) - (ComboBox1.SelectedItem * 2) - ComboBox2.SelectedItem * 4 - ComboBox3.SelectedItem * 8
        If Int(Label11.Text) < 0 Then
            MsgBox("Over selection")
            
            ComboBox3.SelectedIndex = 0
            Return
        End If
        Label11.Text &= " people left"
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        'MsgBox(Form1.DataGridView1.RowCount)
        'Return

        Dim singleR As Integer = ComboBox1.SelectedIndex
        Dim DoubleR As Integer = ComboBox2.SelectedIndex
        Dim Quadruple As Integer = ComboBox3.SelectedIndex
        'MsgBox(Form1.DataGridView1.RowCount - 1)
        'MsgBox(Form1.DataGridView1.Rows(33).Cells(2).Value.ToString & singleR)

        For i As Integer = 0 To Form1.DataGridView1.RowCount - 2
            If (Form1.DataGridView1.Rows(i).Selected = True And i > 1) Then
                Form1.DataGridView1.Rows(i).Selected = False
            End If
            If Form1.DataGridView1.Rows(i).Cells(2).Value.ToString = "Single" And singleR > 0 Then



                Form1.DataGridView1.Rows(i).Selected = True

                singleR = singleR - 1




            End If
            If Form1.DataGridView1.Rows(i).Cells(2).Value.ToString = "Double" And DoubleR > 0 Then
                Form1.DataGridView1.Rows(i).Selected = True
                DoubleR = DoubleR - 1

            End If
            If Form1.DataGridView1.Rows(i).Cells(2).Value.ToString = "Quadruple" And Quadruple > 0 Then
                Form1.DataGridView1.Rows(i).Selected = True
                Quadruple = Quadruple - 1

            End If
        Next
        Me.Close()


    End Sub

    

   
End Class