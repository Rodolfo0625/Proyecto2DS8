Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True
        cmbPrefijo.SelectedIndex = 7
    End Sub

    Public Function VerificarCampos() As Boolean
        If cmbPrefijo.Text = "" Then
            MsgBox("Porfavor ingresa el tu cedula correctamenete")
            Return False
        ElseIf tbTomo.Text = "" Then
            MsgBox("Porfavor ingresa el tu cedula correctamenete")
            Return False
        ElseIf tbAsiento.Text = "" Then
            MsgBox("Porfavor ingresa el tu cedula correctamenete")
            Return False
        ElseIf tbNombre1.Text = "" Then
            MsgBox("Porfavor ingresa tu primer nombre")
            Return False
        ElseIf tbApellido1.Text = "" Then
            MsgBox("Porfavor ingresa tu primer apellido")
            Return False
        ElseIf tbSalarioHora.Text = "" Then
            MsgBox("Porfavor ingresa el salario por hora")
            Return False
        ElseIf tbHorasTrabajadas.Text = "" Then
            MsgBox("Porfavor ingresa las horas trabajadas")
            Return False
        End If
        Return True
    End Function


    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then



            If VerificarCampos() Then
                Dim empleado As New Empleado

                'Datos Personales
                empleado.nombre1 = tbNombre1.Text
                If tbNombre2.Text <> "" Then
                    empleado.nombre2 = tbNombre2.Text
                End If

                empleado.apellido1 = tbApellido1.Text
                If tbApellido2.Text <> "" Then
                    empleado.apellido2 = tbApellido2.Text
                End If

                If tbApellidoCasada.Text <> "" Then
                    empleado.apellido_casada = tbApellidoCasada.Text
                    empleado.usa_ape_casada = True
                Else
                    empleado.usa_ape_casada = False
                End If

                'Datos de salario
                empleado.shora = Decimal.Parse(tbSalarioHora.Text)
                empleado.htrabajadas = Decimal.Parse(tbHorasTrabajadas.Text)

                If tbHorasExtras.Text = "" Then
                    empleado.hextras = 0
                Else
                    empleado.hextras = Decimal.Parse(tbHorasExtras.Text)
                End If

                If tbDescuentos1.Text = "" Then
                    empleado.otros_descuentos1 = 0
                Else
                    empleado.otros_descuentos1 = Decimal.Parse(tbDescuentos1.Text)
                End If

                If tbDescuentos2.Text = "" Then
                    empleado.otros_descuentos2 = 0
                Else
                    empleado.otros_descuentos2 = Decimal.Parse(tbDescuentos2.Text)
                End If

                If tbDescuentos3.Text = "" Then
                    empleado.otros_descuentos3 = 0
                Else
                    empleado.otros_descuentos3 = Decimal.Parse(tbDescuentos3.Text)
                End If

                empleado.CalcularSalario()

                'Mostrar resultado en sus respectivos Text Box formateados a dos decimales'
                tbSueldoBruto.Text = empleado.sbruto.ToString("F2")
                tbSeguroSocial.Text = empleado.ssocial.ToString("F2")
                tbSeguroEducativo.Text = empleado.seducativo.ToString("F2")
                tbImpuestoRenta.Text = empleado.irenta.ToString("F2")
                tbSueldoNeto.Text = empleado.sneto.ToString("F2")
            End If
        End If
    End Sub

    Private Sub tbTomo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbTomo.KeyPress
        ' Verificar si la tecla presionada no es un dígito (0-9) o una tecla de control
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            ' Cancelar el evento de tecla presionada para evitar que se ingrese el carácter
            e.Handled = True
        End If

        If tbTomo.Text.Length = 4 AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If

    End Sub

    Private Sub tbAsiento_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbAsiento.KeyPress
        ' Verificar si la tecla presionada no es un dígito (0-9) o una tecla de control
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            ' Cancelar el evento de tecla presionada para evitar que se ingrese el carácter
            e.Handled = True
        End If

        If tbAsiento.Text.Length = 5 AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub tbNombre1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbNombre1.KeyPress
        ' Evitar teclas que no sean letras alfabéticas, espacios, teclas de control o caracteres con tilde
        If Not Char.IsLetter(e.KeyChar) AndAlso Not e.KeyChar = " " AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not e.KeyChar = "´" Then
            ' Si no es un carácter válido, suprimir el carácter ingresado
            e.Handled = True
        End If
    End Sub

    Private Sub tbNombre2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbNombre2.KeyPress
        ' Evitar teclas que no sean letras alfabéticas, espacios, teclas de control o caracteres con tilde
        If Not Char.IsLetter(e.KeyChar) AndAlso Not e.KeyChar = " " AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not e.KeyChar = "´" Then
            ' Si no es un carácter válido, suprimir el carácter ingresado
            e.Handled = True
        End If
    End Sub

    Private Sub tbApellido1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbApellido1.KeyPress
        ' Evitar teclas que no sean letras alfabéticas, espacios, teclas de control o caracteres con tilde
        If Not Char.IsLetter(e.KeyChar) AndAlso Not e.KeyChar = " " AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not e.KeyChar = "´" Then
            ' Si no es un carácter válido, suprimir el carácter ingresado
            e.Handled = True
        End If
    End Sub

    Private Sub tbApellido2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbApellido2.KeyPress
        ' Evitar teclas que no sean letras alfabéticas, espacios, teclas de control o caracteres con tilde
        If Not Char.IsLetter(e.KeyChar) AndAlso Not e.KeyChar = " " AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not e.KeyChar = "´" Then
            ' Si no es un carácter válido, suprimir el carácter ingresado
            e.Handled = True
        End If
    End Sub

    Private Sub tbApellidoCasada_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbApellidoCasada.KeyPress
        ' Evitar teclas que no sean letras alfabéticas, espacios, teclas de control o caracteres con tilde
        If Not Char.IsLetter(e.KeyChar) AndAlso Not e.KeyChar = " " AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not e.KeyChar = "´" Then
            ' Si no es un carácter válido, suprimir el carácter ingresado
            e.Handled = True
        End If
    End Sub

    Private Sub tbSalarioHora_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbSalarioHora.KeyPress
        ' Evitar teclas que no sean los digitos, el punto decimal o las teclas de control
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "." Then
            e.Handled = True
        End If

        ' Si el primer digito es cero, evitar ingresar otro cero como segundo digito'
        If tbSalarioHora.Text.Length = 1 AndAlso e.KeyChar = "0" Then
            If tbSalarioHora.Text = "0" Then
                e.Handled = True
            End If
        End If

        ' Evitar ingresar mas de un punto decimal
        If e.KeyChar = "." AndAlso (CType(sender, TextBox).Text.Contains(".") OrElse CType(sender, TextBox).SelectionStart = 0) Then
            e.Handled = True
        End If

        ' Evitar ingresar mas de dos digitos decimales'
        If tbSalarioHora.Text.Contains(".") AndAlso Char.IsDigit(e.KeyChar) Then
            Dim partesTexto() As String = tbSalarioHora.Text.Split(".")
            If partesTexto(1).Length = 2 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub tbHorasTrabajadas_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbHorasTrabajadas.KeyPress
        ' Evitar teclas que no sean los digitos o las teclas de control
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "." Then
            ' Cancelar el evento de tecla presionada para evitar que se ingrese el carácter
            e.Handled = True
        End If

        ' Si el primer digito es cero, evitar ingresar otro cero como segundo digito'
        If tbHorasTrabajadas.Text.Length = 1 AndAlso e.KeyChar = "0" Then
            If tbHorasTrabajadas.Text = "0" Then
                e.Handled = True
            End If
        End If

        ' Evitar ingresar mas de un punto decimal
        If e.KeyChar = "." AndAlso (CType(sender, TextBox).Text.Contains(".") OrElse CType(sender, TextBox).SelectionStart = 0) Then
            e.Handled = True
        End If

        ' Evitar ingresar mas de dos digitos decimales'
        If tbHorasTrabajadas.Text.Contains(".") AndAlso Char.IsDigit(e.KeyChar) Then
            Dim partesTexto() As String = tbHorasTrabajadas.Text.Split(".")
            If partesTexto(1).Length = 2 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub tbHorasExtras_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbHorasExtras.KeyPress
        ' Evitar teclas que no sean los digitos o las teclas de control
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "." Then
            ' Cancelar el evento de tecla presionada para evitar que se ingrese el carácter
            e.Handled = True
        End If

        ' Si el primer digito es cero, evitar ingresar otro cero como segundo digito'
        If tbHorasExtras.Text.Length = 1 AndAlso e.KeyChar = "0" Then
            If tbHorasExtras.Text = "0" Then
                e.Handled = True
            End If
        End If

        ' Evitar ingresar mas de un punto decimal
        If e.KeyChar = "." AndAlso (CType(sender, TextBox).Text.Contains(".") OrElse CType(sender, TextBox).SelectionStart = 0) Then
            e.Handled = True
        End If

        ' Evitar ingresar mas de dos digitos decimales'
        If tbHorasExtras.Text.Contains(".") AndAlso Char.IsDigit(e.KeyChar) Then
            Dim partesTexto() As String = tbHorasExtras.Text.Split(".")
            If partesTexto(1).Length = 2 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub tbDescuentos1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbDescuentos1.KeyPress
        ' Evitar teclas que no sean los digitos, el punto decimal o las teclas de control
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "." Then
            e.Handled = True
        End If

        ' Si el primer digito es cero, evitar ingresar otro cero como segundo digito'
        If tbDescuentos1.Text.Length = 1 AndAlso e.KeyChar = "0" Then
            If tbDescuentos1.Text = "0" Then
                e.Handled = True
            End If
        End If

        ' Evitar ingresar mas de un punto decimal
        If e.KeyChar = "." AndAlso (CType(sender, TextBox).Text.Contains(".") OrElse CType(sender, TextBox).SelectionStart = 0) Then
            e.Handled = True
        End If

        ' Evitar ingresar mas de dos digitos decimales'
        If tbDescuentos1.Text.Contains(".") AndAlso Char.IsDigit(e.KeyChar) Then
            Dim partesTexto() As String = tbDescuentos1.Text.Split(".")
            If partesTexto(1).Length = 2 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub tbDescuentos2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbDescuentos2.KeyPress
        ' Evitar teclas que no sean los digitos, el punto decimal o las teclas de control
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "." Then
            e.Handled = True
        End If

        ' Si el primer digito es cero, evitar ingresar otro cero como segundo digito'
        If tbDescuentos2.Text.Length = 1 AndAlso e.KeyChar = "0" Then
            If tbDescuentos1.Text = "0" Then
                e.Handled = True
            End If
        End If

        ' Evitar ingresar mas de un punto decimal
        If e.KeyChar = "." AndAlso (CType(sender, TextBox).Text.Contains(".") OrElse CType(sender, TextBox).SelectionStart = 0) Then
            e.Handled = True
        End If

        ' Evitar ingresar mas de dos digitos decimales'
        If tbDescuentos2.Text.Contains(".") AndAlso Char.IsDigit(e.KeyChar) Then
            Dim partesTexto() As String = tbDescuentos2.Text.Split(".")
            If partesTexto(1).Length = 2 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub tbDescuentos3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbDescuentos3.KeyPress
        ' Evitar teclas que no sean los digitos, el punto decimal o las teclas de control
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "." Then
            e.Handled = True
        End If

        ' Si el primer digito es cero, evitar ingresar otro cero como segundo digito'
        If tbDescuentos2.Text.Length = 1 AndAlso e.KeyChar = "0" Then
            If tbDescuentos1.Text = "0" Then
                e.Handled = True
            End If
        End If

        ' Evitar ingresar mas de un punto decimal
        If e.KeyChar = "." AndAlso (CType(sender, TextBox).Text.Contains(".") OrElse CType(sender, TextBox).SelectionStart = 0) Then
            e.Handled = True
        End If

        ' Evitar ingresar mas de dos digitos decimales'
        If tbDescuentos3.Text.Contains(".") AndAlso Char.IsDigit(e.KeyChar) Then
            Dim partesTexto() As String = tbDescuentos3.Text.Split(".")
            If partesTexto(1).Length = 2 Then
                e.Handled = True
            End If
        End If
    End Sub
End Class
