Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True
        cmbPrefijo.SelectedIndex = 7
        cmbGenero.SelectedIndex = 0
        cmbEstadoCivil.SelectedIndex = 0
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
        ElseIf cmbGenero.Text = "" Then
            MsgBox("Porfavor selecciona tu genero")
            Return False
        ElseIf cmbEstadoCivil.Text = "" Then
            MsgBox("Porfavor selecciona tu genero")
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
        ElseIf cmbTipoHorasExtra1.Text <> "" AndAlso tbHorasExtra1.Text = "" Then
            MsgBox("Porfavor ingresa la cantidad de las horas extras 1")
            Return False
        ElseIf cmbTipoHorasExtra2.Text <> "" AndAlso tbHorasExtra2.Text = "" Then
            MsgBox("Porfavor ingresa la cantidad de las horas extras 2")
            Return False
        ElseIf cmbTipoHorasExtra3.Text <> "" AndAlso tbHorasExtra3.Text = "" Then
            MsgBox("Porfavor ingresa la cantidad de las horas extras 3")
            Return False
        End If
        Return True
    End Function

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then

            If VerificarCampos() Then
                Dim empleado As New Empleado

                'DATOS PERSONALES
                empleado.prefijo = cmbPrefijo.Text
                empleado.tomo = tbTomo.Text
                empleado.asiento = tbAsiento.Text
                empleado.cedula = String.Format("{0}-{1}-{2}", empleado.prefijo, empleado.tomo, empleado.asiento)
                empleado.genero = cmbGenero.Text
                empleado.estado_civil = cmbEstadoCivil.Text

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
                    empleado.usa_apellido_casada = "Si"
                End If

                'SALARIO Y HORAS TRABJADAS
                empleado.shora = Decimal.Parse(tbSalarioHora.Text)
                empleado.htrabajadas = Decimal.Parse(tbHorasTrabajadas.Text)

                'HORAS EXTRAS'
                If tbHorasExtra1.Text <> "" Then
                    Dim tipo As String = cmbTipoHorasExtra1.Text
                    Dim cantidad = Decimal.Parse(tbHorasExtra1.Text)
                    empleado.AsignarHorasExtra1(tipo, cantidad)
                End If

                If tbHorasExtra2.Text <> "" Then
                    Dim tipo As String = cmbTipoHorasExtra2.Text
                    Dim cantidad = Decimal.Parse(tbHorasExtra2.Text)
                    empleado.AsignarHorasExtra2(tipo, cantidad)
                End If

                If tbHorasExtra3.Text <> "" Then
                    Dim tipo As String = cmbTipoHorasExtra3.Text
                    Dim cantidad = Decimal.Parse(tbHorasExtra3.Text)
                    empleado.AsignarHorasExtra3(tipo, cantidad)
                End If

                'DESCUENTOS
                If tbDescuento1.Text <> "" Then
                    empleado.descuento1 = tbDescuento1.Text
                End If

                If tbDescuento2.Text <> "" Then
                    empleado.descuento2 = tbDescuento2.Text
                End If

                If tbDescuento3.Text <> "" Then
                    empleado.descuento3 = tbDescuento3.Text
                End If

                'CALCULOS FINALES'
                empleado.CalcularSalario()

                'MOSTRAR RESULTADOS
                tbMontoHorasExtras1.Text = empleado.mhextra1
                tbMontoHorasExtras2.Text = empleado.mhextra2
                tbMontoHorasExtras3.Text = empleado.mhextra3
                tbSueldoBruto.Text = empleado.sbruto
                tbSeguroSocial.Text = empleado.seguro_social
                tbSeguroEducativo.Text = empleado.seguro_educativo
                tbImpuestoRenta.Text = empleado.impuesto_renta
                tbSueldoNeto.Text = empleado.sneto

                'imprimir una lista de todos los atributos del empleado (SOLO PARA VERIFICAR)
                'empleado.ImprimirAtributos()

                'Registrar datos del empleado
                'empleado.RegistrarDatos()
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

    Private Sub tbHorasExtra1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbHorasExtra1.KeyPress
        ' Evitar teclas que no sean los digitos, el punto decimal o las teclas de control
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "." Then
            e.Handled = True
        End If

        ' Evitar Ingresar 0 como primer digito'
        If tbHorasExtra1.Text.Length = 0 AndAlso e.KeyChar = "0" Then
            e.Handled = True
        End If

        ' Evitar ingresar mas de un punto decimal
        If e.KeyChar = "." AndAlso (CType(sender, TextBox).Text.Contains(".") OrElse CType(sender, TextBox).SelectionStart = 0) Then
            e.Handled = True
        End If

        ' Evitar ingresar mas de dos digitos decimales'
        If tbHorasExtra1.Text.Contains(".") AndAlso Char.IsDigit(e.KeyChar) Then
            Dim partesTexto() As String = tbHorasExtra1.Text.Split(".")
            If partesTexto(1).Length = 2 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub tbHorasExtra2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbHorasExtra2.KeyPress
        ' Evitar teclas que no sean los digitos, el punto decimal o las teclas de control
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "." Then
            e.Handled = True
        End If

        ' Evitar Ingresar 0 como primer digito'
        If tbHorasExtra2.Text.Length = 0 AndAlso e.KeyChar = "0" Then
            e.Handled = True
        End If

        ' Evitar ingresar mas de un punto decimal
        If e.KeyChar = "." AndAlso (CType(sender, TextBox).Text.Contains(".") OrElse CType(sender, TextBox).SelectionStart = 0) Then
            e.Handled = True
        End If

        ' Evitar ingresar mas de dos digitos decimales'
        If tbHorasExtra2.Text.Contains(".") AndAlso Char.IsDigit(e.KeyChar) Then
            Dim partesTexto() As String = tbHorasExtra2.Text.Split(".")
            If partesTexto(1).Length = 2 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub tbHorasExtra3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbHorasExtra3.KeyPress
        ' Evitar teclas que no sean los digitos, el punto decimal o las teclas de control
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "." Then
            e.Handled = True
        End If

        ' Evitar Ingresar 0 como primer digito'
        If tbHorasExtra3.Text.Length = 0 AndAlso e.KeyChar = "0" Then
            e.Handled = True
        End If

        ' Evitar ingresar mas de un punto decimal
        If e.KeyChar = "." AndAlso (CType(sender, TextBox).Text.Contains(".") OrElse CType(sender, TextBox).SelectionStart = 0) Then
            e.Handled = True
        End If

        ' Evitar ingresar mas de dos digitos decimales'
        If tbHorasExtra3.Text.Contains(".") AndAlso Char.IsDigit(e.KeyChar) Then
            Dim partesTexto() As String = tbHorasExtra3.Text.Split(".")
            If partesTexto(1).Length = 2 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub tbDescuentos1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbDescuento1.KeyPress
        ' Evitar teclas que no sean los digitos, el punto decimal o las teclas de control
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "." Then
            e.Handled = True
        End If

        ' Si el primer digito es cero, evitar ingresar otro cero como segundo digito'
        If tbDescuento1.Text.Length = 1 AndAlso e.KeyChar = "0" Then
            If tbDescuento1.Text = "0" Then
                e.Handled = True
            End If
        End If

        ' Evitar ingresar mas de un punto decimal
        If e.KeyChar = "." AndAlso (CType(sender, TextBox).Text.Contains(".") OrElse CType(sender, TextBox).SelectionStart = 0) Then
            e.Handled = True
        End If

        ' Evitar ingresar mas de dos digitos decimales'
        If tbDescuento1.Text.Contains(".") AndAlso Char.IsDigit(e.KeyChar) Then
            Dim partesTexto() As String = tbDescuento1.Text.Split(".")
            If partesTexto(1).Length = 2 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub tbDescuentos2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbDescuento2.KeyPress
        ' Evitar teclas que no sean los digitos, el punto decimal o las teclas de control
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "." Then
            e.Handled = True
        End If

        ' Si el primer digito es cero, evitar ingresar otro cero como segundo digito'
        If tbDescuento2.Text.Length = 1 AndAlso e.KeyChar = "0" Then
            If tbDescuento1.Text = "0" Then
                e.Handled = True
            End If
        End If

        ' Evitar ingresar mas de un punto decimal
        If e.KeyChar = "." AndAlso (CType(sender, TextBox).Text.Contains(".") OrElse CType(sender, TextBox).SelectionStart = 0) Then
            e.Handled = True
        End If

        ' Evitar ingresar mas de dos digitos decimales'
        If tbDescuento2.Text.Contains(".") AndAlso Char.IsDigit(e.KeyChar) Then
            Dim partesTexto() As String = tbDescuento2.Text.Split(".")
            If partesTexto(1).Length = 2 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub tbDescuentos3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbDescuento3.KeyPress
        ' Evitar teclas que no sean los digitos, el punto decimal o las teclas de control
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "." Then
            e.Handled = True
        End If

        ' Si el primer digito es cero, evitar ingresar otro cero como segundo digito'
        If tbDescuento2.Text.Length = 1 AndAlso e.KeyChar = "0" Then
            If tbDescuento1.Text = "0" Then
                e.Handled = True
            End If
        End If

        ' Evitar ingresar mas de un punto decimal
        If e.KeyChar = "." AndAlso (CType(sender, TextBox).Text.Contains(".") OrElse CType(sender, TextBox).SelectionStart = 0) Then
            e.Handled = True
        End If

        ' Evitar ingresar mas de dos digitos decimales'
        If tbDescuento3.Text.Contains(".") AndAlso Char.IsDigit(e.KeyChar) Then
            Dim partesTexto() As String = tbDescuento3.Text.Split(".")
            If partesTexto(1).Length = 2 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub cmbGenero_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbGenero.SelectedIndexChanged
        If cmbGenero.Text = "Femenino" Then
            tbApellidoCasada.Enabled = True
            cmbEstadoCivil.Items.Clear()
            cmbEstadoCivil.Items.Add("Soltera")
            cmbEstadoCivil.Items.Add("Casada")
            cmbEstadoCivil.Items.Add("Viuda")
            cmbEstadoCivil.Items.Add("Divorciada")
            cmbEstadoCivil.SelectedIndex = 0
        Else
            tbApellidoCasada.Enabled = False
            cmbEstadoCivil.Items.Clear()
            cmbEstadoCivil.Items.Add("Soltero")
            cmbEstadoCivil.Items.Add("Casado")
            cmbEstadoCivil.Items.Add("Viudo")
            cmbEstadoCivil.Items.Add("Divorciado")
            cmbEstadoCivil.SelectedIndex = 0
        End If
    End Sub
    Private Sub cmbEstadoCivil_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbEstadoCivil.SelectedIndexChanged
        If cmbGenero.Text = "Femenino" AndAlso cmbEstadoCivil.Text <> "Soltera" Then
            tbApellidoCasada.Enabled = True
        ElseIf cmbGenero.Text = "Femenino" AndAlso cmbEstadoCivil.Text = "Soltera" Then
            tbApellidoCasada.Enabled = False
        End If
    End Sub


    Private Sub cmbTipoHorasExtra1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbTipoHorasExtra1.SelectedIndexChanged
        If cmbTipoHorasExtra1.Text <> "" Then
            tbHorasExtra1.Enabled = True
        Else
            tbHorasExtra1.Text = ""
            tbHorasExtra1.Enabled = False
        End If
    End Sub

    Private Sub cmbTipoHorasExtra2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbTipoHorasExtra2.SelectedIndexChanged
        If cmbTipoHorasExtra2.Text <> "" Then
            tbHorasExtra2.Enabled = True
        Else
            tbHorasExtra2.Text = ""
            tbHorasExtra2.Enabled = False
        End If
    End Sub

    Private Sub cmbTipoHorasExtra3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbTipoHorasExtra3.SelectedIndexChanged
        If cmbTipoHorasExtra3.Text <> "" Then
            tbHorasExtra3.Enabled = True
        Else
            tbHorasExtra3.Text = ""
            tbHorasExtra3.Enabled = False
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        tbTomo.Text = ""
        tbAsiento.Text = ""
        cmbGenero.SelectedIndex = 0

        cmbEstadoCivil.Items.Clear()
        cmbEstadoCivil.Items.Add("Soltero")
        cmbEstadoCivil.Items.Add("Casado")
        cmbEstadoCivil.Items.Add("Viudo")
        cmbEstadoCivil.Items.Add("Divorciado")
        cmbEstadoCivil.SelectedIndex = 0

        tbNombre1.Text = ""
        tbNombre2.Text = ""
        tbApellido1.Text = ""
        tbApellido2.Text = ""
        tbApellidoCasada.Text = ""
        tbSalarioHora.Text = ""
        tbHorasTrabajadas.Text = ""
        cmbTipoHorasExtra1.SelectedIndex = 0
        tbHorasExtra1.Text = ""
        cmbTipoHorasExtra2.SelectedIndex = 0
        tbHorasExtra2.Text = ""
        cmbTipoHorasExtra3.SelectedIndex = 0
        tbHorasExtra3.Text = ""
        tbDescuento1.Text = ""
        tbDescuento2.Text = ""
        tbDescuento3.Text = ""
        tbMontoHorasExtras1.Text = ""
        tbMontoHorasExtras2.Text = ""
        tbMontoHorasExtras3.Text = ""
        tbSueldoBruto.Text = ""
        tbSeguroSocial.Text = ""
        tbSeguroEducativo.Text = ""
        tbImpuestoRenta.Text = ""
        tbSueldoNeto.Text = ""
    End Sub
End Class
