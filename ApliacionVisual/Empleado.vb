Imports MySql.Data.MySqlClient

Public Class Empleado
    Public prefijo As String
    Public tomo As String
    Public asiento As String
    Public cedula As String
    Public genero As String
    Public nombre1 As String
    Public nombre2 As String = ""
    Public apellido1 As String
    Public apellido2 As String = ""
    Public estado_civil As String
    Public apellido_casada As String = ""
    Public usa_apellido_casada As String = "No"

    Public shora As Decimal
    Public htrabajadas As Decimal

    'Horas extras 1'
    Public thextra1 As String = ""
    Public hextra1 As Decimal = 0
    Public mhextra1 As Decimal = 0

    'Horas extras 2'
    Public thextra2 As String = ""
    Public hextra2 As Decimal = 0
    Public mhextra2 As Decimal = 0

    'Horas extras 3'
    Public thextra3 As String = ""
    Public hextra3 As Decimal = 0
    Public mhextra3 As Decimal = 0


    Public sbruto As Decimal
    Public seguro_social As Decimal
    Public seguro_educativo As Decimal
    Public impuesto_renta As Decimal
    Public descuento1 As Decimal = 0
    Public descuento2 As Decimal = 0
    Public descuento3 As Decimal = 0
    Public sneto As Decimal

    Public Function ObtenerPorcentajeHoraExtra(ByVal tipoHoraExtra As String) As Double
        tipoHoraExtra = tipoHoraExtra.TrimEnd()
        Dim porcentaje As Double = 0
        Select Case tipoHoraExtra
            Case "Horas Extra Diurna"
                porcentaje = 1.25
            Case "Horas Extra Nocturna"
                porcentaje = 1.5
            Case "Horas Extra Mixta: Diurna - Nocturna"
                porcentaje = 1.5
            Case "Horas Extra Mixta: Nocturna - Diurna"
                porcentaje = 1.75
            Case "Fiesta Nacional o Duelo Nacional"
                porcentaje = 2.5
            Case "Hora Domingo o Descanso Semanal"
                porcentaje = 1.5
            Case "Horas Extra Diurna con exceso de 3 Horas diarias ó 9 Semanales"
                porcentaje = 1.25 * 1.75
            Case "Horas Extra Nocturna con exceso de 3 Horas diarias ó 9 Semanales"
                porcentaje = 1.5 * 1.75
            Case "Horas Extra Mixta: Diurna - Nocturna con exceso de 3 Horas diarias ó 9 Semanales"
                porcentaje = 1.5 * 1.75
            Case "Horas Extra Mixta: Nocturna - Diurna con exceso de 3 Horas diarias ó 9 Semanales"
                porcentaje = 1.75 * 1.75
            Case "Horas Extra Fiesta Nacional ó Duelo Nacional Diurna"
                porcentaje = 2.5 * 1.25
            Case "Horas Extra Fiesta Nacional ó Duelo Nacional Nocturno"
                porcentaje = 2.5 * 1.5
            Case "Horas Extra Fiesta Nacional ó Duelo Nacional - Mixto: Diurna - Nocturna"
                porcentaje = 2.5 * 1.5
            Case "Horas Extra Fiesta Nacional ó Duelo Nacional - Mixto Nocturna - Diurna"
                porcentaje = 2.5 * 1.75
            Case "Horas Extra Fiesta Nacional Diurno con exceso de 3 Horas Diarias ó 9 Semanales"
                porcentaje = 2.5 * 1.25 * 1.75
            Case "Horas Extra Fiesta Nacional Nocturno con exceso de 3 Horas Diarias ó 9 Semanales"
                porcentaje = 2.5 * 1.5 * 1.75
            Case "Horas Extra Fiesta Nacional Mixto: Diurno-Nocturno con exceso de 3 Horas Diarias ó 9 Semanales"
                porcentaje = 2.5 * 1.5 * 1.75
            Case "Horas Extra Fiesta Nacional Mixto: Nocturno-Diurno con exceso de 3 Horas Diarias ó 9 Semanales"
                porcentaje = 2.5 * 1.75 * 1.75
            Case "Horas Extra Domingo ó Descanso Semanal Diurno"
                porcentaje = 1.5 * 1.25
            Case "Horas Extra Domingo ó Descanso Semanal Nocturno"
                porcentaje = 1.5 * 1.5
            Case "Horas Extra Domingo ó Descanso Semanal Mixto: Diurno-Nocturno"
                porcentaje = 1.5 * 1.5
            Case "Horas Extra Domingo ó Descanso Semanal Mixto: Nocturno-Diurno"
                porcentaje = 1.5 * 1.75
            Case "Horas Extra Domingo ó Descanso Semanal Diurno con exceso de 3 Horas Diarias ó 9 Semanales"
                porcentaje = 1.5 * 1.25 * 1.75
            Case "Horas Extra Domingo ó Descanso Semanal Nocturno con exceso de 3 Horas Diarias ó 9 Semanales"
                porcentaje = 1.5 * 1.5 * 1.75
            Case "Horas Extra Domingo ó Descanso Semanal Mixto: Diurno-Nocturno con exceso de 3 Horas Diarias ó 9 Semanales"
                porcentaje = 1.5 * 1.5 * 1.75
            Case "Horas Extra Domingo ó Descanso Semanal Mixto: Nocturno-Diurno con exceso de 3 Horas Diarias ó 9 Semanales"
                porcentaje = 1.5 * 1.75 * 1.75
            Case Else
                porcentaje = 0
        End Select

        Return porcentaje
    End Function

    Public Sub AsignarHorasExtra1(tipo As String, cantidad As Decimal)
        Dim porcentaje As Decimal = ObtenerPorcentajeHoraExtra(tipo)
        thextra1 = tipo
        hextra1 = cantidad
        mhextra1 = (shora * porcentaje) * hextra1
        mhextra1 = Math.Round(mhextra1, 2)
    End Sub

    Public Sub AsignarHorasExtra2(tipo As String, cantidad As Decimal)
        Dim porcentaje As Decimal = ObtenerPorcentajeHoraExtra(tipo)
        thextra2 = tipo
        hextra2 = cantidad
        mhextra2 = (shora * porcentaje) * hextra2
        mhextra2 = Math.Round(mhextra2, 2)
    End Sub

    Public Sub AsignarHorasExtra3(tipo As String, cantidad As Decimal)
        Dim porcentaje As Decimal = ObtenerPorcentajeHoraExtra(tipo)
        thextra3 = tipo
        hextra3 = cantidad
        mhextra3 = (shora * porcentaje) * hextra3
        mhextra3 = Math.Round(mhextra3, 2)
    End Sub


    Public Sub CalcularSalario()
        sbruto = (shora * htrabajadas) + mhextra1 + mhextra2 + mhextra3
        sbruto = Math.Round(sbruto, 2)

        seguro_social = sbruto * 0.0975
        seguro_social = Math.Round(seguro_social, 2)

        seguro_educativo = sbruto * 0.0125
        seguro_educativo = Math.Round(seguro_educativo, 2)

        Dim anual = sbruto * 12

        If anual > 11000 Then
            If anual > 50000 Then
                impuesto_renta = ((anual - 50000) * 0.25) / 12
            Else
                impuesto_renta = ((anual - 11000) * 0.15) / 12
            End If
        Else
            impuesto_renta = 0
        End If
        impuesto_renta = Math.Round(impuesto_renta, 2)

        sneto = sbruto - (seguro_social + seguro_educativo + impuesto_renta + descuento1 + descuento2 + descuento3)
        sneto = Math.Round(sneto)
    End Sub


    Public Sub ImprimirAtributos()
        MsgBox("empleado.Prefijo: " & Me.prefijo &
               vbCrLf & "empleado.Tomo: " & Me.tomo &
               vbCrLf & "empleado.Asiento: " & Me.asiento &
               vbCrLf & "empleado.Cédula: " & Me.cedula &
               vbCrLf & "empleado.Género: " & Me.genero &
               vbCrLf & "empleado.Nombre1: " & Me.nombre1 &
               vbCrLf & "empleado.Nombre2: " & Me.nombre2 &
               vbCrLf & "empleado.Apellido1: " & Me.apellido1 &
               vbCrLf & "empleado.Apellido2: " & Me.apellido2 &
               vbCrLf & "empleado.Estado Civil: " & Me.estado_civil &
               vbCrLf & "empleado.Apellido Casada: " & Me.apellido_casada &
               vbCrLf & "empleado.Usa Apellido Casada: " & Me.usa_apellido_casada &
               vbCrLf & "empleado.S. Hora: " & Me.shora &
               vbCrLf & "empleado.H. Trabajadas: " & Me.htrabajadas &
               vbCrLf & "empleado.H. Extra 1: " & Me.thextra1 & " - " & Me.hextra1 & " - " & Me.mhextra1 &
               vbCrLf & "empleado.H. Extra 2: " & Me.thextra2 & " - " & Me.hextra2 & " - " & Me.mhextra2 &
               vbCrLf & "empleado.H. Extra 3: " & Me.thextra3 & " - " & Me.hextra3 & " - " & Me.mhextra3 &
               vbCrLf & "empleado.Salario Bruto: " & Me.sbruto &
               vbCrLf & "empleado.Seguro Social: " & Me.seguro_social &
               vbCrLf & "empleado.Seguro Educativo: " & Me.seguro_educativo &
               vbCrLf & "empleado.Impuesto Renta: " & Me.impuesto_renta &
               vbCrLf & "empleado.Descuento 1: " & Me.descuento1 &
               vbCrLf & "empleado.Descuento 2: " & Me.descuento2 &
               vbCrLf & "empleado.Descuento 3: " & Me.descuento3 &
               vbCrLf & "empleado.Salario Neto: " & Me.sneto)
    End Sub

    Public Sub RegistrarDatos()
        ' Lógica para registrar los datos en la base de datos
        Dim insertar As String = "INSERT INTO generales (prefijo, tomo, asiento, genero, cedula, nombre1, nombre2, apellido1, apellido2, estado_civil, apellido_casada, usa_apellido_casada, htrabajadas, shora, thextra1, hextra1, mhextra1, thextra2, hextra2, mhextra2, thextra3, hextra3, mhextra3, sbruto, seguro_social, seguro_educativo, impuesto_renta, descuento1, descuento2, descuento3, sneto) " &
    "VALUES (@prefijo, @tomo, @asiento, @genero, @cedula, @nombre1, @nombre2, @apellido1, @apellido2, @estado_civil, @apellido_casada, @usa_apellido_casada, @htrabajadas, @shora, @thextra1, @hextra1, @mhextra1, @thextra2, @hextra2, @mhextra2, @thextra3, @hextra3, @mhextra3, @sbruto, @seguro_social, @seguro_educativo, @impuesto_renta, @descuento1, @descuento2, @descuento3, @sneto)"

        Using conn As New MySqlConnection("server=localhost;userid=d82023;password=ds8;database=d8_modificado")
            Try
                conn.Open()

                Using command As New MySqlCommand(insertar, conn)
                    ' Asignar valores a los parámetros usando la instancia de Empleado
                    command.Parameters.AddWithValue("@prefijo", If(String.IsNullOrEmpty(prefijo), DBNull.Value, prefijo))
                    command.Parameters.AddWithValue("@tomo", If(String.IsNullOrEmpty(tomo), DBNull.Value, tomo))
                    command.Parameters.AddWithValue("@asiento", If(String.IsNullOrEmpty(asiento), DBNull.Value, asiento))
                    command.Parameters.AddWithValue("@genero", If(String.IsNullOrEmpty(genero), DBNull.Value, genero))
                    command.Parameters.AddWithValue("@cedula", If(String.IsNullOrEmpty(cedula), DBNull.Value, cedula))
                    command.Parameters.AddWithValue("@nombre1", If(String.IsNullOrEmpty(nombre1), DBNull.Value, nombre1))
                    command.Parameters.AddWithValue("@nombre2", If(String.IsNullOrEmpty(nombre2), DBNull.Value, nombre2))
                    command.Parameters.AddWithValue("@apellido1", If(String.IsNullOrEmpty(apellido1), DBNull.Value, apellido1))
                    command.Parameters.AddWithValue("@apellido2", If(String.IsNullOrEmpty(apellido2), DBNull.Value, apellido2))
                    command.Parameters.AddWithValue("@estado_civil", If(String.IsNullOrEmpty(estado_civil), DBNull.Value, estado_civil))
                    command.Parameters.AddWithValue("@apellido_casada", If(String.IsNullOrEmpty(apellido_casada), DBNull.Value, apellido_casada))
                    command.Parameters.AddWithValue("@usa_apellido_casada", If(String.IsNullOrEmpty(usa_apellido_casada), DBNull.Value, usa_apellido_casada))
                    command.Parameters.AddWithValue("@htrabajadas", If(htrabajadas = 0, DBNull.Value, htrabajadas))
                    command.Parameters.AddWithValue("@shora", If(shora = 0, DBNull.Value, shora))
                    command.Parameters.AddWithValue("@thextra1", If(String.IsNullOrEmpty(thextra1), DBNull.Value, thextra1))
                    command.Parameters.AddWithValue("@hextra1", If(hextra1 = 0, DBNull.Value, hextra1))
                    command.Parameters.AddWithValue("@mhextra1", If(mhextra1 = 0, DBNull.Value, mhextra1))
                    command.Parameters.AddWithValue("@thextra2", If(String.IsNullOrEmpty(thextra2), DBNull.Value, thextra2))
                    command.Parameters.AddWithValue("@hextra2", If(hextra2 = 0, DBNull.Value, hextra2))
                    command.Parameters.AddWithValue("@mhextra2", If(mhextra2 = 0, DBNull.Value, mhextra2))
                    command.Parameters.AddWithValue("@thextra3", If(String.IsNullOrEmpty(thextra3), DBNull.Value, thextra3))
                    command.Parameters.AddWithValue("@hextra3", If(hextra3 = 0, DBNull.Value, hextra3))
                    command.Parameters.AddWithValue("@mhextra3", If(mhextra3 = 0, DBNull.Value, mhextra3))
                    command.Parameters.AddWithValue("@sbruto", If(sbruto = 0, DBNull.Value, sbruto))
                    command.Parameters.AddWithValue("@seguro_social", If(seguro_social = 0, DBNull.Value, seguro_social))
                    command.Parameters.AddWithValue("@seguro_educativo", If(seguro_educativo = 0, DBNull.Value, seguro_educativo))
                    command.Parameters.AddWithValue("@impuesto_renta", If(impuesto_renta = 0, DBNull.Value, impuesto_renta))
                    command.Parameters.AddWithValue("@descuento1", If(descuento1 = 0, DBNull.Value, descuento1))
                    command.Parameters.AddWithValue("@descuento2", If(descuento2 = 0, DBNull.Value, descuento2))
                    command.Parameters.AddWithValue("@descuento3", If(descuento3 = 0, DBNull.Value, descuento3))
                    command.Parameters.AddWithValue("@sneto", If(sneto = 0, DBNull.Value, sneto))

                    ' Ejecutar la consulta
                    command.ExecuteNonQuery()

                    MessageBox.Show("Registro Grabado", "AVISO")
                End Using
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Using
    End Sub

End Class



