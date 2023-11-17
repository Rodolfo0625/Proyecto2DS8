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
    Public salarioBase As Decimal

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
End Class



