Public Class Empleado
    Public nombre1 As String
    Public nombre2 As String
    Public apellido1 As String
    Public apellido2 As String
    Public apellido_casada As String
    Public usa_ape_casada As Boolean

    Public shora As Decimal
    Public htrabajadas As Decimal
    Public hextras As Decimal
    Public sbruto As Decimal
    Public ssocial As Decimal
    Public seducativo As Decimal
    Public irenta As Decimal
    Public otros_descuentos1 As Decimal
    Public otros_descuentos2 As Decimal
    Public otros_descuentos3 As Decimal
    Public sneto As Decimal

    Public Sub CalcularSalario()
        sbruto = (shora * htrabajadas) + ((shora + (shora * 0.25)) * hextras)
        ssocial = sbruto * 0.0975
        seducativo = sbruto * 0.0125
        Dim anual = sbruto * 12

        If anual > 11000 Then
            If anual > 50000 Then
                irenta = ((anual - 50000) * 0.25) / 12
            Else
                irenta = ((anual - 11000) * 0.15) / 12
            End If
        Else
            irenta = 0
        End If

        sneto = sbruto - (ssocial + seducativo + irenta + otros_descuentos1 + otros_descuentos2 + otros_descuentos3)
    End Sub
End Class



