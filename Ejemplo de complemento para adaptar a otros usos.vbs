' áéíóú
' Ejemplo de complemento para adaptar a otros usos
Descripcion = "Convierte resúmenes de partidas a minúsculas respetando la inicial y palabras con números"
	Set Obra = GetObject ("", "Presto.App.18") ' Obligatorio para acceder a una obra, en este caso, la obra activa
	Respuesta = MsgBox (Descripcion, vbOKCancel )

If Respuesta <> vbCancel Then

	Set Obra = GetObject ("", "Presto.App.18") ' Obligatorio para acceder a una obra, en este caso, la obra activa
	PrimeraVez = True  ' Para lanzar más adelante el inicio de Deshacer
' Lee la tabla "Conceptos", ordenada por el campo clave "Código", con una máscara y un filtro
	Obra.SetElement 1, "Conceptos","Conceptos.Código" , chr(34)&"*"&chr(34), "Conceptos.Nat == 5"

	While Obra.GetElement(1) = 0	' Lee conceptos que cumplan la máscara y el filtro
		Codigo = Obra.GetField ("Conceptos.Código")
		Resumen = Obra.GetField ("Conceptos.Resumen")
		Pres = Obra.GetField ("Conceptos.Pres")
		Obra.LogMsg Codigo & vbTab& Resumen & vbTab& Pres, 1	' Comprobaciones previas
		' Obra.LogMsg Resumen , 1

		ResumenNuevo = Replace (Resumen, ".", ". ")		' Expande el texto para localizar más abreviaturas
		ResumenNuevo = ConvertirATextoMinusculo(ResumenNuevo)
		ResumenNuevo = UCase (Left(ResumenNuevo, 1)) & Mid (ResumenNuevo, 2, Len(ResumenNuevo))	' Convierte la inicial
		Obra.LogMsg ResumenNuevo, 0
		If PrimeraVez = True Then	' Lanza el incio del deshacer justo cuando es necesario
			Obra.SetModal 0
			Obra.BeginRedo
			PrimeraVez = False
		End If
		Obra.SetField "Conceptos.Resumen", ResumenNuevo
		Obra.UpdateRecord ("Conceptos")
	Wend

	If PrimeraVez = False Then	' Lanza el fin del deshacer si se lanzó el inicio
		Obra.EndRedo
		Obra.SetModal 1
	End If
End If


' Proviene directamente de OpenAI 3.5
' Función para convertir una cadena a minúsculas, excepto las palabras con números
Function ConvertirATextoMinusculo(cadena)
    Dim palabras
    palabras = Split(cadena, " ")
    For i = 0 To UBound(palabras)
        If Not EsPalabraConNumeros(palabras(i)) Then
            palabras(i) = LCase(palabras(i))
        End If
    Next
    ConvertirATextoMinusculo = Join(palabras, " ")
End Function

' Función para verificar si una palabra contiene números
Function EsPalabraConNumeros(palabra)
    Dim i
    EsPalabraConNumeros = False
    For i = 1 To Len(palabra)
        If IsNumeric(Mid(palabra, i, 1)) Then
            EsPalabraConNumeros = True
            Exit Function
        End If
    Next
End Function
