Module TablaMultiplicar

    Sub Main()
        Dim Num As Integer
        Dim NuevaTabla As Boolean = True

        While NuevaTabla
            Num = Solicitarnumero()
            TablaMultiplicar(Num)
            NuevaTabla = Otra() 'Pregunta al usuario si quiere calcular otra tabla de multiplicar
        End While
    End Sub

    'función para solicitar al usuario el número cuya tabla se quiere calcular
    Function Solicitarnumero()
        'variables
        Dim n As Integer

        Console.Write("Introduce el número de la tabla a calcular: ")

        Try
            n = Integer.Parse(Console.ReadLine())
            Console.Clear() 'limpiar consola
        Catch ex As Exception
            Console.WriteLine("El número introducido no es correcto!!")
            Console.WriteLine(vbCrLf)
            Console.WriteLine("Inténtalo otra vez!!!")
            Console.WriteLine("Introduce el número de la tabla a calcular: ")
            n = Integer.Parse(Console.ReadLine())
        End Try

        Return n

    End Function

    'Procedimiento para calcular las tablas de multiplicar
    Sub TablaMultiplicar(n As Integer)
        'variables
        Dim i As Integer 'el contador
        Dim resultado As Integer

        Console.Clear() 'limpiar consola

        Console.WriteLine("La tabla del {0} es: ", n)
        Console.WriteLine(vbCrLf)

        For i = 1 To 10
            resultado = i * n
            Console.WriteLine("{0} x {1} = {2}", n, i, resultado)
        Next

    End Sub

    'función que pregunta al usuario si quiere calcular otra tabla de multiplicar
    Function Otra()
        'variables
        Dim NuevaTabla As Boolean
        Dim respuesta As String

        Console.WriteLine(vbCrLf)
        Console.WriteLine("¿Quieres calcular otra tabla? (S/N): ")

        respuesta = Console.ReadLine

        While Not NuevaTabla
            'para que tanto 'S' o 's' y 'N' o 'n' sean respuestas válidas
            If respuesta.ToUpper = "S" Then
                NuevaTabla = True

            ElseIf respuesta.ToUpper = "N" Then
                NuevaTabla = False
                Environment.Exit(0)
            Else
                Console.WriteLine("La respuesta no es válida. Ha de ser 'S' o 'N'")
                Console.WriteLine(vbCrLf)
                Console.WriteLine("¿Quieres calcular otra tabla? (S/N): ")
                respuesta = Console.ReadLine()
            End If
        End While

        Return NuevaTabla

    End Function

End Module
