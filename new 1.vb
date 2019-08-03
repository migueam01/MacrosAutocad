Sub UbicaDatosPerfil(ByVal param As String) 'ByVal param As String
    'Dibuja perfiles topográficos y los datos de cortes de Alcantarillado a partir de los datos de Excel
    'Dibuja Hatch entre la base de la tuberia y el nivel del agua
    'Requiere Puntos de referencia, abscisas y cotas.
    'Trabaja con el archivo de Excel tipo: DatosDibujo.xls
    'Definir el número de perfiles a generar
    'Fecha de realización 2015/08/17
    'Realizado por AMC
    'Modificado 11-06-2016
    Dim ACADApli As AcadApplication
    Dim Excel As Object
    Dim ExcelSheet As Object
    Dim ExcelWorkbook As Object
    Dim Celdas As Object
    
    Dim col1 As AcadAcCmColor
    Dim col2 As AcadAcCmColor
    Dim objetoHatch As AcadHatch
    Dim linea As AcadLine
    Dim lineaBaseTub As AcadLine
    Dim lineaAgua As AcadLine
    Dim lineaUnion1 As AcadLine
    Dim lineaUnion2 As AcadLine
    Dim lineaUnionAgua As AcadLine
    
    Dim circulo As AcadCircle
    Dim rectangulo As AcadPolyline
    Dim PtsRect(0 To 11) As Double
    Dim RefExterna As AcadBlockReference
    Dim ObjetoTexto As AcadText
    Dim ObjetoTAbscisa As AcadText
    Dim ObjetoLAbscisa As AcadLine
    
    Dim contorno(0 To 3) As AcadEntity
    Dim Abscisa1(0 To 2) As Double
    Dim Abscisa2(0 To 2) As Double
    Dim centro(0 To 2) As Double
    Dim PAbscisa(0 To 2) As Double
    Dim esquina(0 To 2) As Double
    Dim Ubicalle(0 To 2) As Double
    Dim ptoTextPozo(0 To 2) As Double
    
    Dim ptoEsquina1Pozo(0 To 2) As Double
    Dim ptoEsquina2Pozo(0 To 2) As Double
    
    Dim nfil, numFilas As Integer  'filas para la columna desde M
    
    Dim filasRecorre As Integer
    
    Dim Factor As Single
    Dim altura, radioCirculo As Double
    Dim AbsRef As Double
    Dim TAbscisa(0 To 4) As String
    Dim celdaValorA(0 To 300), celdaValorB(0 To 300), celdaValorG(0 To 300) As String
    Dim PLA1(0 To 2) As Double
    Dim PLA2(0 To 2) As Double
    Dim ptoBaseTubInicio(0 To 2) As Double
    Dim ptoSupTubInicio(0 To 2) As Double
    Dim ptoAguaInicio(0 To 2) As Double
    Dim ptoBaseTubAnterior(0 To 2) As Double
    Dim ptoSupTubAnterior(0 To 2) As Double
    Dim ptoAguaAnterior(0 To 2) As Double
    Dim ptoBaseTub1(0 To 2) As Double 'punto 1 de la base de la tuberia
    Dim ptoBaseTub2(0 To 2) As Double 'punto 2 de la base de la tuberia
    Dim ptoAgua1(0 To 2) As Double    'punto 1 para el nivel de agua
    Dim ptoAgua2(0 To 2) As Double    'punto 2 para el nivel de agua
    Dim ptoAux1(0 To 2) As Double
    Dim ptoAux2(0 To 2) As Double
    Dim DesdeF(0 To 150) As Single
    Dim BajarCota As Double
    Dim Narchivo(0 To 150) As String
    Dim Ncalle(0 To 200) As String
    Dim Filas(0 To 200) As Single
    Dim ArchivoModelo, NPozo As String
    Dim j As Single
    Dim contador As Integer
    Dim NTPerf As Single
    Dim FNArch, filInicio, filFin, filaCopia As Single
    Dim espacioLineasH As Integer
    Dim dc(1 To 4)
    Dim dy(1 To 8)
    Dim dx(1 To 8)
    
    'variables usadas para ubicar los textos
    Dim longTramo, alturaTextPozo, datoPendiente, datoDiametro As Double
    Dim puntoTramos(0 To 2) As Double
    Dim objTextoTramos As AcadText
    Dim texto, longitud, velocidad, diametro, caudal, pendiente, relacion As String
    texto = ""
    
    'variables para el texto de la calle
    Dim nombreCalle As String
    
    'Habilita EXCEL
    controlarErrorExcel ("Resumen")
    Set Excel = GetObject(, "Excel.Application")
    'Excel.Application.Visible = True 'Hace visible
    
    ' Activa archivo de Excel y la hoja de trabajo abierta.
    Set ExcelWorkbook = Excel.ActiveWorkbook
'    Set ExcelSheet = ExcelWorkbook.Worksheets(2)
    Set ExcelSheet = Excel.ActiveSheet
    Set Celdas = Excel.ActiveCell
    FNArch = Celdas.row
    
   'Define datos constantes e iniciales
    Directorio = ExcelSheet.cells(9, 35).Value
    Directorio = Directorio & "\"
    altura = 2.25
    alturaTextPozo = 3.75
    Factor = ExcelSheet.cells(3, 3).Value
    espacioLineasH = ExcelSheet.cells(1, 7).Value
    NCol = 2
        
   'Los siguientes datos pueden variar según el requerimiento
    NTPerf = InputBox("No de Perfiles", "Genera Perfiles", 1)
    FNArch = InputBox("Fila Inicial(Ubicación Nombre Archivo)", "Genera Perfiles", FNArch)
    
    'FNArch: Número de fila donde está el 1er. Nombre de Archivo
    Narchivo(0) = ExcelSheet.cells(FNArch, NCol).Value
    Filas(0) = ExcelSheet.cells(FNArch + 1, NCol + 1).Value
    DesdeF(0) = FNArch + 2
    Ncalle(0) = ExcelSheet.cells(FNArch + 1, NCol).Text
    Ubicalle(0) = 0: Ubicalle(2) = 0#
'    contador = 0
    ArchivoRefExt = ExcelSheet.cells(9, 35).Value
    ArchivoRefExt = ArchivoRefExt & "\DatosCotasGuitarra.dwg"
    
    filasRecorre = FNArch
    ExcelSheet.cells(FNArch, 1).Select
    For j = 0 To NTPerf - 1
        contador = 0
        nfil = FNArch + 2
        filaCopia = FNArch + 2
        'obtener fila donde va a empezar a leer datos
        filInicio = ExcelSheet.cells(filasRecorre + 1, 4).Value
        'obtener fila donde termina la lectura de datos
        filFin = ExcelSheet.cells(filasRecorre + 1, 3).Value
        
        numFilas = filFin - filInicio
        k = 0
        'obtener valor de celdas
        For i = filInicio To filFin
            celdaValorA(k) = ExcelSheet.cells(i, 1).Value
            celdaValorB(k) = ExcelSheet.cells(i, 2).Value
            celdaValorG(k) = ExcelSheet.cells(i, 7).Value
            k = k + 1
        Next i
         
        'colocar valores en las celdas
        k = 0
        For i = filaCopia To (filaCopia + numFilas)
            ExcelSheet.cells(i, 13).Value = celdaValorA(k)
            ExcelSheet.cells(i, 15).Value = celdaValorG(k)
            ExcelSheet.cells(i, 17).Value = celdaValorB(k)
            k = k + 1
        Next i
        
        CotaMin = ExcelSheet.cells(DesdeF(j) - 1, 6).Value
        CotaMax = ExcelSheet.cells(DesdeF(j) - 1, 5).Value
        NLH = Math.Round((CotaMax - CotaMin) / espacioLineasH, 0)
        'valor de desplazamiento de la cota del perfil escala vertical y datos.dwg
        BajarCota = CotaMin * Factor
        Application.Documents.Open (ArchivoRefExt)
        AbsInicio = ExcelSheet.cells(DesdeF(j), NCol).Value
        'obtiene el punto donde se va a ubicar el texto de la calle
        Ubicalle(0) = ((ExcelSheet.cells(filFin, 2).Value - ExcelSheet.cells(filInicio, 2).Value) / 2) - (Len(Ncalle(j)) * 2.9)
        Ubicalle(1) = ExcelSheet.cells(DesdeF(j) - 1, NCol + 3).Value * Factor - BajarCota + 20
        Set ObjetoTexto = ThisDrawing.ModelSpace.AddText(Ncalle(j), Ubicalle, altura * 4.4)
        ObjetoTexto.color = 6
         
        Cont = 0: HpozoA = 0
        For Nfila = DesdeF(j) To Filas(j)
            'Lee del Excel las abscisas y traza el perfil del terreno
            AbsRef = ExcelSheet.cells(Nfila, NCol).Value
            Abscisa1(0) = AbsRef - AbsInicio
            If Nfila < Filas(j) Then
                Abscisa1(1) = ExcelSheet.cells(Nfila, 3).Value * Factor - BajarCota
                Abscisa1(2) = 0#
                Abscisa2(0) = ExcelSheet.cells(Nfila + 1, NCol).Value - AbsInicio
                Abscisa2(1) = ExcelSheet.cells(Nfila + 1, NCol + 1).Value * Factor - BajarCota
                Abscisa2(2) = 0#
                Set linea = ThisDrawing.ModelSpace.AddLine(Abscisa1, Abscisa2)
                '  Linea.Linetype = "TRAZOS"
                linea.color = 24
            End If
           
            UbicaY = -90 '(Ubicación del primer dato, en la ordenada Y )
            esquina(2) = 0#
            esquina(0) = ExcelSheet.cells(Nfila, 2).Value - 0.5 - AbsInicio
            'Ubica Textos de abscisas parciales, Corte y Cota Proyecto de la Salida
            esquina(1) = UbicaY
            texto = ExcelSheet.cells(Nfila, 6).Text
            Set ObjetoTexto = ThisDrawing.ModelSpace.AddText(texto, esquina, altura)
            ObjetoTexto.Rotate esquina, 1.5708
            ObjetoTexto.color = acRed
            'Corte de la Salida del pozo
            esquina(0) = esquina(0) + 3.2
            esquina(1) = UbicaY + 27
            texto = ExcelSheet.cells(nfil, 21).Text
            Set ObjetoTexto = ThisDrawing.ModelSpace.AddText(texto, esquina, altura)
            ObjetoTexto.Rotate esquina, 1.5708
            ObjetoTexto.color = acRed
            'Corte y cota de la llegada al siguiente pozo
            If ExcelSheet.cells(Nfila, 15).Text <> "" Then
                If ExcelSheet.cells(Nfila + 1, 1).Text <> "Archivo:" Then
                    esquina(0) = esquina(0) + ExcelSheet.cells(nfil, 16).Value - 3.2
                    texto = ExcelSheet.cells(nfil, 22).Text
                    Set ObjetoTexto = ThisDrawing.ModelSpace.AddText(texto, esquina, altura)
                    ObjetoTexto.Rotate esquina, 1.5708
                    ObjetoTexto.color = acRed
                    'Cota de la llegada al siguiente pozo
                    esquina(1) = UbicaY + 37
                    texto = ExcelSheet.cells(nfil, 25).Text
                    Set ObjetoTexto = ThisDrawing.ModelSpace.AddText(texto, esquina, altura)
                    ObjetoTexto.Rotate esquina, 1.5708
                    ObjetoTexto.color = acRed
                    esquina(0) = esquina(0) - ExcelSheet.cells(nfil, 16).Value + 3.2
                End If
            End If
            'Cota proyecto de la Salida del pozo
            esquina(1) = UbicaY + 37
            
            'texto = ExcelSheet.Cells(nfila, 19).Text
            texto = ExcelSheet.cells(nfil, 19).Text
            
            Set ObjetoTexto = ThisDrawing.ModelSpace.AddText(texto, esquina, altura)
            ObjetoTexto.Rotate esquina, 1.5708
            ObjetoTexto.color = acRed
            
            'Ubica Textos de abscisas acumuladas, cortes, cotas de proyecto y terreno
            esquina(0) = esquina(0) - 3.2
            dc(1) = 2: dc(4) = 3
            dc(2) = 10: dc(3) = 8
            dy(1) = 11: dy(2) = 27: dy(3) = 37: dy(4) = 51
            For k = 1 To 4
                esquina(1) = UbicaY + dy(k)
                TAbscisa(k) = ExcelSheet.cells(Nfila, dc(k)).Text
                Set ObjetoTexto = ThisDrawing.ModelSpace.AddText(TAbscisa(k), esquina, altura)
                ObjetoTexto.Rotate esquina, 1.5708
                If ExcelSheet.cells(Nfila, 7).Text <> "" Then
                    If k = 2 Or k = 3 Then
                        ObjetoTexto.Erase
                    End If
                End If
            Next k
         
           'Ubica líneas verticales de abscisa
            PLA1(0) = Abscisa1(0)
            
            If ExcelSheet.cells(Nfila, 7).Text = "" Then
                PLA1(1) = -23.5
            Else
                PLA1(1) = Abscisa1(1) + 12
            End If
            
            PLA1(2) = 0#
            PLA2(0) = Abscisa1(0)
            PLA2(1) = -90.7
            PLA2(2) = 0#
            If Nfila = Filas(j) Then
                If ExcelSheet.cells(Nfila + 1, 1).Text <> "Archivo:" Then
                    PLA1(1) = 0
                Else
                    PLA1(1) = Abscisa2(1) + 12
                End If
            End If
            
            Set ObjetoLAbscisa = ThisDrawing.ModelSpace.AddLine(PLA1, PLA2)
              
            If ExcelSheet.cells(Nfila, 7).Text <> "" Then
                ObjetoLAbscisa.color = acRed
                'dibuja los círculos de los pozos y el nombre del pozo
                centro(0) = PLA1(0): centro(1) = PLA1(1) + 12: centro(2) = 0
                Set circulo = ThisDrawing.ModelSpace.AddCircle(centro, 10)
                circulo.color = acRed
                NPozo = ExcelSheet.cells(Nfila, 1).Text
                radioCirculo = circulo.Diameter / 2
                'se obtiene el punto para ubicar el nombre del pozo
                ptoTextPozo(0) = centro(0) - (radioCirculo / 2) - Len(NPozo) / 2
                ptoTextPozo(1) = centro(1) - 1.96
                Set ObjetoTexto = ThisDrawing.ModelSpace.AddText(NPozo, ptoTextPozo, alturaTextPozo)
                ObjetoTexto.color = acMagenta
                ObjetoTexto.StyleName = "SANTY"
                abscisaPozo = ExcelSheet.cells(Nfila, 2).Text - AbsInicio
                 'verifica si es el último pozo
                If ExcelSheet.cells(Nfila + 1, 1).Text <> "Archivo:" Then
                    'ubica textos entre tramos
                    longTramo = ExcelSheet.cells(nfil, 16).Text
                    puntoTramos(2) = 0
                    longitud = "L= " & ExcelSheet.cells(nfil, 16).Text & "m"
                    velocidad = "v= " & ExcelSheet.cells(nfil, 31).Text & "m/s"
                    datoDiametro = Round(ExcelSheet.cells(nfil, 30).Text, 0)
                    tipoTub = ExcelSheet.cells(nfil, 27).Text
                    If tipoTub = "CANAL" Then   'cambio antes Left(TipoTub, 5) = "HS"    'HA en lugar de canal
                        diametro = "B=" & Format(datoDiametro / 1000, "0.00") & "m"   'cambio antes BxH=
                        'caudal = "q=" & Round(ExcelSheet.cells(nfil, 32).Text / 1000, 2) & "m3/s"
                        caudal = "q=" & ExcelSheet.cells(nfil, 32).Text & " l/s"   'Para canales pequeños (Ej: Curgua-Guaranda)
                    Else
                        diametro = "%%c" & "=" & datoDiametro & "mm"
                        caudal = "q=" & ExcelSheet.cells(nfil, 32).Text & " l/s"
                    End If
                    datoPendiente = Round((ExcelSheet.cells(nfil, 29).Text / 10), 1)
                    'pendiente = "j= " & ExcelSheet.Cells(nfil, 29).Text / 10 & "%"
                    pendiente = "j= " & datoPendiente & "%"
                    relacion = "y/D= " & ExcelSheet.cells(nfil, 33).Text
                   
                    If longTramo < 30 Then
                        puntoTramos(0) = abscisaPozo + (longTramo / 2) - 6
                        puntoTramos(1) = -3
                        Set objTextoTramos = ThisDrawing.ModelSpace.AddText(longitud, puntoTramos, altura)
                        objTextoTramos.StyleName = "Santy"
                        puntoTramos(1) = puntoTramos(1) - 3.2
                        Set objTextoTramos = ThisDrawing.ModelSpace.AddText(velocidad, puntoTramos, altura)
                        objTextoTramos.StyleName = "Santy"
                        puntoTramos(1) = puntoTramos(1) - 3.2
                        Set objTextoTramos = ThisDrawing.ModelSpace.AddText(diametro, puntoTramos, altura)
                        objTextoTramos.StyleName = "Santy"
                        puntoTramos(1) = puntoTramos(1) - 3.2
                        Set objTextoTramos = ThisDrawing.ModelSpace.AddText(caudal, puntoTramos, altura)
                        objTextoTramos.StyleName = "Santy"
                        puntoTramos(1) = puntoTramos(1) - 3.2
                        Set objTextoTramos = ThisDrawing.ModelSpace.AddText(pendiente, puntoTramos, altura)
                        objTextoTramos.StyleName = "Santy"
                        puntoTramos(1) = puntoTramos(1) - 3.2
                        Set objTextoTramos = ThisDrawing.ModelSpace.AddText(relacion, puntoTramos, altura)
                        objTextoTramos.StyleName = "Santy"
                        puntoTramos(1) = puntoTramos(1) - 3.5
                        Set objTextoTramos = ThisDrawing.ModelSpace.AddText(tipoTub, puntoTramos, altura)
                        objTextoTramos.StyleName = "Santy"
                    Else
                        puntoTramos(0) = abscisaPozo + (longTramo / 2) - 16
                        puntoTramos(1) = -5.75
                        texto = longitud & "; " & velocidad
                        Set objTextoTramos = ThisDrawing.ModelSpace.AddText(texto, puntoTramos, altura)
                        puntoTramos(1) = -9.75
                        texto = diametro & "; " & caudal
                        Set objTextoTramos = ThisDrawing.ModelSpace.AddText(texto, puntoTramos, altura)
                        puntoTramos(1) = -13.75
                        texto = pendiente & "; " & relacion
                        Set objTextoTramos = ThisDrawing.ModelSpace.AddText(texto, puntoTramos, altura)
                        puntoTramos(1) = puntoTramos(1) - 3.5
                        Set objTextoTramos = ThisDrawing.ModelSpace.AddText(tipoTub, puntoTramos, altura)
                        objTextoTramos.StyleName = "Santy"
                    End If
                End If
                 
               'dibuja los rectángulos de los pozos
                If ExcelSheet.cells(Nfila + 1, 1).Text <> "Archivo:" Then
                    columna1 = 20: columna2 = 21
                Else
                    columna1 = 22: columna2 = 23
                End If
                HpozoE = ExcelSheet.cells(nfil, columna1).Value * Factor
                HpozoS = ExcelSheet.cells(nfil, columna2).Value * Factor
               
                If (HpozoE - HpozoS) <= 0 Then
                    HMpozo = HpozoS
                Else
                    HMpozo = HpozoE
                End If
                
                If HpozoA >= HMpozo Then
                    HMpozo = HpozoA
                End If
                HpozoA = ExcelSheet.cells(nfil, 22).Value * Factor
                If ExcelSheet.cells(Nfila + 1, 1).Text <> "Archivo:" Then
                    PtsRect(1) = Abscisa1(1)
                Else
                    PtsRect(1) = Abscisa2(1)
                End If
                PtsRect(0) = PLA1(0) - 1: PtsRect(1) = PtsRect(1) - HMpozo: PtsRect(2) = 0
                PtsRect(3) = PLA1(0) + 1: PtsRect(4) = PtsRect(1): PtsRect(5) = 0
                PtsRect(6) = PtsRect(3): PtsRect(7) = PtsRect(1) + HMpozo: PtsRect(8) = 0
                PtsRect(9) = PtsRect(0): PtsRect(10) = PtsRect(7): PtsRect(11) = 0
                Set rectangulo = ThisDrawing.ModelSpace.AddPolyline(PtsRect)
                rectangulo.Closed = True
                rectangulo.color = acRed
                 
                'verifica si es el último pozo
'                contador = 0
                If ExcelSheet.cells(Nfila + 1, 1).Text <> "Archivo:" Then
                    'dibuja la tubería
                    PLA1(0) = PtsRect(3): PLA1(1) = (ExcelSheet.cells(nfil, 3).Value - ExcelSheet.cells(nfil, 21).Value - CotaMin) * Factor
                    PLA1(2) = 0
                    
                    PLA2(0) = PtsRect(3) + ExcelSheet.cells(nfil, 16).Value - 2
                    PLA2(1) = (ExcelSheet.cells(nfil, 24).Value - ExcelSheet.cells(nfil, 22).Value - CotaMin) * Factor
                    PLA2(2) = 0
                    ptoBaseTub1(0) = PLA1(0)
                    ptoBaseTub1(1) = PLA1(1)
                    ptoBaseTub2(0) = PLA2(0)
                    ptoBaseTub2(1) = PLA2(1)
'                    ptoBaseTubInicio(0) = ptoBaseTub1(0) 'Punto inicial de la tuberia
'                    ptoBaseTubInicio(1) = ptoBaseTub1(1)
'                    Set ObjetoLAbscisa = ThisDrawing.ModelSpace.AddLine(ptoBaseTub1, ptoBaseTub2)
                    Set lineaBaseTub = ThisDrawing.ModelSpace.AddLine(ptoBaseTub1, ptoBaseTub2)
                    
                    lineaBaseTub.color = 134
                    diametro = ExcelSheet.cells(nfil, 30).Value / 1000 * Factor
                    PLA1(1) = PLA1(1) + diametro
                    PLA2(1) = PLA2(1) + diametro
                    Set ObjetoLAbscisa = ThisDrawing.ModelSpace.AddLine(PLA1, PLA2)
                    ObjetoLAbscisa.color = 134
                    y = ExcelSheet.cells(nfil, 33).Value * diametro
                    ptoAgua1(0) = ptoBaseTub1(0)
                    ptoAgua1(1) = ptoBaseTub1(1) + y
                    ptoAgua2(0) = ptoBaseTub2(0)
                    ptoAgua2(1) = ptoBaseTub2(1) + y
                    
'                    ptoAguaInicio(0) = ptoAgua1(0) 'Punto inicial del nivel de agua
'                    ptoAguaInicio(1) = ptoAgua1(1)
                    
'                    PLA1(1) = PLA1(1) - y
'                    PLA2(1) = PLA2(1) - y
'                    Set ObjetoLAbscisa = ThisDrawing.ModelSpace.AddLine(ptoAgua1, ptoAgua2)
                    Set lineaAgua = ThisDrawing.ModelSpace.AddLine(ptoAgua1, ptoAgua2)
                    lineaAgua.color = 140
                    
                    'dibujar hatch entre la base de la tuberia y el nivel del agua
                    'une con lineas la base de la tuberia con el nivel del agua
                    Set lineaUnion1 = ThisDrawing.ModelSpace.AddLine(ptoBaseTub1, ptoAgua1)
                    lineaUnion1.color = acRed
                    Set lineaUnion2 = ThisDrawing.ModelSpace.AddLine(ptoBaseTub2, ptoAgua2)
                    lineaUnion2.color = acRed
                    
                    Set objetoHatch = ThisDrawing.ModelSpace.AddHatch(acPreDefinedGradient, "INVHEMISPHERICAL", True, acGradientObject)
                    Set col1 = AcadApplication.GetInterfaceObject("AutoCAD.AcCmColor.19")  '.19 para Autocad 2014
                    Set col2 = AcadApplication.GetInterfaceObject("AutoCAD.AcCmColor.19")
                    Call col1.SetRGB(255, 191, 0)
                    Call col2.SetRGB(0, 191, 255)
                    objetoHatch.GradientColor1 = col1
                    objetoHatch.GradientColor2 = col2
                    objetoHatch.EntityTransparency = 50
                    
'                    Set contorno(0) = ThisDrawing.ModelSpace.AddLine(ptoBaseTub1, ptoBaseTub2)
'                    Set contorno(1) = ThisDrawing.ModelSpace.AddLine(ptoAgua1, ptoAgua2)
'                    Set contorno(2) = ThisDrawing.ModelSpace.AddLine(ptoBaseTub1, ptoAgua1)
'                    Set contorno(3) = ThisDrawing.ModelSpace.AddLine(ptoBaseTub2, ptoAgua2)
                    Set contorno(0) = lineaBaseTub
                    Set contorno(1) = lineaAgua
                    Set contorno(2) = lineaUnion1
                    Set contorno(3) = lineaUnion2
                    objetoHatch.AppendOuterLoop (contorno)
                    objetoHatch.Evaluate
                    lineaUnion1.Delete
                    lineaUnion2.Delete
                    
                    If contador >= 1 Then
                        Set lineaUnionAgua = ThisDrawing.ModelSpace.AddLine(ptoAguaAnterior, ptoAgua1)
                        lineaUnionAgua.color = 140
                        ptoEsquina1Pozo(0) = PtsRect(0)
                        ptoEsquina1Pozo(1) = PtsRect(1)
                        ptoEsquina1Pozo(2) = PtsRect(2)
                        ptoEsquina2Pozo(0) = PtsRect(3)
                        ptoEsquina2Pozo(1) = PtsRect(4)
                        ptoEsquina2Pozo(2) = PtsRect(5)
                        Set lineaBaseTub = ThisDrawing.ModelSpace.AddLine(ptoEsquina1Pozo, ptoEsquina2Pozo)
                        Set lineaUnion1 = ThisDrawing.ModelSpace.AddLine(ptoAguaAnterior, ptoEsquina1Pozo)
                        Set lineaUnion2 = ThisDrawing.ModelSpace.AddLine(ptoAgua1, ptoEsquina2Pozo)
                        Set contorno(0) = lineaUnionAgua
                        Set contorno(1) = lineaBaseTub
                        Set contorno(2) = lineaUnion1
                        Set contorno(3) = lineaUnion2
                        Set objetoHatch = ThisDrawing.ModelSpace.AddHatch(acPreDefinedGradient, "INVHEMISPHERICAL", True, acGradientObject)
                        objetoHatch.GradientColor1 = col1
                        objetoHatch.GradientColor2 = col2
                        objetoHatch.EntityTransparency = 50
                        objetoHatch.AppendOuterLoop (contorno)
                        objetoHatch.Evaluate
                        lineaBaseTub.Delete
                        lineaUnion1.Delete
                        lineaUnion2.Delete
                    Else
                        ptoBaseTubInicio(0) = ptoBaseTub1(0) 'Punto inicial de la tuberia
                        ptoBaseTubInicio(1) = ptoBaseTub1(1)
                        
                        ptoAguaInicio(0) = ptoAgua1(0) 'Punto inicial del nivel de agua
                        ptoAguaInicio(1) = ptoAgua1(1)
                    End If
                    
                    contador = contador + 1
                    ptoBaseTubAnterior(0) = ptoBaseTub2(0)
                    ptoBaseTubAnterior(1) = ptoBaseTub2(1)
                    ptoAguaAnterior(0) = ptoAgua2(0)
                    ptoAguaAnterior(1) = ptoAgua2(1)
                 End If
                 
                'obtiene y dibuja el texto de la calle que cruza
                nombreCalle = ExcelSheet.cells(nfil, 28).Text
                puntoTramos(0) = abscisaPozo + 3 + 1.3: puntoTramos(2) = 0
                puntoTramos(1) = Abscisa1(1) - HMpozo - 24
                Set objTextoTramos = ThisDrawing.ModelSpace.AddText(nombreCalle, puntoTramos, altura)
                objTextoTramos.Rotate puntoTramos, 1.5708
                End If
                nfil = nfil + 1
            Next Nfila
            
            Set objetoHatch = ThisDrawing.ModelSpace.AddHatch(acPreDefinedGradient, "INVHEMISPHERICAL", True, acGradientObject)
            objetoHatch.GradientColor1 = col1
            objetoHatch.GradientColor2 = col2
            objetoHatch.EntityTransparency = 50
'
'             'Dibuja las lineas y el Hatch en el interior del pozo inicio
             ptoAux1(0) = ptoBaseTubInicio(0) - 2
             ptoAux1(1) = ptoBaseTubInicio(1)
             
             ptoBaseTubInicio(0) = PtsRect(0)
             ptoBaseTubInicio(1) = PtsRect(1)
             ptoBaseTubInicio(2) = PtsRect(2)
             Set lineaBaseTub = ThisDrawing.ModelSpace.AddLine(ptoBaseTubInicio, ptoAux1)
             
             ptoAux2(0) = ptoAguaInicio(0) - 2
             ptoAux2(1) = ptoAguaInicio(1)
             ptoAguaInicio(0) = PtsRect(3)
             ptoAguaInicio(1) = PtsRect(4)
             ptoAguaInicio(2) = PtsRect(5)
             Set lineaAgua = ThisDrawing.ModelSpace.AddLine(ptoAguaInicio, ptoAux2)
             lineaAgua.color = 140
             Set lineaUnion1 = ThisDrawing.ModelSpace.AddLine(ptoAux1, ptoAux2)
             Set lineaUnion2 = ThisDrawing.ModelSpace.AddLine(ptoBaseTubInicio, ptoAguaInicio)
'
             Set contorno(0) = lineaBaseTub
             Set contorno(1) = lineaAgua
             Set contorno(2) = lineaUnion1
             Set contorno(3) = lineaUnion2
             objetoHatch.AppendOuterLoop (contorno)
             objetoHatch.Evaluate
             lineaUnion1.Delete
             lineaUnion2.Delete
             lineaBaseTub.Delete
'
            Set objetoHatch = ThisDrawing.ModelSpace.AddHatch(acPreDefinedGradient, "INVHEMISPHERICAL", True, acGradientObject)
            objetoHatch.GradientColor1 = col1
            objetoHatch.GradientColor2 = col2
            objetoHatch.EntityTransparency = 50
             'Dibuja las lineas y el Hatch en el interior del pozo final
             ptoAux1(0) = ptoBaseTub2(0) + 2
             ptoAux1(1) = ptoBaseTub2(1)
             Set lineaBaseTub = ThisDrawing.ModelSpace.AddLine(ptoBaseTub2, ptoAux1)

             ptoAux2(0) = ptoAgua2(0) + 2
             ptoAux2(1) = ptoAgua2(1)
             Set lineaAgua = ThisDrawing.ModelSpace.AddLine(ptoAgua2, ptoAux2)
             lineaAgua.color = 140
             Set lineaUnion1 = ThisDrawing.ModelSpace.AddLine(ptoAux1, ptoAux2)
             Set lineaUnion2 = ThisDrawing.ModelSpace.AddLine(ptoBaseTub2, ptoAgua2)

             Set contorno(0) = lineaBaseTub
             Set contorno(1) = lineaAgua
             Set contorno(2) = lineaUnion1
             Set contorno(3) = lineaUnion2
             objetoHatch.AppendOuterLoop (contorno)
             objetoHatch.Evaluate
             lineaUnion1.Delete
             lineaUnion2.Delete
             lineaBaseTub.Delete
            
            Celdas.Offset(Filas(j) - DesdeF(j) + 3, 0).Select
            ObjetoLAbscisa.color = acGreen
            ZoomAll
            'Ubica las dos lineas verticales: izquierda (eje) y derecha (extremo perfil)
            PLA1(0) = ExcelSheet.cells(Filas(j), 2).Value - AbsInicio
            PLA1(1) = 0#
            PLA1(2) = 0#
            PLA2(0) = PLA1(0)
            PLA2(1) = (CotaMax - CotaMin) * Factor
            PLA2(2) = 0#
            For k = 1 To 2
                Set ObjetoLAbscisa = ThisDrawing.ModelSpace.AddLine(PLA1, PLA2)
                ObjetoLAbscisa.color = 3
                PLA1(0) = 0#
                PLA2(0) = 0#
            Next k
         
            'Ubica lineas horizontales y cotas referenciales
            esquina(0) = -25 'cambio realizado -25 en lugar de -30
            PLA1(0) = -5#
            PLA1(1) = 0
            PLA1(2) = 0#
            PLA2(0) = ExcelSheet.cells(Filas(j), 2).Value - AbsInicio
            PLA2(1) = PLA1(1)
            PLA2(2) = 0#
            esquina(1) = 0
            texto = Format(CotaMin, "00.00")
            altura2 = Factor / 2.75
            Set ObjetoTexto = ThisDrawing.ModelSpace.AddText(texto, esquina, 3#) 'cambio realizado altura2 en lugar de 3
            ObjetoTexto.color = 40
            
            For lineaH = 1 To NLH
                PLA1(1) = PLA1(1) + espacioLineasH * Factor
                PLA2(1) = PLA1(1)
                esquina(0) = -20
                esquina(1) = esquina(1) + espacioLineasH * Factor
                texto = Format(texto + espacioLineasH, "00.00")
                Set ObjetoLAbscisa = ThisDrawing.ModelSpace.AddLine(PLA1, PLA2)
                If lineaH = NLH Then
                    ObjetoLAbscisa.color = 3
                 Else
                    ObjetoLAbscisa.color = 9
                End If
                esquina(0) = -25
                Set ObjetoTexto = ThisDrawing.ModelSpace.AddText(texto, esquina, 2.5)   'cambio realizado altura2 en lugar de 3
                ObjetoTexto.color = 40
            Next lineaH
            ZoomAll
            'Traza la guitarra
            dx(1) = -15: dx(2) = ExcelSheet.cells(Filas(j), 2).Value - AbsInicio:
            dx(3) = dx(1): dx(4) = dx(2): dx(5) = dx(1): dx(6) = dx(2): dx(7) = dx(1): dx(8) = dx(2)
            dy(1) = 0: dy(2) = -23.5: dy(3) = -40.5: dy(4) = -54.5: dy(5) = -64.5: dy(6) = -79.7: dy(7) = -90.7
            For k = 1 To 7
                PLA1(0) = dx(k): PLA1(1) = dy(k)
                PLA2(0) = dx(k + 1): PLA2(1) = dy(k)
                Set ObjetoLAbscisa = ThisDrawing.ModelSpace.AddLine(PLA1, PLA2)
                If (k = 1 Or k = 7) Then
                    ObjetoLAbscisa.color = 3
                Else
                    ObjetoLAbscisa.color = 1
                End If
            Next k
            PLA1(0) = dx(1): PLA1(1) = dy(1)
            PLA2(0) = dx(1): PLA2(1) = dy(7)
            Set ObjetoLAbscisa = ThisDrawing.ModelSpace.AddLine(PLA1, PLA2)
            ObjetoLAbscisa.color = 3
        
            'Visualiza el dibujo y graba el archivo de dibujo
            ZoomAll
            ThisDrawing.SaveAs (Directorio & Narchivo(j))
            Narchivo(j + 1) = ExcelSheet.cells(Filas(j) + 1, 2).Value
            DesdeF(j + 1) = Filas(j) + 3
            Filas(j + 1) = ExcelSheet.cells(Filas(j) + 2, 3).Value
            Ncalle(j + 1) = ExcelSheet.cells(Filas(j) + 2, NCol).Text
            If j < NTPerf - 1 Then
                ThisDrawing.Close
            End If
            Set Celdas = Excel.ActiveCell
            filasRecorre = Nfila
            FNArch = Nfila   'aumentado
            'limpiar celdas copiadas
            '      k = 0
            '      For i = filaCopia To (filaCopia + numFilas)
            '        ExcelSheet.Cells(i, 13).Value = ""
            '        ExcelSheet.Cells(i, 15).Value = ""
            '        ExcelSheet.Cells(i, 17).Value = ""
            '        k = k + 1
            '      Next i
            'ExcelSheet.cells(i, 13).Value = ""
    Next j
End Sub
