    Shared Function LeerPlanoYCargar(ByVal ruta As String, ByVal tipoCargue As TipoCarguePlano, ByVal usucargue As String, Optional ByVal idTransmisionDispersion As String = "0") As DataTable
        
        Dim dtEsquema As New DataTable
        

        Dim dtTemp As New DataTable

        
        Dim objARCHIVO As New TM56_PLANO
        Dim objLOG As New TM55_PLANO_LOG

        Dim objCARGAGIRO As New TM52_PLANO_CARGAGIROS
        
        Dim objESTADODISPERSION As New TM53_PLANO_ESTADODISPERSION
        

        Dim SqlInsert As String = Nothing

        Dim con As New SqlClient.SqlConnection
        con.ConnectionString = Conexion

        Dim contador As Integer = 1
        Dim Linea As String = ""
        Dim sr As StreamReader = New StreamReader(ruta)
        Dim srUpdate As StreamReader = New StreamReader(ruta)
        Dim ResponsableIdentificacion As String = UsuarioRed.ToString()
        Dim ResponsableId As String = Nothing

        Try
            dtEsquema = CargarDatosEsquemaSP()

            objLOG.TM55_TipoCargue = tipoCargue
            objLOG.TM55_FechaCargue = Now
            objLOG.TM55_Usuario = usucargue
            objLOG.TM55_Descripcion = "INICIO DEL PROCESAMIENTO DEL PLANO: " & ruta & ""

            Dim dataTable As New DataTable
            Dim ds As New DataSet()
            Dim totalReg As Integer = 0

            Dim consultaResponsable As String = "SELECT Id FROM Responsables WHERE Identificacion= '" + ResponsableIdentificacion + "'"
            Dim adapResponsable As New SqlClient.SqlDataAdapter(consultaResponsable, con.ConnectionString.ToString)
            adapResponsable.Fill(ds, "Responsables")
            dataTable = ds.Tables("Responsables")
            totalReg = dataTable.Rows.Count
            If totalReg > 0 Then
                ResponsableId = dataTable.Rows(0)(0).ToString
            End If
            con.Close()

            Dim horaInicial As DateTime = Format(DateTime.Now, "yyyy-MM-dd HH:mm:ss")
            Dim horaFinal As DateTime = Format(DateTime.Now, "yyyy-MM-dd HH:mm:ss")

            Select Case tipoCargue
            
                Case TipoCarguePlano.ESTADODISPERSION
                   
                    Dim TIPOREG1 As String = Nothing
                    Dim FECHAGENERACION As String = Nothing
                    Dim HORAGENERACION As String = Nothing
                    Dim NOIDCLIENTE As String = Nothing
                    Dim TIPOCUENTACLIENTE As String = Nothing
                    Dim NOCUENTACLIENTE As String = Nothing
                    Dim FECHAABONO As String = Nothing
                    Dim VALORTOTALTRANSFER As String = Nothing
                    Dim NOREGISTROS As String = Nothing
                    Dim NOIDTRANSMISION As String = Nothing
                    Dim CEROS As String = Nothing

                    ticket = idTransmisionDispersion       
                    objLOG.TM55_NoCargue = ticket
                    InsertarLOG(objLOG)

                    Dim NOCARGUE As Integer = ticket
                    Dim IDSECTOR As Integer = SDS_FrmPrincipal.entidadLiquidacion.idSector
                    Dim IDENTIDAD As Integer = SDS_FrmPrincipal.entidadLiquidacion.idEntidad

                    Dim factura As String = Nothing
                    Dim estado As String = Nothing
                    Dim estadoDescripcion As String = Nothing

                    Using conectado As New SqlClient.SqlConnection(con.ConnectionString.ToString)

                        While Not sr.EndOfStream
                            Linea = sr.ReadLine()

                            If (Not Linea Is "") Then 
                                If (contador = 1) Then

                                    TIPOREG1 = Mid(Linea, 1, 1)
                                    FECHAGENERACION = Mid(Linea, 2, 8)
                                    HORAGENERACION = Mid(Linea, 10, 6)
                                    NOIDCLIENTE = Mid(Linea, 16, 11)
                                    TIPOCUENTACLIENTE = Mid(Linea, 27, 2)
                                    NOCUENTACLIENTE = Mid(Linea, 29, 11)
                                    FECHAABONO = Mid(Linea, 40, 8)
                                    VALORTOTALTRANSFER = Mid(Linea, 48, 18)
                                    NOREGISTROS = "000000"
                                    NOIDTRANSMISION = Mid(Linea, 66, 20)
                                    CEROS = Mid(Linea, 86, 34)

                                ElseIf Not (contador = 1) Then
                                    factura = Mid(Linea, 130, 10)
                                    estado = Mid(Linea, 140, 3)
                                    estadoDescripcion = Mid(Linea, 143, 250)
                                    SqlInsert = "INSERT INTO TM53_PLANO_ESTADODISPERSION (TIPOREG1, FECHAGENERACION, HORAGENERACION, NOIDCLIENTE, TIPOCUENTACLIENTE, NOCUENTACLIENTE, FECHAABONO, VALORTOTALTRANSFER, NOREGISTROS, NOIDTRANSMISION, CEROS, FECHA_CARGUE, NOCARGUE, IDSECTOR, IDENTIDAD, TIPOREG2, NOMBENEF, TIPOIDENTIFICACIONBENEF, NOIDENTIFICACIONBENEF, TIPOCUENTABENEF, NOCUENTABENEF, CODCOMPENSACIONBANCO, DESCBANCO, VALORTRANSFER, CODCIUDCUENTA, FACTURA, ESTADO, DESCESTADO, FECHADEV, PROCESADO) VALUES ('" + TIPOREG1 + "', '" + FECHAGENERACION + "', '" + HORAGENERACION + "', '" + NOIDCLIENTE + "', '" + TIPOCUENTACLIENTE + "', '" + NOCUENTACLIENTE + "', '" + FECHAABONO + "', '" + VALORTOTALTRANSFER + "', '" + NOREGISTROS + "', '" + NOIDTRANSMISION + "', '" + CEROS + "', '" + Format(Now, "yyyy-MM-dd HH:mm:ss.fff") + "', " + NOCARGUE.ToString + ", " + IDSECTOR.ToString + ", " + IDENTIDAD.ToString + ", '" + Mid(Linea, 1, 1) + "', '" + Mid(Linea, 2, 30) + "', '" + Mid(Linea, 32, 2) + "', '" + Mid(Linea, 34, 11) + "', '" + Mid(Linea, 45, 2) + "', '" + Mid(Linea, 47, 17) + "', '" + Mid(Linea, 64, 4) + "', '" + Mid(Linea, 68, 20) + "', '" + Mid(Linea, 88, 18) + "', '" + Mid(Linea, 106, 24) + "', '" + factura.ToString + "', '" + estado.ToString + "', '" + estadoDescripcion.ToString + "', '0000-00-00', '0')"

                                    Dim Ocomando As New SqlCommand(SqlInsert, conectado)
                                    Ocomando.Connection.Open()
                                    salvar = Ocomando.ExecuteNonQuery()
                                    Ocomando.Connection.Close()

                                    If Not (estado = "S00" Or estado = Nothing) Then
                                            
                                        Dim SqlInsertHistorial As String = "INSERT INTO TM70_Historial (Dato, IdCampo, ValorAnterior, ValorNuevo, IdEntidad, IdSector, IdUsuario, Fecha, Observacion, IdResponsable) VALUES ('" +
                                            factura.TrimStart("0").ToString + "', '" +
                                            estado.ToString + "', ' ', '" +
                                            "(" + estado.ToString + ") " + estadoDescripcion.ToString + "', '" +
                                            SDS_FrmPrincipal.entidadLiquidacion.idEntidad.ToString + "', '" +
                                            SDS_FrmPrincipal.entidadLiquidacion.idSector.ToString + "', '" +
                                            ResponsableIdentificacion.ToString + "', '" +
                                            Format(Now, "yyyy-MM-dd HH:mm:ss.fff").ToString + "', 'Pago rechazado, " +
                                            Format(Now, "yyyy-MM-dd").ToString + ".', '" + ResponsableId.ToString + "')"

                                        Dim OcomandoH As New SqlCommand(SqlInsertHistorial, conectado)
                                        OcomandoH.Connection.Open()
                                        salvar = OcomandoH.ExecuteNonQuery()
                                        OcomandoH.Connection.Close()
                                    ElseIf (estado = "S00" And Not (estado = Nothing)) Then
                                        Dim SqlInsertHistorial As String = "INSERT INTO TM70_Historial (Dato, IdCampo, ValorAnterior, ValorNuevo, IdEntidad, IdSector, IdUsuario, Fecha, Observacion, IdResponsable) VALUES ('" +
                                            factura.TrimStart("0").ToString + "', '" +
                                            estado.ToString + "', ' ', '" +
                                            "(" + estado.ToString + ") " + estadoDescripcion.ToString + "', '" +
                                            SDS_FrmPrincipal.entidadLiquidacion.idEntidad.ToString + "', '" +
                                            SDS_FrmPrincipal.entidadLiquidacion.idSector.ToString + "', '" +
                                            ResponsableIdentificacion.ToString + "', '" +
                                            Format(Now, "yyyy-MM-dd HH:mm:ss.fff").ToString + "', 'Pago Efectuado, " +
                                            Format(Now, "yyyy-MM-dd").ToString + ".', '" + ResponsableId.ToString + "')"

                                        Dim OcomandoH As New SqlCommand(SqlInsertHistorial, conectado)
                                        OcomandoH.Connection.Open()
                                        salvar = OcomandoH.ExecuteNonQuery()
                                        OcomandoH.Connection.Close()
                                    End If
                                End If
                                contador = contador + 1
                            End If
                        End While
                        sr.Close()
                    End Using

                    
                    Using conn As New SqlClient.SqlConnection(con.ConnectionString.ToString)
                        Dim myCmd As New SqlClient.SqlCommand("actualizaEstadoSolicitud", conn)
                        myCmd.CommandType = CommandType.StoredProcedure
                        myCmd.Parameters.AddWithValue("@NOCARGUE", NOCARGUE.ToString())
                        myCmd.Parameters.AddWithValue("@TIPORESPUESTA", 1)
                        myCmd.CommandTimeout = 0

                        myCmd.Connection.Open()
                        salvar = myCmd.ExecuteNonQuery()
                        myCmd.Connection.Close()
                    End Using

                    horaFinal = Format(DateTime.Now, "yyyy-MM-dd HH:mm:ss")

                    MsgBox("Se cargó el archivo de Transferencias (Portal-Cenit) correctamente" + vbCrLf + vbCrLf +
                           "Inicio proceso: " + horaInicial + vbCrLf +
                           "Finalización proceso: " + horaFinal, MsgBoxStyle.Information)
                    '"Archivo: " + ruta.Replace("C:\Validadores\", ""), MsgBoxStyle.Information)

                Case TipoCarguePlano.CARGAGIRO
                    ticket = ObtenerTicket(tipoCargue)
                    objLOG.TM55_NoCargue = ticket
                    InsertarLOG(objLOG)

                    Dim IDSECTOR As Integer = SDS_FrmPrincipal.entidadLiquidacion.idSector
                    Dim IDENTIDAD As Integer = SDS_FrmPrincipal.entidadLiquidacion.idEntidad

                    Using conectado As New SqlClient.SqlConnection(con.ConnectionString.ToString)
                        While Not sr.EndOfStream
                            Linea = sr.ReadLine()

                            If (Not Linea Is "") Then 'OMITE LA LINEA EN BLANCO
                                SqlInsert = "INSERT INTO TM52_PLANO_CARGAGIROS (PROCESADO, NOCARGUE, IDSECTOR, IDENTIDAD, FECHA_CARGUE, TIPOIDENTIFICACIONBENEF, IDENTIFICACIONBENEF, NOGIRO, NOMBENEFICIARIO, OFRADICACION, NOCUENTACLIENTEAUT, TIPOCUENTACLIENTE, FECHADISPER, TIPOIDENTIFICLIENTE, NOCUENTACLIENTE, NOMBRECUENTACLIENTE, FORMAPAGO, VALORPAGO, CODOFICINAPAG, FECHAVENGIRO, TIPOGIRO, DIASVENCIMIENTO, CTRL, NOGIROASIGCLIENTE, COMISION, IVA, RESPUESTA, MENSAJERESPUESTA) VALUES ('0', " + ticket.ToString + ", " + IDSECTOR.ToString + ", " + IDENTIDAD.ToString + ", '" + Format(Now, "yyyy-MM-dd HH:mm:ss.fff").ToString + "', '" + Mid(Linea, 1, 1).ToString + "', '" + Mid(Linea, 2, 11).ToString + "', '" + Mid(Linea, 13, 10).ToString + "', '" + Mid(Linea, 23, 40).ToString + "', '" + Mid(Linea, 63, 3).ToString + "', '" + Mid(Linea, 66, 9).ToString + "', '" + Mid(Linea, 75, 1).ToString + "', '" + Mid(Linea, 76, 8).ToString + "', '" + Mid(Linea, 84, 1).ToString + "', '" + Mid(Linea, 85, 11).ToString + "', '" + Mid(Linea, 96, 40).ToString + "', '" + Mid(Linea, 136, 1).ToString + "', '" + Mid(Linea, 137, 15).ToString + "', '" + Mid(Linea, 152, 3).ToString + "', '" + Mid(Linea, 155, 8).ToString + "', '" + Mid(Linea, 163, 1).ToString + "', '" + Mid(Linea, 164, 3).ToString + "', '" + Mid(Linea, 167, 2).ToString + "', '" + Mid(Linea, 169, 10).ToString + "', '" + Mid(Linea, 179, 13).ToString + "', '" + Mid(Linea, 192, 9).ToString + "', '" + Mid(Linea, 201, 1).ToString + "', '" + Mid(Linea, 202, 56).ToString + "')"

                                Dim Ocomando As New SqlCommand(SqlInsert, conectado)
                                Ocomando.Connection.Open()
                                salvar = Ocomando.ExecuteNonQuery()
                                Ocomando.Connection.Close()
                            End If
                        End While
                        sr.Close()

                    End Using
                    horaFinal = Format(DateTime.Now, "yyyy-MM-dd HH:mm:ss")

                    MsgBox("Se cargó el archivo de Giros (GRSP) correctamente" + vbCrLf + vbCrLf +
                           "Inicio proceso: " + horaInicial + vbCrLf +
                           "Finalización proceso: " + horaFinal, MsgBoxStyle.Information)

                Case TipoCarguePlano.ESTADOGIRO

                    Dim TIPOREG1 As String = Nothing
                    Dim TIPOIDENTIFICCLIENTE As String = Nothing
                    Dim NOIDENTIFICCLIENTE As String = Nothing
                    Dim FECHAPROCESO As String = Nothing
                    Dim FECHAINICIO As String = Nothing
                    Dim FECHFIN As String = Nothing
                    Dim NOCUENTADEB As String = Nothing
                    Dim NOMCLIENTE As String = Nothing
                    Dim CEROS As String = Nothing
                    Dim TIPOPROC As String = Nothing
                    Dim IDENTIFICACIONCLIENTE As String = Nothing
                    Dim PROCESO As String = Nothing
                    Dim FECHAORIGEN As String = Nothing
                    Dim CEROS1 As String = Nothing
                    Dim NOGIRO As String = Nothing

                    ticket = ObtenerTicket(tipoCargue)
                    objLOG.TM55_NoCargue = ticket
                    InsertarLOG(objLOG)

                    Dim NOCARGUE As Integer = ticket
                    Dim IDSECTOR As Integer = SDS_FrmPrincipal.entidadLiquidacion.idSector
                    Dim IDENTIDAD As Integer = SDS_FrmPrincipal.entidadLiquidacion.idEntidad

                    Dim estadoGiro As String = Nothing
                    Using conectado As New SqlClient.SqlConnection(con.ConnectionString.ToString)
                        While Not sr.EndOfStream
                            Linea = sr.ReadLine()

                            If (Not Linea Is "") Then 'OMITE LA LINEA EN BLANCO
                                If (contador = 1) Then

                                    TIPOREG1 = Mid(Linea, 1, 2)
                                    TIPOIDENTIFICCLIENTE = Mid(Linea, 3, 1)
                                    NOIDENTIFICCLIENTE = Mid(Linea, 4, 13)
                                    NOIDENTIFICCLIENTE = NOIDENTIFICCLIENTE.TrimStart("0").PadRight(13, " ").ToString
                                    FECHAPROCESO = Mid(Linea, 17, 8)
                                    FECHAINICIO = Mid(Linea, 25, 8)
                                    FECHFIN = Mid(Linea, 33, 8)
                                    NOCUENTADEB = Mid(Linea, 41, 9)
                                    NOMCLIENTE = Mid(Linea, 50, 40)
                                    CEROS = Mid(Linea, 90, 14)
                                    TIPOPROC = Mid(Linea, 104, 1)
                                    IDENTIFICACIONCLIENTE = Mid(Linea, 105, 13)
                                    PROCESO = Mid(Linea, 118, 2)
                                    FECHAORIGEN = Mid(Linea, 120, 8)
                                    CEROS1 = Mid(Linea, 128, 2)
                                    SqlInsert = "INSERT INTO TM54_PLANO_ESTADOGIROS (TIPOREG1, TIPOIDENTIFICCLIENTE, NOIDENTIFICCLIENTE, FECHAPROCESO, FECHAINICIO, FECHFIN, NOCUENTADEB, NOMCLIENTE, TIPOPROC, IDENTIFICACIONCLIENTE, PROCESO, NOCARGUE, IDSECTOR, IDENTIDAD, FECHA_CARGUE, CEROS1, CEROS, FECHAORIGEN) VALUES ('" + TIPOREG1 + "', '" + TIPOIDENTIFICCLIENTE + "', '" + NOIDENTIFICCLIENTE + "', '" + FECHAPROCESO + "', '" + FECHAINICIO + "', '" + FECHFIN + "', '" + NOCUENTADEB + "', '" + NOMCLIENTE + "', '" + TIPOPROC + "', '" + IDENTIFICACIONCLIENTE + "', '" + PROCESO + "', " + NOCARGUE.ToString + ", " + IDSECTOR.ToString + ", " + IDENTIDAD.ToString + ", '" + Format(Now, "yyyy-MM-dd HH:mm:ss.fff") + "', '" + CEROS1 + "', '" + CEROS + "', '" + FECHAORIGEN + "')"
                                ElseIf Not (contador = 1) Then
                                    NOGIRO = Mid(Linea, 17, 10).ToString
                                    estadoGiro = Mid(Linea, 103, 1).ToString
                                    SqlInsert = "INSERT INTO TM54_PLANO_ESTADOGIROS (TIPOREG1, TIPOIDENTIFICCLIENTE, NOIDENTIFICCLIENTE, FECHAPROCESO, FECHAINICIO, FECHFIN, NOCUENTADEB, NOMCLIENTE, TIPOPROC, IDENTIFICACIONCLIENTE, PROCESO, NOCARGUE, IDSECTOR, IDENTIDAD,  CEROS1, CEROS, FECHAORIGEN, TIPOREG2, TIPOIDBENEF, NOIDENTIFICBENEF, NOGIRO, NOMBENEF, CODOFPAGADORA, FECHACAMBIOESTADO, VALORGIRO, ESTADOGIRO, TIPOPROC1, NOIDENTIFICCLIENTE2, TIPOPROC2, FECHAGIRO, FORMAPAGO, PROCESADO, FECHA_CARGUE) VALUES ('" + TIPOREG1.ToString() + "', '" + TIPOIDENTIFICCLIENTE.ToString() + "', '" + NOIDENTIFICCLIENTE.ToString() + "', '" + FECHAPROCESO.ToString() + "', '" + FECHAINICIO.ToString() + "', '" + FECHFIN.ToString() + "', '" + NOCUENTADEB.ToString() + "', '" + NOMCLIENTE.ToString() + "', '" + TIPOPROC.ToString() + "', '" + IDENTIFICACIONCLIENTE.ToString() + "', '" + PROCESO.ToString() + "', '" + NOCARGUE.ToString + "', '" + IDSECTOR.ToString + "', '" + IDENTIDAD.ToString + "', '" + CEROS1.ToString() + "', ' ', '" + Mid(Linea, 67, 8).ToString + "', '" + Mid(Linea, 1, 2).ToString + "', '" + Mid(Linea, 3, 1).ToString + "', '" + Mid(Linea, 4, 13).TrimStart("0").PadRight(13, " ").ToString + "', '" + NOGIRO + "', '" + Mid(Linea, 27, 40).ToString + "', '" + Mid(Linea, 75, 3).ToString + "', '" + Mid(Linea, 78, 8).ToString + "', '" + Mid(Linea, 86, 17).ToString + "', '" + estadoGiro + "', '" + Mid(Linea, 104, 1).ToString + "', '" + Mid(Linea, 105, 13).ToString + "', '" + Mid(Linea, 118, 2).ToString + "', '" + Mid(Linea, 120, 8).ToString + "', '" + Mid(Linea, 128, 1).ToString + "', '0','" + Format(Now, "yyyy-MM-dd HH:mm:ss.fff") + "')"
                                End If
                                Dim Ocomando As New SqlCommand(SqlInsert, conectado)
                                Ocomando.Connection.Open()
                                salvar = Ocomando.ExecuteNonQuery()
                                Ocomando.Connection.Close()
                                contador = contador + 1
                            End If
                        End While
                        sr.Close()
                    End Using

                    Dim estadoUpdate As Integer = 0
                    Using conn As New SqlClient.SqlConnection(con.ConnectionString.ToString)

                        Dim myCmd As New SqlClient.SqlCommand("actualizaEstadoSolicitud", conn)
                        myCmd.CommandType = CommandType.StoredProcedure
                        myCmd.Parameters.AddWithValue("@NOCARGUE", NOCARGUE.ToString())
                        myCmd.Parameters.AddWithValue("@TIPORESPUESTA", 3)
                        myCmd.CommandTimeout = 0

                        myCmd.Connection.Open()
                        salvar = myCmd.ExecuteNonQuery()
                        myCmd.Connection.Close()
                        
                    End Using
                    horaFinal = Format(DateTime.Now, "yyyy-MM-dd HH:mm:ss")

                    MsgBox("Se cargó el archivo de Giros (GIDIA) correctamente" + vbCrLf + vbCrLf +
                           "Inicio proceso: " + horaInicial + vbCrLf +
                           "Finalización proceso: " + horaFinal, MsgBoxStyle.Information)
                
                Case TipoCarguePlano.ESTADODISPERSIONHOST

                    Dim TIPOREG1 As String = Nothing
                    Dim FECHAGENERACION As String = Nothing
                    Dim HORAGENERACION As String = Nothing
                    Dim NOIDCLIENTE As String = Nothing
                    Dim TIPOCUENTACLIENTE As String = Nothing
                    Dim NOCUENTACLIENTE As String = Nothing
                    Dim FECHAABONO As String = Nothing
                    Dim VALORTOTALTRANSFER As String = Nothing
                    Dim NOREGISTROS As String = Nothing
                    Dim NUMRECHAZADOS As String = Nothing
                    Dim NOIDTRANSMISION As String = Nothing
                    Dim ESTADOENC As String = Nothing
                    Dim DESCESTADOENC As String = Nothing
                    Dim CEROS As String = Nothing

                    ticket = idTransmisionDispersion
                    objLOG.TM55_NoCargue = ticket
                    InsertarLOG(objLOG)

                    Dim NOCARGUE As Integer = ticket
                    Dim IDSECTOR As Integer = SDS_FrmPrincipal.entidadLiquidacion.idSector
                    Dim IDENTIDAD As Integer = SDS_FrmPrincipal.entidadLiquidacion.idEntidad
                    Dim facturaHtH As String = Nothing
                    Dim estadoHtH As String = Nothing
                    Dim estadoDescripcionHtH As String = Nothing

                    Using conectado As New SqlClient.SqlConnection(con.ConnectionString.ToString)
                        While Not sr.EndOfStream
                            Linea = sr.ReadLine()

                            If (Not Linea Is "") Then 'OMITE LA LINEA EN BLANCO
                                If (contador = 1) Then

                                    TIPOREG1 = Mid(Linea, 1, 1)
                                    FECHAGENERACION = Mid(Linea, 2, 8)
                                    HORAGENERACION = Mid(Linea, 10, 6)
                                    NOIDCLIENTE = Mid(Linea, 16, 11)
                                    TIPOCUENTACLIENTE = Mid(Linea, 27, 2)
                                    NOCUENTACLIENTE = Mid(Linea, 29, 17)
                                    FECHAABONO = Mid(Linea, 46, 8)
                                    VALORTOTALTRANSFER = Mid(Linea, 54, 18)
                                    NOREGISTROS = Mid(Linea, 72, 6)
                                    NUMRECHAZADOS = Mid(Linea, 78, 6)
                                    NOIDTRANSMISION = Mid(Linea, 84, 20)
                                    ESTADOENC = Mid(Linea, 104, 3)
                                    DESCESTADOENC = Mid(Linea, 107, 260)
                                    CEROS = Mid(Linea, 367, 34)

                                ElseIf Not (contador = 1) Then
                                    facturaHtH = Mid(Linea, 130, 10)
                                    estadoHtH = Mid(Linea, 140, 3)
                                    estadoDescripcionHtH = Mid(Linea, 143, 250)
                                    SqlInsert = "INSERT INTO TM53_PLANO_ESTADODISPERSION_HOST (TIPOREG1, FECHAGENERACION, HORAGENERACION, NOIDCLIENTE, TIPOCUENTACLIENTE, NOCUENTACLIENTE, FECHAABONO, VALORTOTALTRANSFER, NOREGISTROS, NUMRECHAZADOS, NOIDTRANSMISION, ESTADOENC, DESCESTADOENC, CEROS, FECHA_CARGUE, NOCARGUE, IDSECTOR, IDENTIDAD, TIPOREG2, NOMBENEF, TIPOIDENTIFICACIONBENEF, NOIDENTIFICACIONBENEF, TIPOCUENTABENEF, NOCUENTABENEF, CODCOMPENSACIONBANCO, DESCBANCO, VALORTRANSFER, CODCIUDCUENTA, FACTURA, ESTADO, DESCESTADO, FECHADEV, PROCESADO) VALUES ('" + TIPOREG1 + "', '" + FECHAGENERACION + "', '" + HORAGENERACION + "', '" + NOIDCLIENTE + "', '" + TIPOCUENTACLIENTE + "', '" + NOCUENTACLIENTE + "', '" + FECHAABONO + "', '" + VALORTOTALTRANSFER + "', '" + NOREGISTROS + "', '" + NUMRECHAZADOS + "', '" + NOIDTRANSMISION + "', '" + ESTADOENC + "', '" + DESCESTADOENC + "', '" + CEROS + "', '" + Format(Now, "yyyy-MM-dd HH:mm:ss.fff") + "', " + NOCARGUE.ToString + ", " + IDSECTOR.ToString + ", " + IDENTIDAD.ToString + ", '" + Mid(Linea, 1, 1) + "', '" + Mid(Linea, 2, 30) + "', '" + Mid(Linea, 32, 2) + "', '" + Mid(Linea, 34, 11) + "', '" + Mid(Linea, 45, 2) + "', '" + Mid(Linea, 47, 17) + "', '" + Mid(Linea, 64, 4) + "', '" + Mid(Linea, 68, 20) + "', '" + Mid(Linea, 88, 18) + "', '" + Mid(Linea, 106, 24) + "', '" + facturaHtH.ToString + "', '" + estadoHtH.ToString + "', '" + estadoDescripcionHtH.ToString + "', '" + Mid(Linea, 393, 8) + "', '0')"
                                    Dim Ocomando As New SqlCommand(SqlInsert, conectado)
                                    Ocomando.Connection.Open()
                                    salvar = Ocomando.ExecuteNonQuery()
                                    Ocomando.Connection.Close()

                                    If Not (estadoHtH = "S00" Or estadoHtH = Nothing) Then
                                        '                                         ok      ok          ''        
                                        Dim SqlInsertHistorial As String = "INSERT INTO TM70_Historial (Dato, IdCampo, ValorAnterior, ValorNuevo, IdEntidad, IdSector, IdUsuario, Fecha, Observacion, IdResponsable) VALUES ('" +
                                            facturaHtH.TrimStart("0").ToString + "', '" +
                                            estadoHtH.ToString + "', ' ', '" +
                                             "(" + estadoHtH.ToString + ") " + estadoDescripcionHtH.ToString + "', '" +
                                            SDS_FrmPrincipal.entidadLiquidacion.idEntidad.ToString + "', '" +
                                            SDS_FrmPrincipal.entidadLiquidacion.idSector.ToString + "', '" +
                                            ResponsableIdentificacion.ToString + "', '" +
                                            Format(Now, "yyyy-MM-dd HH:mm:ss.fff").ToString + "', 'Pago rechazado, " +
                                            Format(Now, "yyyy-MM-dd").ToString + ".', '" + ResponsableId.ToString + "')"

                                        Dim OcomandoH As New SqlCommand(SqlInsertHistorial, conectado)
                                        OcomandoH.Connection.Open()
                                        salvar = OcomandoH.ExecuteNonQuery()
                                        OcomandoH.Connection.Close()
                                    ElseIf (estadoHtH = "S00") Then
                                        Dim SqlInsertHistorial As String = "INSERT INTO TM70_Historial (Dato, IdCampo, ValorAnterior, ValorNuevo, IdEntidad, IdSector, IdUsuario, Fecha, Observacion, IdResponsable) VALUES ('" +
                                            facturaHtH.TrimStart("0").ToString + "', '" +
                                            estadoHtH.ToString + "', ' ', '" +
                                             "(" + estadoHtH.ToString + ") " + estadoDescripcionHtH.ToString + "', '" +
                                            SDS_FrmPrincipal.entidadLiquidacion.idEntidad.ToString + "', '" +
                                            SDS_FrmPrincipal.entidadLiquidacion.idSector.ToString + "', '" +
                                            ResponsableIdentificacion.ToString + "', '" +
                                            Format(Now, "yyyy-MM-dd HH:mm:ss.fff").ToString + "', 'Pago Efectuado, " +
                                            Format(Now, "yyyy-MM-dd").ToString + ".', '" + ResponsableId.ToString + "')"

                                        Dim OcomandoH As New SqlCommand(SqlInsertHistorial, conectado)
                                        OcomandoH.Connection.Open()
                                        salvar = OcomandoH.ExecuteNonQuery()
                                        OcomandoH.Connection.Close()
                                    End If
                                End If
                                contador = contador + 1
                            End If
                        End While
                        sr.Close()
                    End Using


                    Dim estadoUpdate As Integer = 0
                    Using conn As New SqlClient.SqlConnection(con.ConnectionString.ToString)

                        Dim myCmd As New SqlClient.SqlCommand("actualizaEstadoSolicitud", conn)
                        myCmd.CommandType = CommandType.StoredProcedure
                        myCmd.Parameters.AddWithValue("@NOCARGUE", NOCARGUE.ToString())
                        myCmd.Parameters.AddWithValue("@TIPORESPUESTA", 2)
                        myCmd.CommandTimeout = 0

                        myCmd.Connection.Open()
                        salvar = myCmd.ExecuteNonQuery()
                        myCmd.Connection.Close()
                    End Using
                    horaFinal = Format(DateTime.Now, "yyyy-MM-dd HH:mm:ss")

                    MsgBox("Se cargó el archivo de Transferencia (Host to Host) correctamente" + vbCrLf + vbCrLf +
                           "Inicio proceso: " + horaInicial + vbCrLf +
                           "Finalización proceso: " + horaFinal, MsgBoxStyle.Information)
            End Select

            objLOG.TM55_TipoCargue = tipoCargue
            objLOG.TM55_FechaCargue = Now
            objLOG.TM55_Usuario = usucargue
            objLOG.TM55_Descripcion = "FIN DEL PROCESAMIENTO DEL PLANO: " & ruta & " REGISTROS CARGADOS  " & contador
            objLOG.TM55_NoCargue = ticket
            InsertarLOG(objLOG)

            objARCHIVO.TM56_Archivo = FileIO.FileSystem.GetName(ruta)
            objARCHIVO.TM56_Ruta = ruta
            objARCHIVO.TM56_Ext = FileIO.FileSystem.GetFileInfo(ruta).Extension
            objARCHIVO.TM56_Tamano = FileIO.FileSystem.GetFileInfo(ruta).Length
            objARCHIVO.TM56_Fecha_Carga = Now
            objARCHIVO.TM56_TipoCargue = tipoCargue
            objARCHIVO.TM56_NoCargue = ticket
            objARCHIVO.TM56_Corte = Nothing
            objARCHIVO.TM56_Obs = Nothing
            objARCHIVO.TM56_idEntidad = SDS_FrmPrincipal.entidadLiquidacion.idEntidad
            objARCHIVO.TM56_idSector = SDS_FrmPrincipal.entidadLiquidacion.idSector
            InsertarARCHIVO(objARCHIVO)

            dtTemp = Nothing
            dtTemp = registrosCargados(tipoCargue, ticket)

            Return dtTemp

            ticket = 0
            longitud = Nothing
            dtTemp = Nothing

        Catch ex As Exception

            Dim objLOGERROR As New TM55_PLANO_LOG
            objLOGERROR.TM55_TipoCargue = tipoCargue
            objLOGERROR.TM55_FechaCargue = Now
            objLOGERROR.TM55_Usuario = usucargue
            objLOGERROR.TM55_Descripcion = "FIN DEL PROCESAMIENTO DEL PLANO: " & ruta & " ERROR " & ex.Message
            objLOGERROR.TM55_NoCargue = ticket
            InsertarLOG(objLOGERROR)

            Return Nothing
            longitud = Nothing
            campo = ""
            dtTemp = Nothing
        End Try

End Function