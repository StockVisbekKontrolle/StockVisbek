Imports MySql.Data.MySqlClient
Imports System.IO



Public Class clsDatenTransfer

    Dim strDatenBankZugang As String =
    "server=192.168.0.2;" &
    "port=3307;" &
    "uid=root;" &
    "pwd=csm101;" &
    "database=csd2;"


    Dim conMySQL As New MySqlConnection(strDatenBankZugang)
    Dim datumVon As String = ""
    Dim datumBis As String = ""

    Dim strbldContainer As New System.Text.StringBuilder

    Public Sub DatenbankTabelleLaden()

        '*** Datenbank öffnen ***
        conMySQL.Open()

        '*** Daten an Tabelle übergeben ***
        Dim strSQL As String = "SELECT Orders.FirstLoadingContainerNum AS Containernummer, " &
                        "Pallet_Types.MatchCode AS Containerart, " &
                        "Orders.OrderNum AS Fahrtnummer, " &
                        "Waypoints.WayPointFromTime AS Gestellung, " &
                        "Waypoints.AddressName AS Firmenname, " &
                        "Waypoints.City AS Ort, " &
                        "Waypoints.WayPointToTime AS Closing, " &
                        "Ship_Owners.MatchCode AS Reederei, " &
                        "Waypoints.WaypointDetailTypeID AS WegePunktTyp, " &
                        "Orders.ImportExport " &
                        "FROM Orders " &
                        "LEFT JOIN Addresses ON (Orders.OrdererAddressID) = (Addresses.AddressID) " &
                        "INNER JOIN WayPoints ON (Orders.OrderID = Waypoints.WayPointOwnerId) " &
                        "LEFT JOIN Pallet_Types ON (Waypoints.PalletTypeID = Pallet_Types.PalletTypeID) " &
                        "LEFT JOIN Ship_Owners ON (Waypoints.ShipOwnerID = Ship_Owners.ShipOwnerID) " &
                        "WHERE (Orders.OrderDispoDate BETWEEN '" & datumVon & "' AND '" & datumBis & "') " &
                        "AND (Orders.Deleted = 0) " &
                        "AND (Waypoints.WaypointDetailTypeID IN (16, 20, 5)) " & '5 eingebunden, um Closing zu ermitteln
                        "AND (Waypoints.WayPointOwnerType = 'order')"

        '1 = Aufnahme
        '2 = ImportGestellung?
        '3 = ExportGestellung?
        '4 = Zollstop
        '5 = Rückgabe
        '14 = absatteln
        '15 = aufsatteln
        '16 = Gestellung
        '20 = Ab-/Aufkranen

        Dim cmd As New MySqlCommand(strSQL, conMySQL)
        Dim dr As MySqlDataReader = cmd.ExecuteReader

        Form1.dt.Load(dr)

        'Datareader schliessen, sonst blockiert er die conn
        'es gilt: nur ein Reader pro conn Objekt möglich!
        'und: ein conn Objekt pro Reader
        dr.Close()
        cmd.Dispose()

        conMySQL.Close()
        conMySQL.Dispose()

        subKennzeichenHolen


    End Sub

    Sub subKennzeichenHolen()

        conMySQL.Open()

        Dim strSQL As String =
            "SELECT Orders.OrderNum AS Fahrtnummer, " &
            "Trailers.Matchcode AS Chassis " &
            "FROM Orders " &
            "LEFT JOIN Trips ON (Orders.OrderID = Trips.OrderID) " &
            "LEFT JOIN Trailers ON (Trips.TrailerID = Trailers.TrailerID) " &
            "LEFT JOIN WayPoints ON (Orders.OrderID = Waypoints.WayPointOwnerId) " &
            "WHERE (Waypoints.WaypointFromtime BETWEEN '" & datumVon & "' AND '" & datumBis & "') " &
            "AND (Orders.Deleted = 0) " &
            "AND (Trips.Deleted = 0) " &
            "AND (Waypoints.WayPointOwnerType = 'order')"


        Dim cmd As New MySqlCommand(strSQL, conMySQL)

        Dim dr As MySqlDataReader = cmd.ExecuteReader

        Form1.dtKennzeichen.Load(dr)


        cmd.Dispose()
        dr.Close()
        dr.Dispose()
        conMySQL.Close()
        conMySQL.Dispose()
    End Sub


    Sub DataTableErgaenzen()
        'Datatable ergänzen
        'mit ContainerNummer ohne Format zum leichteren Filtern


        For Each row In Form1.dt.Rows

            'Sonderfall Containernummer
            'Ereignis KeyDown wird in TxtBoxContainernummer ausgelöst
            'Vergleich der Angabe soll aber in dem GridViewFeld ContainernummerOhneFormat
            'stattfinden
            strbldContainer.Clear()
            strbldContainer.Append(row!ContainerNummer)
            strbldContainer.Replace(" ", "")
            strbldContainer.Replace(".", "")
            strbldContainer.Replace("-", "")
            row!ContainerNummerOhneFormat = strbldContainer.ToString


        Next



    End Sub

    Public Sub ZeitRaumeinrichten()
        'Zeitraum der Datenübertragung festlegen
        datumVon = DateAdd(DateInterval.Day, -90, Today)
        datumBis = DateAdd(DateInterval.Day, 30, Today)

        'Neuzuordnung wegen Test....Programmierung der neuen Statustabellen
        ' datumVon = "24.07.2021"
        'datumBis = "24.07.2021"

        datumVon = Strings.Right(datumVon, 4) & "-" & Strings.Mid(datumVon, 4, 2) & "-" & Strings.Left(datumVon, 2) & " 00:00"
        datumBis = Strings.Right(datumBis, 4) & "-" & Strings.Mid(datumBis, 4, 2) & "-" & Strings.Left(datumBis, 2) & " 23:59"
    End Sub

    Public Function fncClosing(strFbNummer As String) As String

        Dim result() As DataRow

        result = Form1.dt.Select("[Fahrtnummer] = '" & strFbNummer & "' AND ([WegePunktTyp] = 5) ")


        If result.Length > 0 Then
            fncClosing = result(0)!Closing
        Else
            fncClosing = ""
        End If

        If fncClosing = "00:00:00" Then fncClosing = ""

        Return Strings.Left(fncClosing, 16)

    End Function

    Public Function fncReederei(strFbNummer As String) As String

        Dim result() As DataRow

        result = Form1.dt.Select("[Fahrtnummer] = '" & strFbNummer & "' AND ([WegePunktTyp] = 5) ")


        If result.Length > 0 Then
            fncReederei = result(0)!Reederei
        Else
            fncReederei = ""
        End If


        Return fncReederei

    End Function

    Public Function fncKennzeichen(strFbNummer As String) As String

        Dim result() As DataRow

        result = Form1.dtKennzeichen.Select("[Fahrtnummer] = '" & strFbNummer & "'")

        If result(0)!Chassis Is DBNull.Value Then result(0)!Chassis = ""

        If result.Length > 0 Then
            fncKennzeichen = result(0)!Chassis
        Else
            fncKennzeichen = ""
        End If

        Return fncKennzeichen

    End Function

    Public Sub ChassisKennzeichenHolen()


        '*** TextDatei Kennzeichen einlesen
        Dim strPfadDatei As String = "F:\StockVisbek\KennzeichenChassis.txt"
        Dim strZeile As String = ""


        Form1.dtKennzeichenChassis.Columns.Add("Kennzeichen", GetType(String))

        Dim strmRead As New System.IO.StreamReader(strPfadDatei)

        strZeile = strmRead.ReadLine

        Do While (Not strZeile Is Nothing)

            'neue Daten der Tabelle hinzufügen
            Dim row As DataRow = Form1.dtKennzeichenChassis.NewRow

            row!Kennzeichen = Strings.Mid(strZeile, 36, 5)

            Form1.dtKennzeichenChassis.Rows.Add(row)


            'nächste Zeile lesen
            strZeile = strmRead.ReadLine
        Loop




    End Sub

    Public Sub XMLspeichern()

        If File.Exists(Form1.strPfad & Form1.strDateiName & Form1.txtBoxDatum.Text & ".xml") Then
            'Datei nicht überschreiben
            'wenn sie bereits besteht
            Try
                Form1.dtStock.WriteXml(Form1.strPfad & Form1.strDateiName & Form1.txtBoxDatum.Text & "_" & Now.ToString("HHmm") & ".xml", XmlWriteMode.WriteSchema)
                Form1.lblSpeicherDatum.Text = Now.ToString("HH:mm")
            Catch ex As Exception
                Form1.lblSpeicherDatum.Text = ""
                MsgBox("Daten konnten nicht gespeichert werden." & vbCrLf & "Keine Verbindung zum Server F:\")
            End Try
        Else
            Try
                Form1.dtStock.WriteXml(Form1.strPfad & Form1.strDateiName & Form1.txtBoxDatum.Text & ".xml", XmlWriteMode.WriteSchema)
                Form1.lblSpeicherDatum.Text = Now.ToString("HH:mm")
            Catch ex As Exception
                Form1.lblSpeicherDatum.Text = ""
                MsgBox("Daten konnten nicht gespeichert werden." & vbCrLf & "Keine Verbindung zum Server F:\")
            End Try
        End If
    End Sub


End Class