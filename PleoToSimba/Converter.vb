Imports System.IO
Imports System.IO.Compression
Imports Microsoft.VisualBasic.FileIO
Imports System.Linq
Imports System.Text


Module Converter
    Public Sub MOSSExport(zipPath As String)
        If String.IsNullOrWhiteSpace(zipPath) OrElse Not File.Exists(zipPath) Then
            Throw New FileNotFoundException("Die angegebene ZIP-Datei wurde nicht gefunden.", zipPath)
        End If

        Dim zipOrdner As String = Path.GetDirectoryName(zipPath)
        Dim arbeitsOrdner As String = Path.Combine(zipOrdner, "MOSS_TEMP")

        If Directory.Exists(arbeitsOrdner) Then
            FileIO.FileSystem.DeleteDirectory(arbeitsOrdner, FileIO.DeleteDirectoryOption.DeleteAllContents)
        End If
        Directory.CreateDirectory(arbeitsOrdner)

        ZipFile.ExtractToDirectory(zipPath, arbeitsOrdner)

        ' attachments.zip suchen
        Dim attachmentZip As String = Nothing
        Dim attachmentZips = Directory.GetFiles(arbeitsOrdner, "attachments.zip", System.IO.SearchOption.TopDirectoryOnly)

        If attachmentZips.Length > 0 Then
            attachmentZip = attachmentZips(0)
        End If

        If String.IsNullOrWhiteSpace(attachmentZip) Then
            Throw New Exception("Keine attachments.zip im MOSS-Export gefunden.")
        End If

        Dim attachmentFolder As String = Path.Combine(arbeitsOrdner, "attachments")
        Directory.CreateDirectory(attachmentFolder)
        ZipFile.ExtractToDirectory(attachmentZip, attachmentFolder)

        ' Alle Dateien aus attachments ins Root verschieben
        For Each attachmentFile As String In Directory.GetFiles(attachmentFolder, "*.*", System.IO.SearchOption.AllDirectories)
            Dim target As String = Path.Combine(arbeitsOrdner, Path.GetFileName(attachmentFile))

            If File.Exists(target) Then
                File.Delete(target)
            End If

            File.Move(attachmentFile, target)
        Next

        FileIO.FileSystem.DeleteDirectory(attachmentFolder, FileIO.DeleteDirectoryOption.DeleteAllContents)
        File.Delete(attachmentZip)

        ' CSV suchen
        Dim csvFiles = Directory.GetFiles(arbeitsOrdner, "*.csv", System.IO.SearchOption.TopDirectoryOnly)

        If csvFiles.Length = 0 Then
            Throw New Exception("Keine CSV-Datei im MOSS-Export gefunden.")
        End If

        Dim csvFile As String = csvFiles(0)

        ' Zielname mit Zählerlogik
        Dim datum As String = DateTime.Now.ToString("yyyy_MM_dd")
        Dim basisName As String = "EXTF_" & datum
        Dim csvName As String = basisName & ".csv"

        Dim counter As Integer = 1
        While File.Exists(Path.Combine(zipOrdner, csvName & ".zip"))
            csvName = basisName & "_" & counter.ToString("000") & ".csv"
            counter += 1
        End While

        Dim neuerCsvPfad As String = Path.Combine(arbeitsOrdner, csvName)

        If File.Exists(neuerCsvPfad) Then
            File.Delete(neuerCsvPfad)
        End If

        File.Move(csvFile, neuerCsvPfad)

        ' EXTF-Header ergänzen
        Dim extfHeader As String = BuildMossExtfHeader(neuerCsvPfad)
        Dim bestehendeZeilen As New List(Of String)(File.ReadAllLines(neuerCsvPfad, Encoding.UTF8))
        bestehendeZeilen.Insert(0, extfHeader)

        Dim utf8Bom As New UTF8Encoding(True)
        File.WriteAllLines(neuerCsvPfad, bestehendeZeilen, utf8Bom)

        ' ZIP erzeugen: EXTF_YYYY_MM_DD(.csv).zip bzw. mit _001, _002 ...
        Dim newZip As String = Path.Combine(zipOrdner, csvName & ".zip")

        If File.Exists(newZip) Then
            File.Delete(newZip)
        End If

        ZipFile.CreateFromDirectory(arbeitsOrdner, newZip)

        FileIO.FileSystem.DeleteDirectory(arbeitsOrdner, FileIO.DeleteDirectoryOption.DeleteAllContents)
    End Sub


    Public Sub PLEOExport(zipPfad As String)
        If String.IsNullOrWhiteSpace(zipPfad) OrElse Not File.Exists(zipPfad) Then
            Throw New FileNotFoundException("Die angegebene ZIP-Datei wurde nicht gefunden.", zipPfad)
        End If

        ' Basisordner = Ordner der ZIP-Datei
        Dim zipOrdner As String = Path.GetDirectoryName(zipPfad)
        Dim arbeitsOrdner As String = Path.Combine(zipOrdner, "PLEO_TEMP")
        Dim PZiel As String = arbeitsOrdner

        ' Arbeitsordner neu anlegen
        If Directory.Exists(arbeitsOrdner) Then
            FileIO.FileSystem.DeleteDirectory(arbeitsOrdner, FileIO.DeleteDirectoryOption.DeleteAllContents)
        End If
        Directory.CreateDirectory(arbeitsOrdner)

        ' ZIP entpacken
        ZipFile.ExtractToDirectory(zipPfad, PZiel)

        ' PDFs aus "receipts" nach oben ziehen
        Dim receiptsOrdner As String = Path.Combine(PZiel, "receipts")
        If Directory.Exists(receiptsOrdner) Then
            FileIO.FileSystem.CopyDirectory(receiptsOrdner, PZiel, overwrite:=True)
            FileIO.FileSystem.DeleteDirectory(receiptsOrdner, FileIO.DeleteDirectoryOption.DeleteAllContents)
        End If

        ' EXTF-CSV suchen
        Dim csvFiles = FileIO.FileSystem.GetFiles(
    PZiel,
    Microsoft.VisualBasic.FileIO.SearchOption.SearchTopLevelOnly,
    "EXTF*.csv")


        If csvFiles.Count = 0 Then
            Throw New Exception("Keine EXTF-CSV im entpackten Ordner gefunden.")
        End If

        Dim PLEOcsv As String = csvFiles(0)

        ' Temp-CSV anlegen
        Dim csvTempPfad As String = Path.Combine(PZiel, "csvtemp.csv")

        FileIO.FileSystem.CopyFile(PLEOcsv, csvTempPfad, overwrite:=True)
        FileIO.FileSystem.DeleteFile(PLEOcsv)

        Dim zeileninhalt As String = String.Empty

        ' Neue CSV schreiben basierend auf csvtemp
        Using MyReader As New TextFieldParser(csvTempPfad)
            MyReader.TextFieldType = FieldType.Delimited
            MyReader.SetDelimiters(";")

            Dim currentRow As String()
            Dim zaehler As Integer = 1

            While Not MyReader.EndOfData
                Try
                    currentRow = MyReader.ReadFields()

                    zeileninhalt = String.Empty

                    If zaehler > 2 Then
                        ' Ab Zeile 3: Spalte K (Index 10) → Spalte T (Index 19) + ".pdf"
                        If currentRow.Length > 19 AndAlso currentRow.Length > 10 Then
                            currentRow(19) = currentRow(10) & ".pdf"
                        End If
                    End If

                    For Each currentField As String In currentRow
                        zeileninhalt &= currentField & ";"
                    Next

                    ' Letztes Semikolon entfernen
                    If zeileninhalt.Length > 0 Then
                        zeileninhalt = zeileninhalt.Substring(0, zeileninhalt.Length - 1)
                    End If

                    zeileninhalt &= Environment.NewLine

                    ' An Zieldatei anhängen (wird bei erstem Aufruf erstellt)
                    My.Computer.FileSystem.WriteAllText(PLEOcsv, zeileninhalt, True)

                    zaehler += 1

                Catch ex As MalformedLineException
                    ' Fehlerhafte Zeile überspringen, optional loggen
                    ' Hier könntest du bei Bedarf eine Protokollierung einbauen
                End Try
            End While
        End Using

        ' Temp-CSV löschen
        FileIO.FileSystem.DeleteFile(csvTempPfad)

        ' Dateinamen der CSV ermitteln (ohne Pfad)
        Dim csvDateiname As String = Path.GetFileName(PLEOcsv)

        ' Ziel-ZIP: gleicher Ordner wie Original-ZIP, Name: EXTF_xxx.csv.zip
        Dim newZipPfad As String = Path.Combine(zipOrdner, csvDateiname & ".zip")

        ' Falls schon vorhanden: löschen
        If File.Exists(newZipPfad) Then
            File.Delete(newZipPfad)
        End If

        ' ZIP mit CSV + PDFs erstellen
        ZipFile.CreateFromDirectory(PZiel, newZipPfad)

        ' Arbeitsordner aufräumen
        FileIO.FileSystem.DeleteDirectory(PZiel, FileIO.DeleteDirectoryOption.DeleteAllContents)

        ' Optional: Original-PLEO-ZIP löschen? -> hier bewusst NICHT
        ' FileIO.FileSystem.DeleteFile(zipPfad)

    End Sub

    Private Function ParseCsvLine(line As String) As List(Of String)
        Dim result As New List(Of String)
        Dim current As New StringBuilder()
        Dim inQuotes As Boolean = False

        For i As Integer = 0 To line.Length - 1
            Dim ch As Char = line(i)

            If ch = """"c Then
                If inQuotes AndAlso i < line.Length - 1 AndAlso line(i + 1) = """"c Then
                    current.Append(""""c)
                    i += 1
                Else
                    inQuotes = Not inQuotes
                End If
            ElseIf ch = ";"c AndAlso Not inQuotes Then
                result.Add(current.ToString())
                current.Clear()
            Else
                current.Append(ch)
            End If
        Next

        result.Add(current.ToString())
        Return result
    End Function

    Private Function ErmittleSachkontenlaenge(csvPath As String) As Integer
        Dim lines() As String = File.ReadAllLines(csvPath, Encoding.UTF8)

        If lines.Length < 2 Then
            Return 4
        End If

        Dim laengen As New Dictionary(Of Integer, Integer)

        ' Zeile 1 = EXTF-Header (kommt bei MOSS später dazu)
        ' Zeile 2 = Spaltenüberschrift
        ' Daten beginnen aktuell bei MOSS in Zeile 2, daher hier ab Zeile 2 der Datei
        For i As Integer = 1 To lines.Length - 1
            If String.IsNullOrWhiteSpace(lines(i)) Then Continue For

            Dim row As List(Of String) = ParseCsvLine(lines(i))

            ' Konto = Spalte 7 (Index 6)
            ' Gegenkonto = Spalte 8 (Index 7)
            If row.Count > 7 Then
                ZaehleKontolaenge(row(6), laengen)
                ZaehleKontolaenge(row(7), laengen)
            End If
        Next

        If laengen.Count = 0 Then
            Return 4
        End If

        Dim haeufigsteLaenge As Integer = 4
        Dim maxTreffer As Integer = -1

        For Each eintrag In laengen
            If eintrag.Value > maxTreffer Then
                haeufigsteLaenge = eintrag.Key
                maxTreffer = eintrag.Value
            ElseIf eintrag.Value = maxTreffer AndAlso eintrag.Key = 4 Then
                ' Bei Gleichstand lieber 4-stellig bevorzugen
                haeufigsteLaenge = 4
            End If
        Next

        Return haeufigsteLaenge
    End Function

    Private Sub ZaehleKontolaenge(wert As String, laengen As Dictionary(Of Integer, Integer))
        If String.IsNullOrWhiteSpace(wert) Then Exit Sub

        Dim nurZiffern As String = New String(wert.Where(Function(c) Char.IsDigit(c)).ToArray())

        If String.IsNullOrWhiteSpace(nurZiffern) Then Exit Sub

        Dim len As Integer = nurZiffern.Length

        If Not laengen.ContainsKey(len) Then
            laengen(len) = 0
        End If

        laengen(len) += 1
    End Sub

    Private Function BuildMossExtfHeader(csvPath As String) As String
        Dim lines() As String = File.ReadAllLines(csvPath, Encoding.UTF8)

        If lines.Length < 2 Then
            Throw New Exception("Die MOSS-CSV enthält keine Buchungsdaten.")
        End If

        Dim timestamp As String = DateTime.Now.ToString("yyyyMMddHHmmss")
        Dim wirtschaftsjahr As String = DateTime.Now.Year.ToString() & "0101"

        Dim minDate As DateTime? = Nothing
        Dim maxDate As DateTime? = Nothing

        For i As Integer = 1 To lines.Length - 1
            If String.IsNullOrWhiteSpace(lines(i)) Then Continue For

            Dim row As List(Of String) = ParseCsvLine(lines(i))

            ' Belegdatum = Spalte 10 (Index 9), Format z.B. 2001 = 20.01.
            If row.Count > 9 Then
                Dim belegdatum As String = row(9).Trim()

                If belegdatum.Length = 4 AndAlso IsNumeric(belegdatum) Then
                    Dim tag As Integer = Integer.Parse(belegdatum.Substring(0, 2))
                    Dim monat As Integer = Integer.Parse(belegdatum.Substring(2, 2))
                    Dim jahr As Integer = DateTime.Now.Year

                    Try
                        Dim dt As New DateTime(jahr, monat, tag)

                        If Not minDate.HasValue OrElse dt < minDate.Value Then minDate = dt
                        If Not maxDate.HasValue OrElse dt > maxDate.Value Then maxDate = dt
                    Catch
                        ' ungültiges Datum ignorieren
                    End Try
                End If
            End If
        Next

        If Not minDate.HasValue OrElse Not maxDate.HasValue Then
            minDate = New DateTime(DateTime.Now.Year, DateTime.Now.Month, 1)
            maxDate = minDate.Value.AddMonths(1).AddDays(-1)
        End If

        wirtschaftsjahr = minDate.Value.Year.ToString() & "0101"

        Dim sachkontenlaenge As Integer = ErmittleSachkontenlaenge(csvPath)

        Dim fields As New List(Of String) From {
            "EXTF",
            "700",
            "21",
            "Buchungsstapel",
            "9",
            timestamp,
            "",
            "RE",
            "",
            "",
            "1",
            "1",
            wirtschaftsjahr,
            sachkontenlaenge.ToString(),
            minDate.Value.ToString("yyyyMMdd"),
            maxDate.Value.ToString("yyyyMMdd"),
            "MOSS"
        }

        ' Restliche leere Felder wie beim funktionierenden PLEO-Header auffüllen
        For i As Integer = 1 To 14
            fields.Add("")
        Next

        Return String.Join(";", fields)
    End Function

End Module
