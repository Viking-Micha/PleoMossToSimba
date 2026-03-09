Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class MainForm

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Prüfen, ob das Programm mit einer Datei als Argument gestartet wurde (Drag&Drop)
        Dim args() As String = Environment.GetCommandLineArgs()
        Me.AcceptButton = btnStart
        Me.StartPosition = FormStartPosition.CenterScreen
        ToolTip1.SetToolTip(picCoffee, "Wenn dir das Tool Arbeit spart ☕")
        ToolTip1.SetToolTip(Label1, "Digitale Bananenrepublik auf Instagram 🍌")

        ' args(0) = Pfad zur EXE, args(1) = erste übergebene Datei (falls vorhanden)
        If args.Length = 2 Then
            Dim zipPfad As String = args(1)

            If File.Exists(zipPfad) Then
                Try
                    Converter.PLEOExport(zipPfad)
                    MessageBox.Show("Fertig. Die Simba-ZIP wurde im gleichen Ordner wie die PLEO-ZIP erstellt.",
                                    "PleoToSimba", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show("Fehler bei der Verarbeitung:" & Environment.NewLine & ex.Message,
                                    "PleoToSimba", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Finally
                    ' Im Drag&Drop-Fall brauchen wir das Fenster nicht offen
                    Me.Close()
                End Try
            End If
        End If

        ' Wenn args.Length <= 1 ist: ganz normal öffnen, Button verwenden
    End Sub

    Private Sub picCoffee_Click(sender As Object, e As EventArgs) Handles picCoffee.Click
        Process.Start(New ProcessStartInfo("https://buymeacoffee.com/VikingMicha") With {.UseShellExecute = True})
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        Process.Start(New ProcessStartInfo("https://www.instagram.com/digitale_bananenrepublik?igsh=N2lla2tjM3N6YmVw&utm_source=qr") With {.UseShellExecute = True})
    End Sub

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        Using ofd As New OpenFileDialog()
            ofd.Filter = "ZIP-Dateien|*.zip"
            ofd.Title = "PLEO-Export auswählen"

            If ofd.ShowDialog() = DialogResult.OK Then
                Try
                    Converter.PLEOExport(ofd.FileName)
                    MessageBox.Show("Fertig. Die Simba-ZIP liegt im gleichen Ordner wie die PLEO-ZIP.",
                                    "PleoToSimba", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show("Fehler bei der Verarbeitung:" & Environment.NewLine & ex.Message,
                                    "PleoToSimba", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End Using
    End Sub

    Private Sub btnMoss_Click(sender As Object, e As EventArgs) Handles btnMoss.Click

        Dim openFileDialog As New OpenFileDialog
        openFileDialog.Filter = "ZIP-Dateien (*.zip)|*.zip"
        openFileDialog.Title = "MOSS-Export auswählen"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Converter.MOSSExport(openFileDialog.FileName)
            MessageBox.Show("MOSS-Datei erfolgreich konvertiert.")
        End If

    End Sub

End Class
