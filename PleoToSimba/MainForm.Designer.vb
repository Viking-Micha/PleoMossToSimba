<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MainForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        components = New ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
        btnStart = New Button()
        Label1 = New Label()
        btnMoss = New Button()
        picCoffee = New PictureBox()
        ToolTip1 = New ToolTip(components)
        Label2 = New Label()
        CType(picCoffee, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' btnStart
        ' 
        btnStart.Location = New Point(12, 12)
        btnStart.Name = "btnStart"
        btnStart.Size = New Size(235, 23)
        btnStart.TabIndex = 0
        btnStart.Text = "PLEO-ZIP auswählen und konvertieren"
        btnStart.UseVisualStyleBackColor = True
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Cursor = Cursors.Hand
        Label1.Location = New Point(195, 89)
        Label1.Name = "Label1"
        Label1.Size = New Size(119, 15)
        Label1.TabIndex = 1
        Label1.Text = "(c)2026 Viking-Micha"
        ' 
        ' btnMoss
        ' 
        btnMoss.Location = New Point(12, 41)
        btnMoss.Name = "btnMoss"
        btnMoss.Size = New Size(235, 23)
        btnMoss.TabIndex = 2
        btnMoss.Text = "MOSS-ZIP auswählen und konvertieren"
        btnMoss.UseVisualStyleBackColor = True
        ' 
        ' picCoffee
        ' 
        picCoffee.Cursor = Cursors.Hand
        picCoffee.Image = CType(resources.GetObject("picCoffee.Image"), Image)
        picCoffee.Location = New Point(12, 89)
        picCoffee.Name = "picCoffee"
        picCoffee.Size = New Size(24, 24)
        picCoffee.SizeMode = PictureBoxSizeMode.StretchImage
        picCoffee.TabIndex = 0
        picCoffee.TabStop = False
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(251, 104)
        Label2.Name = "Label2"
        Label2.Size = New Size(63, 15)
        Label2.TabIndex = 3
        Label2.Text = "Version 1.0"
        ' 
        ' MainForm
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(326, 125)
        Controls.Add(Label2)
        Controls.Add(picCoffee)
        Controls.Add(btnMoss)
        Controls.Add(Label1)
        Controls.Add(btnStart)
        Name = "MainForm"
        Text = "Pleo & MOSS → SIMBA Converter"
        CType(picCoffee, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents btnStart As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents btnMoss As Button

    'Friend WithEvents linkCoffee As LinkLabel
    Friend WithEvents picCoffee As PictureBox
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents Label2 As Label

End Class
