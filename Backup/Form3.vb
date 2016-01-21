Public Class Form3
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Chart1 As SoftwareFX.ChartFX.Lite.Chart
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Chart1 = New SoftwareFX.ChartFX.Lite.Chart
        Me.SuspendLayout()
        '
        'Chart1
        '
        Me.Chart1.Location = New System.Drawing.Point(0, 0)
        Me.Chart1.Name = "Chart1"
        Me.Chart1.Size = New System.Drawing.Size(768, 456)
        Me.Chart1.TabIndex = 0
        '
        'Form3
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(768, 454)
        Me.Controls.Add(Me.Chart1)
        Me.Name = "Form3"
        Me.Text = "Equity Curve"
        Me.ResumeLayout(False)

    End Sub

#End Region

End Class
