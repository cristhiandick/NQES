
Imports System.Drawing
Imports System.Windows.Forms
Imports System.ComponentModel
Public Class Form2
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
    Friend WithEvents DataGrid2 As System.Windows.Forms.DataGrid
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents DataGrid3 As System.Windows.Forms.DataGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.DataGrid2 = New System.Windows.Forms.DataGrid
        Me.DataGrid1 = New System.Windows.Forms.DataGrid
        Me.DataGrid3 = New System.Windows.Forms.DataGrid
        CType(Me.DataGrid2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGrid3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGrid2
        '
        Me.DataGrid2.DataMember = ""
        Me.DataGrid2.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid2.Location = New System.Drawing.Point(8, 16)
        Me.DataGrid2.Name = "DataGrid2"
        Me.DataGrid2.Size = New System.Drawing.Size(1056, 456)
        Me.DataGrid2.TabIndex = 15
        Me.DataGrid2.Visible = False
        '
        'DataGrid1
        '
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Location = New System.Drawing.Point(24, 16)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.Size = New System.Drawing.Size(1040, 456)
        Me.DataGrid1.TabIndex = 16
        '
        'DataGrid3
        '
        Me.DataGrid3.DataMember = ""
        Me.DataGrid3.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid3.Location = New System.Drawing.Point(12, 11)
        Me.DataGrid3.Name = "DataGrid3"
        Me.DataGrid3.Size = New System.Drawing.Size(1056, 456)
        Me.DataGrid3.TabIndex = 17
        Me.DataGrid3.Visible = False
        '
        'Form2
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.GrayText
        Me.ClientSize = New System.Drawing.Size(1080, 478)
        Me.Controls.Add(Me.DataGrid3)
        Me.Controls.Add(Me.DataGrid2)
        Me.Controls.Add(Me.DataGrid1)
        Me.Name = "Form2"
        Me.Text = "Form2"
        CType(Me.DataGrid2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGrid3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
