
Imports System.Data.OleDb
Imports TA
Imports System.Drawing
Imports System.Windows.Forms
Imports System.ComponentModel
Imports System.IO

Public Class Form1
    Inherits System.Windows.Forms.Form
    Private myConnection As OleDbConnection
    Private myConnection2 As OleDbConnection
    Private myConnection3 As OleDbConnection
    Private myAdapter As OleDbDataAdapter
    Private myAdapter2 As OleDbDataAdapter
    Private myAdapter3 As OleDbDataAdapter
    Private myDataset As DataSet
    Private myDataset2 As DataSet
    Public myDataset3 As DataSet
    Public myDataset4 As DataSet
    Private myStats As Stats
    Dim myFormData As Form2
    Dim myFormChart As Form3
    Dim myFormStats As Form2
    Dim myFormTrades As Form2
    Dim lossdollar As Double
    Dim windollar As Double
    Dim daycountl As Integer
    Dim daycounts As Integer
    Dim exportcount As Integer
    Dim exportcountcurve As Integer
    Dim runloop As Boolean




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
    Friend WithEvents GrabData As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtMA1 As System.Windows.Forms.TextBox
    Friend WithEvents txtMA2 As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rdSMA As System.Windows.Forms.RadioButton
    Friend WithEvents rdEMA As System.Windows.Forms.RadioButton
    Friend WithEvents rdWEMA As System.Windows.Forms.RadioButton
    Friend WithEvents txtEndId As System.Windows.Forms.TextBox
    Friend WithEvents txtStartID As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtPL As System.Windows.Forms.Button
    Friend WithEvents txtExport As System.Windows.Forms.Button
    Friend WithEvents Trade As System.Windows.Forms.Button
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents RunAll As System.Windows.Forms.Button
    Friend WithEvents ClearData As System.Windows.Forms.Button
    Friend WithEvents groupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents lbsepl As System.Windows.Forms.Label
    Friend WithEvents txtSEPL2 As System.Windows.Forms.Label
    Friend WithEvents txtLEPL2 As System.Windows.Forms.Label
    Friend WithEvents lblepl As System.Windows.Forms.Label
    Friend WithEvents label49 As System.Windows.Forms.Label
    Friend WithEvents lbunithigh As System.Windows.Forms.Label
    Friend WithEvents label44 As System.Windows.Forms.Label
    Friend WithEvents label46 As System.Windows.Forms.Label
    Friend WithEvents lbDayHigh As System.Windows.Forms.Label
    Friend WithEvents label40 As System.Windows.Forms.Label
    Friend WithEvents lbratiodoll As System.Windows.Forms.Label
    Friend WithEvents label38 As System.Windows.Forms.Label
    Friend WithEvents lbnetdailyaverage As System.Windows.Forms.Label
    Friend WithEvents label15 As System.Windows.Forms.Label
    Friend WithEvents label27 As System.Windows.Forms.Label
    Friend WithEvents label28 As System.Windows.Forms.Label
    Friend WithEvents ratioD As System.Windows.Forms.Label
    Friend WithEvents label30 As System.Windows.Forms.Label
    Friend WithEvents label31 As System.Windows.Forms.Label
    Friend WithEvents lbnetdailylow As System.Windows.Forms.Label
    Friend WithEvents label33 As System.Windows.Forms.Label
    Friend WithEvents label34 As System.Windows.Forms.Label
    Friend WithEvents lblossday As System.Windows.Forms.Label
    Friend WithEvents lbwinday As System.Windows.Forms.Label
    Friend WithEvents lbnetdailyhigh As System.Windows.Forms.Label
    Friend WithEvents lbtotalpl As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents label17 As System.Windows.Forms.Label
    Friend WithEvents txtwlratio As System.Windows.Forms.Label
    Friend WithEvents label18 As System.Windows.Forms.Label
    Friend WithEvents label19 As System.Windows.Forms.Label
    Friend WithEvents lbeql As System.Windows.Forms.Label
    Friend WithEvents label25 As System.Windows.Forms.Label
    Friend WithEvents label26 As System.Windows.Forms.Label
    Friend WithEvents lbloss As System.Windows.Forms.Label
    Friend WithEvents lbwin As System.Windows.Forms.Label
    Friend WithEvents lbeqh As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents rdShowForm2 As System.Windows.Forms.CheckBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtTickerA As System.Windows.Forms.TextBox
    Friend WithEvents txtTickerB As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtZExit As System.Windows.Forms.TextBox
    Friend WithEvents txtZEntry As System.Windows.Forms.TextBox
    Friend WithEvents txtMultA As System.Windows.Forms.TextBox
    Friend WithEvents txtMultB As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents rbMVHR As System.Windows.Forms.RadioButton
    Friend WithEvents rbFixed As System.Windows.Forms.RadioButton
    Friend WithEvents txtRatioB As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtRatioA As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents rbPrice As System.Windows.Forms.RadioButton
    Friend WithEvents rbYield As System.Windows.Forms.RadioButton
    Friend WithEvents rbDVO1 As System.Windows.Forms.RadioButton
    Friend WithEvents cbTradeLog As System.Windows.Forms.CheckBox
    Friend WithEvents cbCurve As System.Windows.Forms.CheckBox
    Friend WithEvents txtMaxUnits As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Chart As System.Windows.Forms.Button
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GrabData = New System.Windows.Forms.Button
        Me.Trade = New System.Windows.Forms.Button
        Me.ClearData = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtMA1 = New System.Windows.Forms.TextBox
        Me.txtMA2 = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.rdWEMA = New System.Windows.Forms.RadioButton
        Me.rdEMA = New System.Windows.Forms.RadioButton
        Me.rdSMA = New System.Windows.Forms.RadioButton
        Me.txtEndId = New System.Windows.Forms.TextBox
        Me.txtStartID = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtMultB = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtPL = New System.Windows.Forms.Button
        Me.txtExport = New System.Windows.Forms.Button
        Me.txtMultA = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtTickerA = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.RunAll = New System.Windows.Forms.Button
        Me.groupBox3 = New System.Windows.Forms.GroupBox
        Me.lbsepl = New System.Windows.Forms.Label
        Me.txtSEPL2 = New System.Windows.Forms.Label
        Me.txtLEPL2 = New System.Windows.Forms.Label
        Me.lblepl = New System.Windows.Forms.Label
        Me.label49 = New System.Windows.Forms.Label
        Me.lbunithigh = New System.Windows.Forms.Label
        Me.label44 = New System.Windows.Forms.Label
        Me.label46 = New System.Windows.Forms.Label
        Me.lbDayHigh = New System.Windows.Forms.Label
        Me.label40 = New System.Windows.Forms.Label
        Me.lbratiodoll = New System.Windows.Forms.Label
        Me.label38 = New System.Windows.Forms.Label
        Me.lbnetdailyaverage = New System.Windows.Forms.Label
        Me.label15 = New System.Windows.Forms.Label
        Me.label27 = New System.Windows.Forms.Label
        Me.label28 = New System.Windows.Forms.Label
        Me.ratioD = New System.Windows.Forms.Label
        Me.label30 = New System.Windows.Forms.Label
        Me.label31 = New System.Windows.Forms.Label
        Me.lbnetdailylow = New System.Windows.Forms.Label
        Me.label33 = New System.Windows.Forms.Label
        Me.label34 = New System.Windows.Forms.Label
        Me.lblossday = New System.Windows.Forms.Label
        Me.lbwinday = New System.Windows.Forms.Label
        Me.lbnetdailyhigh = New System.Windows.Forms.Label
        Me.lbtotalpl = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.Label24 = New System.Windows.Forms.Label
        Me.Label22 = New System.Windows.Forms.Label
        Me.label17 = New System.Windows.Forms.Label
        Me.txtwlratio = New System.Windows.Forms.Label
        Me.label18 = New System.Windows.Forms.Label
        Me.label19 = New System.Windows.Forms.Label
        Me.lbeql = New System.Windows.Forms.Label
        Me.label25 = New System.Windows.Forms.Label
        Me.label26 = New System.Windows.Forms.Label
        Me.lbloss = New System.Windows.Forms.Label
        Me.lbwin = New System.Windows.Forms.Label
        Me.lbeqh = New System.Windows.Forms.Label
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.rdShowForm2 = New System.Windows.Forms.CheckBox
        Me.txtZExit = New System.Windows.Forms.TextBox
        Me.txtZEntry = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.txtTickerB = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.rbDVO1 = New System.Windows.Forms.RadioButton
        Me.rbMVHR = New System.Windows.Forms.RadioButton
        Me.rbFixed = New System.Windows.Forms.RadioButton
        Me.txtRatioB = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.txtRatioA = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.rbPrice = New System.Windows.Forms.RadioButton
        Me.rbYield = New System.Windows.Forms.RadioButton
        Me.cbTradeLog = New System.Windows.Forms.CheckBox
        Me.cbCurve = New System.Windows.Forms.CheckBox
        Me.txtMaxUnits = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Chart = New System.Windows.Forms.Button
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.GroupBox1.SuspendLayout()
        Me.groupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'GrabData
        '
        Me.GrabData.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
        Me.GrabData.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GrabData.Location = New System.Drawing.Point(8, 16)
        Me.GrabData.Name = "GrabData"
        Me.GrabData.Size = New System.Drawing.Size(128, 40)
        Me.GrabData.TabIndex = 0
        Me.GrabData.Text = "Grab Data"
        '
        'Trade
        '
        Me.Trade.BackColor = System.Drawing.Color.FromArgb(CType(128, Byte), CType(255, Byte), CType(128, Byte))
        Me.Trade.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Trade.Location = New System.Drawing.Point(8, 96)
        Me.Trade.Name = "Trade"
        Me.Trade.Size = New System.Drawing.Size(128, 40)
        Me.Trade.TabIndex = 2
        Me.Trade.Text = "Trade"
        '
        'ClearData
        '
        Me.ClearData.BackColor = System.Drawing.Color.FromArgb(CType(128, Byte), CType(128, Byte), CType(255, Byte))
        Me.ClearData.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ClearData.Location = New System.Drawing.Point(8, 256)
        Me.ClearData.Name = "ClearData"
        Me.ClearData.Size = New System.Drawing.Size(128, 40)
        Me.ClearData.TabIndex = 3
        Me.ClearData.Text = "Clear Data"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(432, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 23)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "MA"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(432, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(40, 23)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "StDev"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtMA1
        '
        Me.txtMA1.Location = New System.Drawing.Point(480, 16)
        Me.txtMA1.Name = "txtMA1"
        Me.txtMA1.Size = New System.Drawing.Size(40, 20)
        Me.txtMA1.TabIndex = 6
        Me.txtMA1.Text = "40"
        '
        'txtMA2
        '
        Me.txtMA2.Location = New System.Drawing.Point(480, 48)
        Me.txtMA2.Name = "txtMA2"
        Me.txtMA2.Size = New System.Drawing.Size(40, 20)
        Me.txtMA2.TabIndex = 7
        Me.txtMA2.Text = "40"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rdWEMA)
        Me.GroupBox1.Controls.Add(Me.rdEMA)
        Me.GroupBox1.Controls.Add(Me.rdSMA)
        Me.GroupBox1.Location = New System.Drawing.Point(184, 16)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(112, 120)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "MA Type"
        '
        'rdWEMA
        '
        Me.rdWEMA.Location = New System.Drawing.Point(16, 88)
        Me.rdWEMA.Name = "rdWEMA"
        Me.rdWEMA.Size = New System.Drawing.Size(80, 24)
        Me.rdWEMA.TabIndex = 2
        Me.rdWEMA.Text = "Weighted"
        '
        'rdEMA
        '
        Me.rdEMA.Location = New System.Drawing.Point(16, 56)
        Me.rdEMA.Name = "rdEMA"
        Me.rdEMA.Size = New System.Drawing.Size(88, 24)
        Me.rdEMA.TabIndex = 1
        Me.rdEMA.Text = "Exponential"
        '
        'rdSMA
        '
        Me.rdSMA.Checked = True
        Me.rdSMA.Location = New System.Drawing.Point(16, 24)
        Me.rdSMA.Name = "rdSMA"
        Me.rdSMA.Size = New System.Drawing.Size(72, 24)
        Me.rdSMA.TabIndex = 0
        Me.rdSMA.TabStop = True
        Me.rdSMA.Text = "Simple"
        '
        'txtEndId
        '
        Me.txtEndId.Location = New System.Drawing.Point(480, 112)
        Me.txtEndId.Name = "txtEndId"
        Me.txtEndId.Size = New System.Drawing.Size(40, 20)
        Me.txtEndId.TabIndex = 13
        Me.txtEndId.Text = "200"
        '
        'txtStartID
        '
        Me.txtStartID.Location = New System.Drawing.Point(480, 80)
        Me.txtStartID.Name = "txtStartID"
        Me.txtStartID.Size = New System.Drawing.Size(40, 20)
        Me.txtStartID.TabIndex = 12
        Me.txtStartID.Text = "1"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(432, 112)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(40, 23)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "EndID"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(432, 80)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(40, 23)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "StartID"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtMultB
        '
        Me.txtMultB.Location = New System.Drawing.Point(368, 112)
        Me.txtMultB.Name = "txtMultB"
        Me.txtMultB.Size = New System.Drawing.Size(40, 20)
        Me.txtMultB.TabIndex = 16
        Me.txtMultB.Text = "50"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(320, 112)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(40, 23)
        Me.Label5.TabIndex = 15
        Me.Label5.Text = "Mult B"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtPL
        '
        Me.txtPL.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(128, Byte))
        Me.txtPL.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPL.Location = New System.Drawing.Point(8, 176)
        Me.txtPL.Name = "txtPL"
        Me.txtPL.Size = New System.Drawing.Size(128, 40)
        Me.txtPL.TabIndex = 17
        Me.txtPL.Text = "P/L"
        '
        'txtExport
        '
        Me.txtExport.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(192, Byte), CType(128, Byte))
        Me.txtExport.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtExport.Location = New System.Drawing.Point(8, 416)
        Me.txtExport.Name = "txtExport"
        Me.txtExport.Size = New System.Drawing.Size(128, 40)
        Me.txtExport.TabIndex = 18
        Me.txtExport.Text = "Export"
        '
        'txtMultA
        '
        Me.txtMultA.Location = New System.Drawing.Point(368, 80)
        Me.txtMultA.Name = "txtMultA"
        Me.txtMultA.Size = New System.Drawing.Size(40, 20)
        Me.txtMultA.TabIndex = 127
        Me.txtMultA.Text = "20"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(320, 80)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(40, 23)
        Me.Label7.TabIndex = 126
        Me.Label7.Text = "Mult A"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtTickerA
        '
        Me.txtTickerA.Location = New System.Drawing.Point(368, 16)
        Me.txtTickerA.Name = "txtTickerA"
        Me.txtTickerA.Size = New System.Drawing.Size(40, 20)
        Me.txtTickerA.TabIndex = 129
        Me.txtTickerA.Text = "NQClose"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(304, 16)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(56, 23)
        Me.Label8.TabIndex = 128
        Me.Label8.Text = "Symbol A"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'RunAll
        '
        Me.RunAll.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.RunAll.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RunAll.Location = New System.Drawing.Point(8, 496)
        Me.RunAll.Name = "RunAll"
        Me.RunAll.Size = New System.Drawing.Size(128, 40)
        Me.RunAll.TabIndex = 130
        Me.RunAll.Text = "Run All"
        '
        'groupBox3
        '
        Me.groupBox3.Controls.Add(Me.lbsepl)
        Me.groupBox3.Controls.Add(Me.txtSEPL2)
        Me.groupBox3.Controls.Add(Me.txtLEPL2)
        Me.groupBox3.Controls.Add(Me.lblepl)
        Me.groupBox3.Controls.Add(Me.label49)
        Me.groupBox3.Controls.Add(Me.lbunithigh)
        Me.groupBox3.Controls.Add(Me.label44)
        Me.groupBox3.Controls.Add(Me.label46)
        Me.groupBox3.Controls.Add(Me.lbDayHigh)
        Me.groupBox3.Controls.Add(Me.label40)
        Me.groupBox3.Controls.Add(Me.lbratiodoll)
        Me.groupBox3.Controls.Add(Me.label38)
        Me.groupBox3.Controls.Add(Me.lbnetdailyaverage)
        Me.groupBox3.Controls.Add(Me.label15)
        Me.groupBox3.Controls.Add(Me.label27)
        Me.groupBox3.Controls.Add(Me.label28)
        Me.groupBox3.Controls.Add(Me.ratioD)
        Me.groupBox3.Controls.Add(Me.label30)
        Me.groupBox3.Controls.Add(Me.label31)
        Me.groupBox3.Controls.Add(Me.lbnetdailylow)
        Me.groupBox3.Controls.Add(Me.label33)
        Me.groupBox3.Controls.Add(Me.label34)
        Me.groupBox3.Controls.Add(Me.lblossday)
        Me.groupBox3.Controls.Add(Me.lbwinday)
        Me.groupBox3.Controls.Add(Me.lbnetdailyhigh)
        Me.groupBox3.Controls.Add(Me.lbtotalpl)
        Me.groupBox3.Controls.Add(Me.Label20)
        Me.groupBox3.Controls.Add(Me.Label24)
        Me.groupBox3.Controls.Add(Me.Label22)
        Me.groupBox3.Controls.Add(Me.label17)
        Me.groupBox3.Controls.Add(Me.txtwlratio)
        Me.groupBox3.Controls.Add(Me.label18)
        Me.groupBox3.Controls.Add(Me.label19)
        Me.groupBox3.Controls.Add(Me.lbeql)
        Me.groupBox3.Controls.Add(Me.label25)
        Me.groupBox3.Controls.Add(Me.label26)
        Me.groupBox3.Controls.Add(Me.lbloss)
        Me.groupBox3.Controls.Add(Me.lbwin)
        Me.groupBox3.Controls.Add(Me.lbeqh)
        Me.groupBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.groupBox3.ForeColor = System.Drawing.Color.Blue
        Me.groupBox3.Location = New System.Drawing.Point(184, 240)
        Me.groupBox3.Name = "groupBox3"
        Me.groupBox3.Size = New System.Drawing.Size(528, 248)
        Me.groupBox3.TabIndex = 131
        Me.groupBox3.TabStop = False
        Me.groupBox3.Text = "Trade Results"
        '
        'lbsepl
        '
        Me.lbsepl.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbsepl.Location = New System.Drawing.Point(248, 112)
        Me.lbsepl.Name = "lbsepl"
        Me.lbsepl.Size = New System.Drawing.Size(112, 16)
        Me.lbsepl.TabIndex = 148
        Me.lbsepl.Text = "0"
        '
        'txtSEPL2
        '
        Me.txtSEPL2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSEPL2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtSEPL2.Location = New System.Drawing.Point(168, 112)
        Me.txtSEPL2.Name = "txtSEPL2"
        Me.txtSEPL2.Size = New System.Drawing.Size(72, 16)
        Me.txtSEPL2.TabIndex = 147
        Me.txtSEPL2.Text = "SEPL"
        '
        'txtLEPL2
        '
        Me.txtLEPL2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLEPL2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.txtLEPL2.Location = New System.Drawing.Point(168, 80)
        Me.txtLEPL2.Name = "txtLEPL2"
        Me.txtLEPL2.Size = New System.Drawing.Size(72, 23)
        Me.txtLEPL2.TabIndex = 146
        Me.txtLEPL2.Text = "LEPL"
        '
        'lblepl
        '
        Me.lblepl.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblepl.Location = New System.Drawing.Point(248, 80)
        Me.lblepl.Name = "lblepl"
        Me.lblepl.Size = New System.Drawing.Size(104, 24)
        Me.lblepl.TabIndex = 145
        Me.lblepl.Text = "0"
        '
        'label49
        '
        Me.label49.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label49.ForeColor = System.Drawing.SystemColors.ControlText
        Me.label49.Location = New System.Drawing.Point(272, 136)
        Me.label49.Name = "label49"
        Me.label49.Size = New System.Drawing.Size(48, 16)
        Me.label49.TabIndex = 144
        Me.label49.Text = "Units"
        '
        'lbunithigh
        '
        Me.lbunithigh.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbunithigh.Location = New System.Drawing.Point(280, 160)
        Me.lbunithigh.Name = "lbunithigh"
        Me.lbunithigh.Size = New System.Drawing.Size(32, 16)
        Me.lbunithigh.TabIndex = 142
        Me.lbunithigh.Text = "0"
        '
        'label44
        '
        Me.label44.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label44.ForeColor = System.Drawing.SystemColors.ControlText
        Me.label44.Location = New System.Drawing.Point(168, 136)
        Me.label44.Name = "label44"
        Me.label44.Size = New System.Drawing.Size(96, 16)
        Me.label44.TabIndex = 140
        Me.label44.Text = "Days Per Trade"
        '
        'label46
        '
        Me.label46.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label46.ForeColor = System.Drawing.SystemColors.ControlText
        Me.label46.Location = New System.Drawing.Point(168, 160)
        Me.label46.Name = "label46"
        Me.label46.Size = New System.Drawing.Size(32, 16)
        Me.label46.TabIndex = 138
        Me.label46.Text = "High"
        '
        'lbDayHigh
        '
        Me.lbDayHigh.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbDayHigh.Location = New System.Drawing.Point(224, 160)
        Me.lbDayHigh.Name = "lbDayHigh"
        Me.lbDayHigh.Size = New System.Drawing.Size(32, 16)
        Me.lbDayHigh.TabIndex = 136
        Me.lbDayHigh.Text = "0"
        '
        'label40
        '
        Me.label40.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label40.ForeColor = System.Drawing.SystemColors.ControlText
        Me.label40.Location = New System.Drawing.Point(8, 216)
        Me.label40.Name = "label40"
        Me.label40.Size = New System.Drawing.Size(96, 16)
        Me.label40.TabIndex = 135
        Me.label40.Text = "Ratio Win$/Loss$"
        '
        'lbratiodoll
        '
        Me.lbratiodoll.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbratiodoll.Location = New System.Drawing.Point(112, 216)
        Me.lbratiodoll.Name = "lbratiodoll"
        Me.lbratiodoll.Size = New System.Drawing.Size(40, 16)
        Me.lbratiodoll.TabIndex = 134
        Me.lbratiodoll.Text = "0"
        '
        'label38
        '
        Me.label38.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label38.ForeColor = System.Drawing.SystemColors.ControlText
        Me.label38.Location = New System.Drawing.Point(360, 192)
        Me.label38.Name = "label38"
        Me.label38.Size = New System.Drawing.Size(48, 16)
        Me.label38.TabIndex = 133
        Me.label38.Text = "Average"
        '
        'lbnetdailyaverage
        '
        Me.lbnetdailyaverage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbnetdailyaverage.Location = New System.Drawing.Point(416, 192)
        Me.lbnetdailyaverage.Name = "lbnetdailyaverage"
        Me.lbnetdailyaverage.Size = New System.Drawing.Size(56, 16)
        Me.lbnetdailyaverage.TabIndex = 132
        Me.lbnetdailyaverage.Text = "0"
        '
        'label15
        '
        Me.label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label15.ForeColor = System.Drawing.SystemColors.ControlText
        Me.label15.Location = New System.Drawing.Point(392, 120)
        Me.label15.Name = "label15"
        Me.label15.Size = New System.Drawing.Size(96, 16)
        Me.label15.TabIndex = 131
        Me.label15.Text = "Daily Net Change"
        '
        'label27
        '
        Me.label27.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label27.ForeColor = System.Drawing.SystemColors.ControlText
        Me.label27.Location = New System.Drawing.Point(392, 16)
        Me.label27.Name = "label27"
        Me.label27.Size = New System.Drawing.Size(56, 24)
        Me.label27.TabIndex = 130
        Me.label27.Text = "Daily"
        '
        'label28
        '
        Me.label28.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label28.ForeColor = System.Drawing.SystemColors.ControlText
        Me.label28.Location = New System.Drawing.Point(368, 96)
        Me.label28.Name = "label28"
        Me.label28.Size = New System.Drawing.Size(40, 16)
        Me.label28.TabIndex = 129
        Me.label28.Text = "Ratio"
        '
        'ratioD
        '
        Me.ratioD.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ratioD.Location = New System.Drawing.Point(416, 96)
        Me.ratioD.Name = "ratioD"
        Me.ratioD.Size = New System.Drawing.Size(64, 16)
        Me.ratioD.TabIndex = 128
        Me.ratioD.Text = "0"
        '
        'label30
        '
        Me.label30.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label30.ForeColor = System.Drawing.SystemColors.ControlText
        Me.label30.Location = New System.Drawing.Point(368, 168)
        Me.label30.Name = "label30"
        Me.label30.Size = New System.Drawing.Size(40, 16)
        Me.label30.TabIndex = 127
        Me.label30.Text = " Low"
        '
        'label31
        '
        Me.label31.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label31.ForeColor = System.Drawing.SystemColors.ControlText
        Me.label31.Location = New System.Drawing.Point(376, 144)
        Me.label31.Name = "label31"
        Me.label31.Size = New System.Drawing.Size(32, 16)
        Me.label31.TabIndex = 126
        Me.label31.Text = "High"
        '
        'lbnetdailylow
        '
        Me.lbnetdailylow.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbnetdailylow.Location = New System.Drawing.Point(416, 168)
        Me.lbnetdailylow.Name = "lbnetdailylow"
        Me.lbnetdailylow.Size = New System.Drawing.Size(64, 16)
        Me.lbnetdailylow.TabIndex = 125
        Me.lbnetdailylow.Text = "0"
        '
        'label33
        '
        Me.label33.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label33.ForeColor = System.Drawing.SystemColors.ControlText
        Me.label33.Location = New System.Drawing.Point(368, 72)
        Me.label33.Name = "label33"
        Me.label33.Size = New System.Drawing.Size(40, 16)
        Me.label33.TabIndex = 123
        Me.label33.Text = "Loss"
        '
        'label34
        '
        Me.label34.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label34.ForeColor = System.Drawing.SystemColors.ControlText
        Me.label34.Location = New System.Drawing.Point(368, 48)
        Me.label34.Name = "label34"
        Me.label34.Size = New System.Drawing.Size(40, 16)
        Me.label34.TabIndex = 122
        Me.label34.Text = "Win"
        '
        'lblossday
        '
        Me.lblossday.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblossday.Location = New System.Drawing.Point(416, 72)
        Me.lblossday.Name = "lblossday"
        Me.lblossday.Size = New System.Drawing.Size(64, 16)
        Me.lblossday.TabIndex = 121
        Me.lblossday.Text = "0"
        '
        'lbwinday
        '
        Me.lbwinday.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbwinday.Location = New System.Drawing.Point(416, 48)
        Me.lbwinday.Name = "lbwinday"
        Me.lbwinday.Size = New System.Drawing.Size(64, 16)
        Me.lbwinday.TabIndex = 120
        Me.lbwinday.Text = "0"
        '
        'lbnetdailyhigh
        '
        Me.lbnetdailyhigh.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbnetdailyhigh.Location = New System.Drawing.Point(416, 144)
        Me.lbnetdailyhigh.Name = "lbnetdailyhigh"
        Me.lbnetdailyhigh.Size = New System.Drawing.Size(64, 16)
        Me.lbnetdailyhigh.TabIndex = 124
        Me.lbnetdailyhigh.Text = "0"
        '
        'lbtotalpl
        '
        Me.lbtotalpl.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbtotalpl.Location = New System.Drawing.Point(248, 56)
        Me.lbtotalpl.Name = "lbtotalpl"
        Me.lbtotalpl.Size = New System.Drawing.Size(104, 16)
        Me.lbtotalpl.TabIndex = 119
        Me.lbtotalpl.Text = "0"
        '
        'Label20
        '
        Me.Label20.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label20.Location = New System.Drawing.Point(168, 56)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(72, 16)
        Me.Label20.TabIndex = 118
        Me.Label20.Text = "Total P/L"
        '
        'Label24
        '
        Me.Label24.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label24.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label24.Location = New System.Drawing.Point(56, 144)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(56, 16)
        Me.Label24.TabIndex = 117
        Me.Label24.Text = "Equity"
        '
        'Label22
        '
        Me.Label22.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label22.Location = New System.Drawing.Point(56, 24)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(56, 24)
        Me.Label22.TabIndex = 116
        Me.Label22.Text = "Trades"
        '
        'label17
        '
        Me.label17.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label17.ForeColor = System.Drawing.SystemColors.ControlText
        Me.label17.Location = New System.Drawing.Point(32, 104)
        Me.label17.Name = "label17"
        Me.label17.Size = New System.Drawing.Size(40, 16)
        Me.label17.TabIndex = 115
        Me.label17.Text = "Ratio"
        '
        'txtwlratio
        '
        Me.txtwlratio.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtwlratio.Location = New System.Drawing.Point(80, 104)
        Me.txtwlratio.Name = "txtwlratio"
        Me.txtwlratio.Size = New System.Drawing.Size(64, 16)
        Me.txtwlratio.TabIndex = 114
        Me.txtwlratio.Text = "0"
        '
        'label18
        '
        Me.label18.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label18.ForeColor = System.Drawing.SystemColors.ControlText
        Me.label18.Location = New System.Drawing.Point(32, 192)
        Me.label18.Name = "label18"
        Me.label18.Size = New System.Drawing.Size(40, 16)
        Me.label18.TabIndex = 113
        Me.label18.Text = " Low"
        '
        'label19
        '
        Me.label19.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label19.ForeColor = System.Drawing.SystemColors.ControlText
        Me.label19.Location = New System.Drawing.Point(40, 168)
        Me.label19.Name = "label19"
        Me.label19.Size = New System.Drawing.Size(32, 16)
        Me.label19.TabIndex = 112
        Me.label19.Text = "High"
        '
        'lbeql
        '
        Me.lbeql.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbeql.Location = New System.Drawing.Point(80, 192)
        Me.lbeql.Name = "lbeql"
        Me.lbeql.Size = New System.Drawing.Size(64, 16)
        Me.lbeql.TabIndex = 111
        Me.lbeql.Text = "0"
        '
        'label25
        '
        Me.label25.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label25.ForeColor = System.Drawing.SystemColors.ControlText
        Me.label25.Location = New System.Drawing.Point(32, 80)
        Me.label25.Name = "label25"
        Me.label25.Size = New System.Drawing.Size(40, 16)
        Me.label25.TabIndex = 105
        Me.label25.Text = "Loss"
        '
        'label26
        '
        Me.label26.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label26.ForeColor = System.Drawing.SystemColors.ControlText
        Me.label26.Location = New System.Drawing.Point(32, 56)
        Me.label26.Name = "label26"
        Me.label26.Size = New System.Drawing.Size(40, 16)
        Me.label26.TabIndex = 104
        Me.label26.Text = "Win"
        '
        'lbloss
        '
        Me.lbloss.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbloss.Location = New System.Drawing.Point(80, 80)
        Me.lbloss.Name = "lbloss"
        Me.lbloss.Size = New System.Drawing.Size(64, 16)
        Me.lbloss.TabIndex = 103
        Me.lbloss.Text = "0"
        '
        'lbwin
        '
        Me.lbwin.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbwin.Location = New System.Drawing.Point(80, 56)
        Me.lbwin.Name = "lbwin"
        Me.lbwin.Size = New System.Drawing.Size(64, 16)
        Me.lbwin.TabIndex = 102
        Me.lbwin.Text = "0"
        '
        'lbeqh
        '
        Me.lbeqh.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbeqh.Location = New System.Drawing.Point(80, 168)
        Me.lbeqh.Name = "lbeqh"
        Me.lbeqh.Size = New System.Drawing.Size(64, 16)
        Me.lbeqh.TabIndex = 110
        Me.lbeqh.Text = "0"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(184, 504)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(528, 23)
        Me.ProgressBar1.TabIndex = 134
        '
        'rdShowForm2
        '
        Me.rdShowForm2.Location = New System.Drawing.Point(320, 144)
        Me.rdShowForm2.Name = "rdShowForm2"
        Me.rdShowForm2.TabIndex = 135
        Me.rdShowForm2.Text = "Show Form"
        '
        'txtZExit
        '
        Me.txtZExit.Location = New System.Drawing.Point(480, 176)
        Me.txtZExit.Name = "txtZExit"
        Me.txtZExit.Size = New System.Drawing.Size(40, 20)
        Me.txtZExit.TabIndex = 139
        Me.txtZExit.Text = "0"
        '
        'txtZEntry
        '
        Me.txtZEntry.Location = New System.Drawing.Point(480, 144)
        Me.txtZEntry.Name = "txtZEntry"
        Me.txtZEntry.Size = New System.Drawing.Size(40, 20)
        Me.txtZEntry.TabIndex = 137
        Me.txtZEntry.Text = "1"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(432, 176)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(40, 23)
        Me.Label9.TabIndex = 138
        Me.Label9.Text = "Exit"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(432, 144)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(40, 23)
        Me.Label10.TabIndex = 136
        Me.Label10.Text = "Entry"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtTickerB
        '
        Me.txtTickerB.Location = New System.Drawing.Point(368, 48)
        Me.txtTickerB.Name = "txtTickerB"
        Me.txtTickerB.Size = New System.Drawing.Size(40, 20)
        Me.txtTickerB.TabIndex = 141
        Me.txtTickerB.Text = "ESClose"
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(304, 48)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(56, 23)
        Me.Label11.TabIndex = 140
        Me.Label11.Text = "Symbol B"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.rbDVO1)
        Me.GroupBox2.Controls.Add(Me.rbMVHR)
        Me.GroupBox2.Controls.Add(Me.rbFixed)
        Me.GroupBox2.Location = New System.Drawing.Point(552, 112)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(112, 120)
        Me.GroupBox2.TabIndex = 144
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Ratio Type"
        '
        'rbDVO1
        '
        Me.rbDVO1.Location = New System.Drawing.Point(16, 88)
        Me.rbDVO1.Name = "rbDVO1"
        Me.rbDVO1.Size = New System.Drawing.Size(88, 24)
        Me.rbDVO1.TabIndex = 2
        Me.rbDVO1.Text = "DVO1"
        '
        'rbMVHR
        '
        Me.rbMVHR.Location = New System.Drawing.Point(16, 56)
        Me.rbMVHR.Name = "rbMVHR"
        Me.rbMVHR.Size = New System.Drawing.Size(88, 24)
        Me.rbMVHR.TabIndex = 1
        Me.rbMVHR.Text = "MVHR"
        '
        'rbFixed
        '
        Me.rbFixed.Checked = True
        Me.rbFixed.Location = New System.Drawing.Point(16, 24)
        Me.rbFixed.Name = "rbFixed"
        Me.rbFixed.Size = New System.Drawing.Size(72, 24)
        Me.rbFixed.TabIndex = 0
        Me.rbFixed.TabStop = True
        Me.rbFixed.Text = "Fixed"
        '
        'txtRatioB
        '
        Me.txtRatioB.Location = New System.Drawing.Point(600, 80)
        Me.txtRatioB.Name = "txtRatioB"
        Me.txtRatioB.Size = New System.Drawing.Size(40, 20)
        Me.txtRatioB.TabIndex = 148
        Me.txtRatioB.Text = "1"
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(552, 80)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(40, 23)
        Me.Label14.TabIndex = 147
        Me.Label14.Text = "B"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtRatioA
        '
        Me.txtRatioA.Location = New System.Drawing.Point(600, 48)
        Me.txtRatioA.Name = "txtRatioA"
        Me.txtRatioA.Size = New System.Drawing.Size(40, 20)
        Me.txtRatioA.TabIndex = 146
        Me.txtRatioA.Text = "2"
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(552, 48)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(40, 23)
        Me.Label16.TabIndex = 145
        Me.Label16.Text = "A"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label21
        '
        Me.Label21.Location = New System.Drawing.Point(560, 16)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(72, 23)
        Me.Label21.TabIndex = 151
        Me.Label21.Text = "Fixed Ratios"
        Me.Label21.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.rbPrice)
        Me.GroupBox4.Controls.Add(Me.rbYield)
        Me.GroupBox4.Location = New System.Drawing.Point(184, 144)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(112, 88)
        Me.GroupBox4.TabIndex = 152
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Spread Type"
        '
        'rbPrice
        '
        Me.rbPrice.Checked = True
        Me.rbPrice.Location = New System.Drawing.Point(16, 56)
        Me.rbPrice.Name = "rbPrice"
        Me.rbPrice.Size = New System.Drawing.Size(88, 24)
        Me.rbPrice.TabIndex = 1
        Me.rbPrice.TabStop = True
        Me.rbPrice.Text = "Price"
        '
        'rbYield
        '
        Me.rbYield.Location = New System.Drawing.Point(16, 24)
        Me.rbYield.Name = "rbYield"
        Me.rbYield.Size = New System.Drawing.Size(72, 24)
        Me.rbYield.TabIndex = 0
        Me.rbYield.Text = "Yield"
        '
        'cbTradeLog
        '
        Me.cbTradeLog.Location = New System.Drawing.Point(320, 176)
        Me.cbTradeLog.Name = "cbTradeLog"
        Me.cbTradeLog.TabIndex = 153
        Me.cbTradeLog.Text = "TradeLog"
        '
        'cbCurve
        '
        Me.cbCurve.Location = New System.Drawing.Point(320, 208)
        Me.cbCurve.Name = "cbCurve"
        Me.cbCurve.TabIndex = 154
        Me.cbCurve.Text = "Curve"
        '
        'txtMaxUnits
        '
        Me.txtMaxUnits.Location = New System.Drawing.Point(480, 208)
        Me.txtMaxUnits.Name = "txtMaxUnits"
        Me.txtMaxUnits.Size = New System.Drawing.Size(40, 20)
        Me.txtMaxUnits.TabIndex = 156
        Me.txtMaxUnits.Text = "10"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(432, 208)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(40, 23)
        Me.Label6.TabIndex = 155
        Me.Label6.Text = "Max Units"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Chart
        '
        Me.Chart.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(192, Byte), CType(192, Byte))
        Me.Chart.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Chart.Location = New System.Drawing.Point(8, 336)
        Me.Chart.Name = "Chart"
        Me.Chart.Size = New System.Drawing.Size(128, 40)
        Me.Chart.TabIndex = 157
        Me.Chart.Text = "Chart"
        '
        'CheckBox1
        '
        Me.CheckBox1.Checked = True
        Me.CheckBox1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox1.Location = New System.Drawing.Point(672, 40)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.TabIndex = 158
        Me.CheckBox1.Text = "EndofDay"
        '
        'CheckBox2
        '
        Me.CheckBox2.Location = New System.Drawing.Point(672, 72)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.TabIndex = 159
        Me.CheckBox2.Text = "PriceAdjust"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.GrayText
        Me.ClientSize = New System.Drawing.Size(960, 542)
        Me.Controls.Add(Me.CheckBox2)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Chart)
        Me.Controls.Add(Me.txtMaxUnits)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.cbCurve)
        Me.Controls.Add(Me.cbTradeLog)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.Label21)
        Me.Controls.Add(Me.txtRatioB)
        Me.Controls.Add(Me.txtRatioA)
        Me.Controls.Add(Me.txtTickerB)
        Me.Controls.Add(Me.txtZExit)
        Me.Controls.Add(Me.txtZEntry)
        Me.Controls.Add(Me.txtTickerA)
        Me.Controls.Add(Me.txtMultA)
        Me.Controls.Add(Me.txtMultB)
        Me.Controls.Add(Me.txtEndId)
        Me.Controls.Add(Me.txtStartID)
        Me.Controls.Add(Me.txtMA2)
        Me.Controls.Add(Me.txtMA1)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.rdShowForm2)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.groupBox3)
        Me.Controls.Add(Me.RunAll)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.txtExport)
        Me.Controls.Add(Me.txtPL)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ClearData)
        Me.Controls.Add(Me.Trade)
        Me.Controls.Add(Me.GrabData)
        Me.Name = "Form1"
        Me.Text = "FutureSpreadTester"
        Me.GroupBox1.ResumeLayout(False)
        Me.groupBox3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Dim intLength As Integer
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        myConnection2 = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Temp\ContFutureData.mdb")
        myConnection3 = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Temp\;Extended Properties=""Text;FMT=Delimited""")
        runloop = False
    End Sub

    Private Sub GrabData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GrabData.Click
        If rdShowForm2.Checked Then
            myFormData = New Form2
            myFormData.Text = "Data"
            myFormData.DataGrid2.Visible() = False
            myFormData.DataGrid3.Visible() = False
        End If

        ClearData.Text = "Clear Data"


        Trade.Text = "Trade"
        txtExport.Text = "Export"
        txtPL.Text = "P/L"
        RunAll.Text = "Run ALL "

        myDataset = New DataSet
        ' 
        Try
            myAdapter = New OleDbDataAdapter("SELECT * FROM  NQES1MIN.csv", myConnection3)
            myAdapter.Fill(myDataset, "myData")

        Catch ex As OleDbException
            MessageBox.Show(ex.ToString)
        End Try
        If rdShowForm2.Checked Then
            myFormData.DataGrid1.DataSource = myDataset
            myFormData.DataGrid1.DataMember = "myData"
            myFormData.Show()
        End If



        GrabData.Text = myDataset.Tables("myData").Rows.Count

    End Sub



    Private Sub Trade_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Trade.Click

        If rdShowForm2.Checked Then
            myFormStats = New Form2
            myFormStats.Text = "Stats"
        End If

        Dim myStats As Stats = New Stats
        Dim myTA As New TA.Lib.Core
        intLength = myDataset.Tables("myData").Rows.Count
        If runloop = True Then
            txtStartID.Text = 0
            txtEndId.Text = intLength - 1
        End If
        Dim timeframe As Integer
        Dim timeframe2 As Integer
        Dim max As Integer
        max = Math.Max(Convert.ToInt64(txtMA1.Text), Convert.ToInt64(txtMA2.Text))
        timeframe = Convert.ToInt64(txtEndId.Text) - Convert.ToInt64(txtStartID.Text) - max
        timeframe2 = Convert.ToInt64(txtEndId.Text) - Convert.ToInt64(txtStartID.Text)



        Dim y As Integer
        Dim x As Integer


        Dim closeA As Double() = New Double(intLength) {}
        Dim closeB As Double() = New Double(intLength) {}
        Dim closeC As Double() = New Double(intLength) {}
        Dim yieldA As Double() = New Double(intLength) {}
        Dim yieldB As Double() = New Double(intLength) {}
        Dim yieldC As Double() = New Double(intLength) {}
        Dim Spread As Double() = New Double(intLength) {}
        Dim MA1 As Double() = New Double(timeframe) {}
        Dim MA2 As Double() = New Double(timeframe) {}
        Dim inttime As Integer() = New Integer(intLength) {}
        Dim StDev As Double() = New Double(timeframe) {}
        Dim ZSCORE As Double() = New Double(timeframe) {}
        Dim DayUp As Double() = New Double(timeframe) {}
        Dim DayDown As Double() = New Double(timeframe) {}


        For x = 0 To intLength - 1
            y = x - 1
            'set dataset values for ease of use in passing values into function

            closeA(x) = myDataset.Tables(0).Rows(x).Item("" & txtTickerA.Text)
            'yieldA(x) = myDataset.Tables(0).Rows(x).Item("" & txtTickerA.Text & "yld")
            closeB(x) = myDataset.Tables(0).Rows(x).Item("" & txtTickerB.Text)
            ' yieldB(x) = myDataset.Tables(0).Rows(x).Item("" & txtTickerB.Text & "yld")
            inttime(x) = myDataset.Tables(0).Rows(x).Item("ESTime")

            If rbYield.Checked Then
                Spread(x) = yieldB(x) - yieldA(x)
            ElseIf CheckBox2.Checked Then
                Spread(x) = (txtMultA.Text * txtRatioA.Text * closeA(x)) - (txtMultB.Text * txtRatioB.Text * closeB(x))
            Else
                Spread(x) = closeA(x) - closeB(x)
            End If
            'loop to do all stat calculations from stat public functions
        Next

        Dim n, m, j As Integer

        myTA.MA(max, timeframe2, Spread, Convert.ToInt64(txtMA1.Text), MA_type(), n, m, MA1)
        myTA.STDDEV(max, timeframe2, Spread, txtMA2.Text, 1, n, m, StDev)


        Check_Stats()
        j = 1
        Dim k As Integer = 0
        For x = 0 To timeframe
            k = x + max
            myDataset2.Tables(0).Rows(x).Item("DateID") = k
            ' myDataset2.Tables(0).Rows(x).Item("A") = myDataset.Tables(0).Rows(k).Item("NQDate")
            myDataset2.Tables(0).Rows(x).Item("B") = closeA(k)
            myDataset2.Tables(0).Rows(x).Item("C") = yieldA(k)
            myDataset2.Tables(0).Rows(x).Item("D") = closeB(k)
            myDataset2.Tables(0).Rows(x).Item("E") = yieldB(k)
            myDataset2.Tables(0).Rows(x).Item("F") = Spread(k)
            myDataset2.Tables(0).Rows(x).Item("G") = MA1(x)
            myDataset2.Tables(0).Rows(x).Item("H") = StDev(x)

            If (Math.Abs(Spread(k) - MA1(x)) > 0.0001) Then
                myDataset2.Tables(0).Rows(x).Item("I") = (Spread(k) - MA1(x)) / StDev(x)
                ZSCORE(x) = (Spread(k) - MA1(x)) / StDev(x)
            Else
                ZSCORE(x) = ZSCORE(x) - 1

            End If

        Next

        Grab_Template()

        Trade_Zscore(ZSCORE, Spread, inttime)

        Trade_Units(closeA, closeB)
        Trade.Text = "Trade Done"





    End Sub

    Private Sub ClearData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearData.Click

        Dim strCom As String
        strCom = New String("Delete * from EQCURVE where ID > -1")
        myConnection2.Open()
        Dim command2 As New OleDbCommand(strCom, myConnection2)
        command2.ExecuteNonQuery()
        myConnection2.Close()

        strCom = New String("Delete * from Results where ID > -1")
        myConnection2.Open()
        command2 = New OleDbCommand(strCom, myConnection2)
        command2.ExecuteNonQuery()
        myConnection2.Close()

        ClearData.Text = "Clear Data Done"
     

    End Sub
    Public Function MA_type()
        Dim a As Integer
        If (rdSMA.Checked) Then
            a = 0
        ElseIf (rdEMA.Checked) Then
            a = 1
        ElseIf (rdWEMA.Checked) Then
            a = 2
        End If

        Return a

    End Function



    Public Function Check_Stats()


        myDataset2 = New DataSet
        myAdapter2 = New OleDbDataAdapter("SELECT * FROM tblCheck.csv", myConnection3)

        myConnection2.Open()

        myAdapter2.Fill(myDataset2, "myData2")

        If rdShowForm2.Checked Then
            myFormStats.DataGrid2.Visible() = True
            myFormStats.DataGrid2.DataSource = myDataset2
            myFormStats.DataGrid2.DataMember = "myData2"
            myFormStats.Show()
        End If
        myConnection2.Close()
    End Function

    Private Sub DataGrid2_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs)

    End Sub

    Private Sub txtExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtExport.Click

        exportcount = exportcount + 1
        Dim strCom As String
        strCom = New String("INSERT INTO Results (ID,Product,StartTime,TimeFrame,Win,Loss,WLRatio,EquityHigh,EquityLow,TotalPL,RatioDollar,DayHigh,UnitHigh,DailyWin,DailyLoss,DailyRatio,DailyChngHigh,DailyChngLow,DailyChngAve) VALUES (")
        Dim strCom5 As String
        strCom5 = New String(txtTickerA.Text + "-" + txtTickerB.Text)


        Dim strCom2 As String
        strCom2 = New String("" & exportcount & ",'" & strCom5 & "' ," & txtStartID.Text & "," & txtEndId.Text & "," & lbwin.Text & "," & lbloss.Text & "," & txtwlratio.Text & "," & lbeqh.Text & "," & lbeql.Text & "," & lbtotalpl.Text & "," & lbratiodoll.Text & "," & lbDayHigh.Text & "," & lbunithigh.Text & "," & lbwinday.Text & "," & lblossday.Text & "," & ratioD.Text & "," & lbnetdailyhigh.Text & "," & lbnetdailylow.Text & "," & lbnetdailyaverage.Text & ")")
        strCom = strCom.Concat(strCom, strCom2)
        myConnection2.Open()
        Dim command2 As New OleDbCommand(strCom, myConnection2)
        command2.ExecuteNonQuery()
        myConnection2.Close()


        'Dim I As Integer
        'intLength = myDataset.Tables("myData").Rows.Count
        'Dim timeframe As Integer
        'Dim timeframe2 As Integer
        'Dim max As Integer
        'max = Math.Max(Convert.ToInt64(txtMA1.Text), Convert.ToInt64(txtMA2.Text))
        'timeframe = Convert.ToInt64(txtEndId.Text) - Convert.ToInt64(txtStartID.Text) - max
        'timeframe2 = Convert.ToInt64(txtEndId.Text) - Convert.ToInt64(txtStartID.Text)
        'myConnection2.Open()
        'For I = max + 1 To timeframe - 1
        '    exportcountcurve = exportcountcurve + 1
        '    Dim x As Double = Convert.ToDouble(myDataset3.Tables(0).Rows(I).Item("TotalPL"))
        '    Dim y As Date = myDataset3.Tables(0).Rows(I).Item("Date")
        '    Dim strCom3 As String
        '    strCom3 = New String("INSERT INTO EQCURVE (ID,[Date],Ticker,EQCURVE) VALUES (")
        '    Dim strCom4 As String
        '    strCom4 = New String("" & exportcountcurve & ",# " & y & " #,'" & strCom5 & "'," & x & ")")
        '    strCom3 = strCom3.Concat(strCom3, strCom4)
        '    Try
        '        Dim command4 As New OleDbCommand(strCom3, myConnection2)
        '        command4.ExecuteNonQuery()
        '    Catch ex As OleDbException
        '        MessageBox.Show(ex.ToString)
        '    End Try

        'Next
        'myConnection2.Close()


        If (cbTradeLog.Checked) Then
            Export_Trade_Log()
        End If
        If (cbCurve.Checked) Then
            Export_Curve()
        End If
        txtExport.Text = "Export Done"
    End Sub

    Private Sub txtPL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPL.Click
        Dim I As Integer
        intLength = myDataset.Tables("myData").Rows.Count
        Dim timeframe As Integer
        Dim timeframe2 As Integer
        Dim max As Integer
        max = Math.Max(Convert.ToInt64(txtMA1.Text), Convert.ToInt64(txtMA2.Text))
        timeframe = Convert.ToInt64(txtEndId.Text) - Convert.ToInt64(txtStartID.Text) - max
        timeframe2 = Convert.ToInt64(txtEndId.Text) - Convert.ToInt64(txtStartID.Text)
        Dim x As Integer
        CleanValues()




        For I = max + 1 To timeframe - 1
            Ratio(I)
            LEPL(I)
            SEPL(I)
            LETradePL(I)
            SETradePL(I)
            TotalPL(I)
            NetChngD(I)
            TotalEQHigh(I)
            TotalEQLow(I)
            Winners(I)
            Losers(I)
            UnitHigh(I)
            DayHigh(I)
            WinLossDays(I)
            EQHighLowDays(I)
            DollarRatio(I)

        Next

        WLRatio()
        WLDaysRatio()
        FinalTotalPL(I - 1)




        txtPL.Text = "P/L Done"
    End Sub

    Public Function Grab_Template()
        If rdShowForm2.Checked Then
            myFormTrades = New Form2
            myFormTrades.Text = "Trades"
            myFormTrades.DataGrid3.Visible() = True
        End If

        myDataset3 = New DataSet
        myAdapter3 = New OleDbDataAdapter("SELECT * FROM Template.csv", myConnection3)

        myConnection3.Open()

        myAdapter3.Fill(myDataset3, "myData3")
        If rdShowForm2.Checked Then
            myFormTrades.DataGrid3.DataSource = myDataset3
            myFormTrades.DataGrid3.DataMember = "myData3"
            myFormTrades.Show()
        End If
        myConnection3.Close()


        'Dim workRow As DataRow
        Dim i As Integer

        For i = 0 To Convert.ToInt64(txtEndId.Text) - 1
            ' workRow = myDataset3.Tables("myData3").NewRow()
            ' myDataset3.Tables("myData3").Rows.Add(workRow)
            myDataset3.Tables(0).Rows(i).Item("ID") = i
            myDataset3.Tables(0).Rows(i).Item("Date") = myDataset.Tables(0).Rows(i).Item("NQDate")
            myDataset3.Tables(0).Rows(i).Item("PriceA") = myDataset.Tables(0).Rows(i).Item("" & txtTickerA.Text)
            myDataset3.Tables(0).Rows(i).Item("Time") = myDataset.Tables(0).Rows(i).Item("ESTime")
            'myDataset3.Tables(0).Rows(i).Item("YieldA") = myDataset.Tables(0).Rows(i).Item("" & txtTickerA.Text & "yld")
            myDataset3.Tables(0).Rows(i).Item("PriceB") = myDataset.Tables(0).Rows(i).Item("" & txtTickerB.Text)
            ' myDataset3.Tables(0).Rows(i).Item("YieldB") = myDataset.Tables(0).Rows(i).Item("" & txtTickerB.Text & "yld")
       
        Next





    End Function
    Public Function Trade_Zscore(ByRef z As Double(), ByRef Spread As Double(), ByRef inttime As Integer())



        'Dim workRow As DataRow
        Dim I As Integer
        intLength = myDataset.Tables("myData").Rows.Count
        Dim timeframe As Integer
        Dim timeframe2 As Integer
        Dim max As Integer
        max = Math.Max(Convert.ToInt64(txtMA1.Text), Convert.ToInt64(txtMA2.Text))
        timeframe = Convert.ToInt64(txtEndId.Text) - Convert.ToInt64(txtStartID.Text) - max
        timeframe2 = Convert.ToInt64(txtEndId.Text) - Convert.ToInt64(txtStartID.Text)
        Dim x As Integer


        For I = max + 1 To timeframe - 1
            ' workRow = myDataset3.Tables("myData3").NewRow()
            ' myDataset3.Tables("myData3").Rows.Add(workRow)
            x = I - max
            If z(x) < Convert.ToDouble(txtZEntry.Text) * -1 Then
                myDataset3.Tables(0).Rows(I).Item("LE") = Spread(I)
            Else
                myDataset3.Tables(0).Rows(I).Item("LE") = 0
            End If

            If z(x) > Convert.ToDouble(txtZExit.Text) * -1 And z(x - 1) < Convert.ToDouble(txtZExit.Text) * -1 Then
                myDataset3.Tables(0).Rows(I).Item("LX") = Spread(I)
            Else
                myDataset3.Tables(0).Rows(I).Item("LX") = 0
            End If


            If z(x) > Convert.ToDouble(txtZEntry.Text) Then
                myDataset3.Tables(0).Rows(I).Item("SE") = Spread(I)
            Else
                myDataset3.Tables(0).Rows(I).Item("SE") = 0
            End If

            If z(x) < Convert.ToDouble(txtZExit.Text) * -1 And z(x - 1) > Convert.ToDouble(txtZExit.Text) Then
                myDataset3.Tables(0).Rows(I).Item("SX") = Spread(I)
            Else
                myDataset3.Tables(0).Rows(I).Item("SX") = 0
            End If
            If inttime(x) = 1515 And CheckBox1.Checked Then
                myDataset3.Tables(0).Rows(I).Item("SX") = Spread(I)
                myDataset3.Tables(0).Rows(I).Item("LX") = Spread(I)
            End If


        Next

    End Function

    Public Function Trade_Units(ByRef closeA As Double(), ByRef closeB As Double())



        'Dim workRow As DataRow
        Dim I As Integer
        intLength = myDataset.Tables("myData").Rows.Count
        Dim timeframe As Integer
        Dim timeframe2 As Integer
        Dim max As Integer
        max = Math.Max(Convert.ToInt64(txtMA1.Text), Convert.ToInt64(txtMA2.Text))
        timeframe = Convert.ToInt64(txtEndId.Text) - Convert.ToInt64(txtStartID.Text) - max
        timeframe2 = Convert.ToInt64(txtEndId.Text) - Convert.ToInt64(txtStartID.Text)
        Dim x As Integer




        On Error Resume Next

        For I = max + 1 To timeframe - 1


            If myDataset3.Tables(0).Rows(I).Item("LE") <> 0 And myDataset3.Tables(0).Rows(I).Item("LEUNITS") < txtMaxUnits.Text Then
                myDataset3.Tables(0).Rows(I + 1).Item("LEUNITS") = myDataset3.Tables(0).Rows(I).Item("LEUNITS") + 1
            Else
                myDataset3.Tables(0).Rows(I + 1).Item("LEUNITS") = myDataset3.Tables(0).Rows(I).Item("LEUNITS")
            End If
            If myDataset3.Tables(0).Rows(I).Item("LX") <> 0 Then
                myDataset3.Tables(0).Rows(I + 1).Item("LEUNITS") = 0
            End If


            If myDataset3.Tables(0).Rows(I).Item("SE") <> 0 And myDataset3.Tables(0).Rows(I).Item("SEUNITS") < txtMaxUnits.Text Then
                myDataset3.Tables(0).Rows(I + 1).Item("SEUNITS") = myDataset3.Tables(0).Rows(I).Item("SEUNITS") + 1
            Else
                myDataset3.Tables(0).Rows(I + 1).Item("SEUNITS") = myDataset3.Tables(0).Rows(I).Item("SEUNITS")
            End If
            If myDataset3.Tables(0).Rows(I).Item("SX") <> 0 Then
                myDataset3.Tables(0).Rows(I + 1).Item("SEUNITS") = 0
            End If

            myDataset3.Tables(0).Rows(I).Item("NetChngA") = closeA(I) - closeA(I - 1)
            myDataset3.Tables(0).Rows(I).Item("NetChngB") = closeB(I) - closeB(I - 1)

        Next

    End Function
    Public Function TotalEQHigh(ByRef i As Integer)
        If myDataset3.Tables(0).Rows(i).Item("TotalPL") > myDataset3.Tables(0).Rows(i - 1).Item("TotalPL") Then
            If lbeqh.Text = "" Then
                lbeqh.Text = 0
            Else
                If myDataset3.Tables(0).Rows(i).Item("TotalPL") > Convert.ToDouble(lbeqh.Text) Then
                    lbeqh.Text = Format(myDataset3.Tables(0).Rows(i).Item("TotalPL"), "#.##")
                End If
            End If
        End If
    End Function
    Public Function TotalEQLow(ByRef i As Integer)
        If myDataset3.Tables(0).Rows(i).Item("TotalPL") < myDataset3.Tables(0).Rows(i - 1).Item("TotalPL") Then
            If myDataset3.Tables(0).Rows(i).Item("TotalPL") < Convert.ToDouble(lbeql.Text) Then
                lbeql.Text = Format(myDataset3.Tables(0).Rows(i).Item("TotalPL"), "#.##")
            End If
        End If
    End Function

    Public Function LETradePL(ByRef i As Integer)
        If myDataset3.Tables(0).Rows(i).Item("LEUNITS") > 0 Then
            myDataset3.Tables(0).Rows(i).Item("LETradePL") = _
          (myDataset3.Tables(0).Rows(i).Item("NetChngA") * Convert.ToDouble(txtMultA.Text) * myDataset3.Tables(0).Rows(i).Item("RatioA") + _
             myDataset3.Tables(0).Rows(i).Item("NetChngB") * -1 * Convert.ToDouble(txtMultB.Text) * myDataset3.Tables(0).Rows(i).Item("RatioB")) _
   * myDataset3.Tables(0).Rows(i).Item("LEUNITS") + myDataset3.Tables(0).Rows(i - 1).Item("LETradePL")

        Else
            myDataset3.Tables(0).Rows(i).Item("LETradePL") = 0
        End If

    End Function
    Public Function SETradePL(ByRef i As Integer)
        If myDataset3.Tables(0).Rows(i).Item("SEUNITS") > 0 Then
            myDataset3.Tables(0).Rows(i).Item("SETradePL") = _
          (myDataset3.Tables(0).Rows(i).Item("NetChngA") * -1 * Convert.ToDouble(txtMultA.Text) * myDataset3.Tables(0).Rows(i).Item("RatioA") + _
             myDataset3.Tables(0).Rows(i).Item("NetChngB") * Convert.ToDouble(txtMultB.Text) * myDataset3.Tables(0).Rows(i).Item("RatioB")) _
              * myDataset3.Tables(0).Rows(i).Item("SEUNITS") + myDataset3.Tables(0).Rows(i - 1).Item("SETradePL")
        Else
            myDataset3.Tables(0).Rows(i).Item("SETradePL") = 0
        End If

    End Function

    Public Function LEPL(ByRef i As Integer)
        If myDataset3.Tables(0).Rows(i).Item("LEUNITS") = 0 Then
            myDataset3.Tables(0).Rows(i).Item("LEPL") = myDataset3.Tables(0).Rows(i - 1).Item("LEPL")
        Else
            myDataset3.Tables(0).Rows(i).Item("LEPL") = _
          (myDataset3.Tables(0).Rows(i).Item("NetChngA") * Convert.ToDouble(txtMultA.Text) * myDataset3.Tables(0).Rows(i).Item("RatioA") + _
             myDataset3.Tables(0).Rows(i).Item("NetChngB") * -1 * Convert.ToDouble(txtMultB.Text) * myDataset3.Tables(0).Rows(i).Item("RatioB")) _
              * myDataset3.Tables(0).Rows(i).Item("LEUNITS") + myDataset3.Tables(0).Rows(i - 1).Item("LEPL")
        End If


    End Function
    Public Function SEPL(ByRef i As Integer)
        If myDataset3.Tables(0).Rows(i).Item("SEUNITS") = 0 Then
            myDataset3.Tables(0).Rows(i).Item("SEPL") = myDataset3.Tables(0).Rows(i - 1).Item("SEPL")
        Else
            myDataset3.Tables(0).Rows(i).Item("SEPL") = _
          (myDataset3.Tables(0).Rows(i).Item("NetChngA") * -1 * Convert.ToDouble(txtMultA.Text) * myDataset3.Tables(0).Rows(i).Item("RatioA") + _
             myDataset3.Tables(0).Rows(i).Item("NetChngB") * Convert.ToDouble(txtMultB.Text) * myDataset3.Tables(0).Rows(i).Item("RatioB")) _
              * myDataset3.Tables(0).Rows(i).Item("SEUNITS") + myDataset3.Tables(0).Rows(i - 1).Item("SEPL")
        End If
    End Function
    Public Function TotalPL(ByRef i As Integer)

        myDataset3.Tables(0).Rows(i).Item("TotalPL") = myDataset3.Tables(0).Rows(i).Item("LEPL") + myDataset3.Tables(0).Rows(i).Item("SEPL")

    End Function
    Public Function NetChngD(ByRef i As Integer)

        myDataset3.Tables(0).Rows(i).Item("NetChngPL") = myDataset3.Tables(0).Rows(i).Item("TotalPL") - myDataset3.Tables(0).Rows(i - 1).Item("TotalPL")

    End Function
    Public Function DollarRatio(ByRef i As Integer)


        If myDataset3.Tables(0).Rows(i).Item("LETradePL") = 0 Then
            If myDataset3.Tables(0).Rows(i - 1).Item("LETradePL") < 0 Then
                lossdollar = lossdollar + myDataset3.Tables(0).Rows(i - 1).Item("LETradePL")
            End If
        End If


        If myDataset3.Tables(0).Rows(i).Item("SETradePL") = 0 Then
            If myDataset3.Tables(0).Rows(i - 1).Item("SETradePL") < 0 Then
                lossdollar = lossdollar + myDataset3.Tables(0).Rows(i - 1).Item("SETradePL")
            End If
        End If

        If myDataset3.Tables(0).Rows(i).Item("LETradePL") = 0 Then
            If myDataset3.Tables(0).Rows(i - 1).Item("LETradePL") > 0 Then
                windollar = windollar + myDataset3.Tables(0).Rows(i - 1).Item("LETradePL")
            End If
        End If


        If myDataset3.Tables(0).Rows(i).Item("SETradePL") = 0 Then
            If myDataset3.Tables(0).Rows(i - 1).Item("SETradePL") > 0 Then
                windollar = windollar + myDataset3.Tables(0).Rows(i - 1).Item("SETradePL")
            End If
        End If
        Dim dbltest As Double = 0

        If lossdollar = 0 Or windollar = 0 Then
            lbratiodoll.Text = 0
        Else
            lbratiodoll.Text = Format(windollar / lossdollar * -1, "#.##")
            dbltest = windollar / lossdollar * -1
        End If
        If dbltest < 0.01 Then
            lbratiodoll.Text = 0
        End If
    End Function
    Public Function Winners(ByRef i As Integer)


        If myDataset3.Tables(0).Rows(i).Item("LEUnits") = 0 And myDataset3.Tables(0).Rows(i - 1).Item("LEUnits") <> 0 And myDataset3.Tables(0).Rows(i - 1).Item("LETradePL") > 0 Then

            lbwin.Text = Convert.ToInt64(lbwin.Text) + 1

        End If



        If myDataset3.Tables(0).Rows(i).Item("SEUnits") = 0 And myDataset3.Tables(0).Rows(i - 1).Item("SEUnits") <> 0 And myDataset3.Tables(0).Rows(i - 1).Item("SETradePL") > 0 Then

            lbwin.Text = Convert.ToInt64(lbwin.Text) + 1

        End If


    End Function


    Public Function Losers(ByRef i As Integer)

        If myDataset3.Tables(0).Rows(i).Item("LEUnits") = 0 And myDataset3.Tables(0).Rows(i - 1).Item("LEUnits") <> 0 And myDataset3.Tables(0).Rows(i - 1).Item("LETradePL") < 0 Then

            lbloss.Text = Convert.ToInt64(lbloss.Text) + 1

        End If



        If myDataset3.Tables(0).Rows(i).Item("SEUnits") = 0 And myDataset3.Tables(0).Rows(i - 1).Item("SEUnits") <> 0 And myDataset3.Tables(0).Rows(i - 1).Item("SETradePL") < 0 Then

            lbloss.Text = Convert.ToInt64(lbloss.Text) + 1

        End If


    End Function
    Public Function WinLossDays(ByRef i As Integer)

        If myDataset3.Tables(0).Rows(i).Item("NetChngPL") > 0 Then
            lbwinday.Text = Convert.ToInt64(lbwinday.Text) + 1
        End If
        If myDataset3.Tables(0).Rows(i).Item("NetChngPL") < 0 Then
            lblossday.Text = Convert.ToInt64(lblossday.Text) + 1
        End If

    End Function
    Public Function EQHighLowDays(ByRef i As Integer)

        If myDataset3.Tables(0).Rows(i).Item("NetChngPL") > Convert.ToDouble(lbnetdailyhigh.Text) Then
            lbnetdailyhigh.Text = Format(myDataset3.Tables(0).Rows(i).Item("NetChngPL"), "#.##")
        End If
        If myDataset3.Tables(0).Rows(i).Item("NetChngPL") < Convert.ToDouble(lbnetdailylow.Text) Then
            lbnetdailylow.Text = Format(myDataset3.Tables(0).Rows(i).Item("NetChngPL"), "#.##")
        End If

    End Function


    Public Function WLRatio()

        If Convert.ToInt64(lbloss.Text) = 0 Or Convert.ToInt64(lbwin.Text) = 0 Then
            txtwlratio.Text = 0
        Else
            txtwlratio.Text = Format(Convert.ToInt64(lbwin.Text) / Convert.ToInt64(lbloss.Text), "#.##")
        End If

    End Function
    Public Function WLDaysRatio()

        If Convert.ToInt64(lblossday.Text) = 0 Or Convert.ToInt64(lbwinday.Text) = 0 Then
            ratioD.Text = 0
        Else
            ratioD.Text = Format(Convert.ToInt64(lbwinday.Text) / Convert.ToInt64(lblossday.Text), "#.##")
        End If

    End Function

    Public Function UnitHigh(ByRef i As Integer)

        If myDataset3.Tables(0).Rows(i).Item("LEUNITS") > 0 Then
            If myDataset3.Tables(0).Rows(i).Item("LEUNITS") > Convert.ToInt64(lbunithigh.Text) Then
                lbunithigh.Text = myDataset3.Tables(0).Rows(i).Item("LEUNITS")
            End If
        End If
        If myDataset3.Tables(0).Rows(i).Item("SEUNITS") > 0 Then
            If myDataset3.Tables(0).Rows(i).Item("SEUNITS") > Convert.ToInt64(lbunithigh.Text) Then
                lbunithigh.Text = myDataset3.Tables(0).Rows(i).Item("SEUNITS")
            End If
        End If

    End Function
    Public Function FinalTotalPL(ByRef i As Integer)
        If myDataset3.Tables(0).Rows(i).Item("TotalPL") <> 0 Then
            lbtotalpl.Text = Format(myDataset3.Tables(0).Rows(i).Item("TotalPL"), "#.##")
        End If
        If myDataset3.Tables(0).Rows(i).Item("LEPL") <> 0 Then
            lblepl.Text = Format(myDataset3.Tables(0).Rows(i).Item("LEPL"), "#.##")
        End If
        If myDataset3.Tables(0).Rows(i).Item("SEPL") <> 0 Then
            lbsepl.Text = Format(myDataset3.Tables(0).Rows(i).Item("SEPL"), "#.##")
        End If


    End Function
    Public Function DayHigh(ByRef i As Integer)

        If myDataset3.Tables(0).Rows(i).Item("LEUNITS") > 0 Then
            daycountl = daycountl + 1
        End If

        If myDataset3.Tables(0).Rows(i).Item("LEUNITS") = 0 Then
            daycountl = 0
        End If

        If myDataset3.Tables(0).Rows(i).Item("SEUNITS") > 0 Then
            daycounts = daycounts + 1
        End If

        If myDataset3.Tables(0).Rows(i).Item("LEUNITS") = 0 Then
            daycounts = 0
        End If

        If daycountl > Convert.ToInt64(lbDayHigh.Text) Then
            lbDayHigh.Text = daycountl
        End If

        If daycounts > Convert.ToInt64(lbDayHigh.Text) Then
            lbDayHigh.Text = daycounts
        End If


    End Function
    Public Function DayAve(ByRef i As Integer)
        Dim count As Integer = 0
        Dim timeframe As Integer
        Dim timeframe2 As Integer
        Dim max As Integer
        max = Math.Max(Convert.ToInt64(txtMA1.Text), Convert.ToInt64(txtMA2.Text))
        timeframe = Convert.ToInt64(txtEndId.Text) - Convert.ToInt64(txtStartID.Text) - max
        timeframe2 = Convert.ToInt64(txtEndId.Text) - Convert.ToInt64(txtStartID.Text)
        For i = max + 1 To timeframe - 1


        Next


    End Function
    Public Function Ratio(ByRef i As Integer)

        If (rbFixed.Checked) Then
            myDataset3.Tables(0).Rows(i).Item("RatioA") = Convert.ToDouble(txtRatioA.Text)
            myDataset3.Tables(0).Rows(i).Item("RatioB") = Convert.ToDouble(txtRatioB.Text)
        End If


    End Function


    Public Function CleanValues()
        daycountl = 0
        daycounts = 0
        lossdollar = 0
        windollar = 0
        lbwin.Text = 0
        lbloss.Text = 0
        txtwlratio.Text = 0
        lbeqh.Text = 0.0
        lbeql.Text = 0.0
        lbratiodoll.Text = 0

        lbtotalpl.Text = 0
        lblepl.Text = 0
        lbsepl.Text = 0
        lbDayHigh.Text = 0

        lbunithigh.Text = 0

        lbwinday.Text = 0
        lblossday.Text = 0
        ratioD.Text = 0
        lbnetdailyhigh.Text = 0
        lbnetdailylow.Text = 0
        lbnetdailyaverage.Text = 0



    End Function

    Public Function SharpeCalc(ByRef i As Integer)

        Dim sharpecount, sortinocount As Integer
        Dim avsharpesum, avsharpe, stdsharpe, sharpe, stdsortino, sortino, avesortino, avesortinosum As Double

        'sharpe
        If myDataset3.Tables(0).Rows(i).Item("NetChngPL") <> 0 Then
            sharpecount = sharpecount + 1
            avsharpesum = avsharpesum + myDataset3.Tables(0).Rows(i).Item("NetChngPL")
            If myDataset3.Tables(0).Rows(i).Item("NetChngPL") < 0 Then
                sortinocount = sortinocount + 1
                avesortinosum = avesortinosum + myDataset3.Tables(0).Rows(i).Item("NetChngPL")
            End If

        End If



        avsharpe = (avsharpesum / sharpecount)
        avesortino = (avesortinosum / sortinocount)
        For i = 0 To 500
            If myDataset3.Tables(0).Rows(i).Item("NetChngPL") <> 0 Then

                stdsharpe = (myDataset3.Tables(0).Rows(i).Item("NetChngPL") - avsharpe) ^ 2 / (sharpecount - 1) + stdsharpe

                If myDataset3.Tables(0).Rows(i).Item("NetChngPL") < 0 Then
                    stdsortino = (myDataset3.Tables(0).Rows(i).Item("NetChngPL") - avesortino) ^ 2 / (sortinocount - 1) + stdsortino
                End If
            End If
        Next


        sortino = avsharpe / stdsortino ^ 0.5
        sharpe = avsharpe / stdsharpe ^ 0.5
        '  lbsharpe.Text = Format((sharpe * 252 ^ 0.5), "0.00")
        'lbsortino.Text = Format((sortino * 252 ^ 0.5), "0.00")


    End Function



    Private Sub groupBox3_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles groupBox3.Enter

    End Sub

    Private Sub RunAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RunAll.Click
        runloop = True
        ClearData.PerformClick()
        myDataset4 = New DataSet
        myAdapter3 = New OleDbDataAdapter("SELECT * FROM  Tickers.csv", myConnection3)
        myAdapter3.Fill(myDataset4, "myData4")
        Dim tickerLength As Integer
        tickerLength = myDataset4.Tables("myData4").Rows.Count
        Dim x As Integer
        'myConnection3.Open()
        ProgressBar1.Maximum = tickerLength
        For x = 0 To tickerLength - 1
            ProgressBar1.Value = x
            txtTickerA.Text = myDataset4.Tables(0).Rows(x).Item("Symbol")
            txtMultA.Text = myDataset4.Tables(0).Rows(x).Item("Multiplier")
            GrabData.PerformClick()
            Trade.PerformClick()
            txtPL.PerformClick()
            txtExport.PerformClick()

        Next
        ProgressBar1.Visible = False
        runloop = False
        RunAll.Text = "Run ALL Done "
    End Sub
    Public Function MA_type_str()
        Dim a As String
        If (rdSMA.Checked) Then
            a = "Simple"
        ElseIf (rdEMA.Checked) Then
            a = "Exponential"
        ElseIf (rdWEMA.Checked) Then
            a = "Weighted"
        End If

        Return a

    End Function
    Public Function Spread_type_str()
        Dim a As String
        If (rbYield.Checked) Then
            a = "Yield"
        Else
            a = "Price"
        End If

        Return a

    End Function
    Public Function Ratio_type_str()
        Dim a As String
        If (rbFixed.Checked) Then
            a = "Fixed"
        ElseIf (rbMVHR.Checked) Then
            a = "MVHR"
        ElseIf (rbDVO1.Checked) Then
            a = "DVO1"
        End If

        Return a

    End Function
    Private Function Export_Trade_Log()



        Dim intColumn As Integer
        Dim intRows As Integer
        Dim x As Integer
        Dim y As Integer
        intColumn = myDataset3.Tables("myData3").Columns.Count
        intRows = myDataset3.Tables("myData3").Rows.Count
        Dim matype As String
        matype = MA_type_str()
        Dim timeframe As Integer
        Dim timeframe2 As Integer
        Dim max As Integer
        max = Math.Max(Convert.ToInt64(txtMA1.Text), Convert.ToInt64(txtMA2.Text))
        timeframe = Convert.ToInt64(txtEndId.Text) - Convert.ToInt64(txtStartID.Text) - max
        timeframe2 = Convert.ToInt64(txtEndId.Text) - Convert.ToInt64(txtStartID.Text)



        Dim sw = New StreamWriter("c:\temp\FutureSpreadsTradeLog\" & txtTickerA.Text + "_" + txtTickerB.Text + "_MA" + txtMA1.Text + matype + "_StDev" + txtMA2.Text + "Log.csv")

        sw.Write("Spread = " + Spread_type_str())
        sw.Write(",")
        sw.Write("Ratio = " + Ratio_type_str())
        sw.Write(",")
        sw.Write("EntryZ = " + txtZEntry.Text)
        sw.Write(",")
        sw.Write("ExitZ = " + txtZExit.Text)
        sw.Write(",")
        sw.Write("MaxUnits = " + txtMaxUnits.Text)
        sw.Write(",")

        sw.Write(sw.NewLine)
        sw.Write(sw.NewLine)

        For x = 0 To intColumn - 1
            sw.Write(myDataset3.Tables("myData3").Columns(x))
            If x < intColumn - 1 Then
                sw.Write(",")
            End If
        Next x


        For y = max + 1 To timeframe - 1
            For x = 0 To intColumn - 1
                sw.Write(myDataset3.Tables("myData3").Rows(y).Item(x))
                If x < intColumn - 1 Then
                    sw.Write(",")
                End If
            Next x
            sw.Write(sw.NewLine)
        Next y
        sw.Close()


    End Function

    Private Function Export_Curve()

        Dim intColumn As Integer
        Dim intRows As Integer
        Dim x, y As Integer


        Dim timeframe As Integer
        Dim timeframe2 As Integer
        Dim max As Integer
        max = Math.Max(Convert.ToInt64(txtMA1.Text), Convert.ToInt64(txtMA2.Text))
        timeframe = Convert.ToInt64(txtEndId.Text) - Convert.ToInt64(txtStartID.Text) - max
        timeframe2 = Convert.ToInt64(txtEndId.Text) - Convert.ToInt64(txtStartID.Text)

        Dim matype As String
        matype = MA_type_str()


        Dim sw = New StreamWriter("c:\temp\FutureSpreadsTradeLog\" & txtTickerA.Text + "_" + txtTickerB.Text + "_MA" + txtMA1.Text + matype + "_StDev" + txtMA2.Text + "Curve.csv")


        sw.Write("Spread = " + Spread_type_str())
        sw.Write(",")
        sw.Write("Ratio = " + Ratio_type_str())
        sw.Write(",")
        sw.Write("EntryZ = " + txtZEntry.Text)
        sw.Write(",")
        sw.Write("ExitZ = " + txtZExit.Text)
        sw.Write(",")
        sw.Write("MaxUnits = " + txtMaxUnits.Text)
        sw.Write(",")

        sw.Write(sw.NewLine)
        sw.Write(sw.NewLine)

        sw.Write("Date")
        sw.Write(",")
        sw.Write("EQCurve")
        sw.Write(",")
        sw.Write(sw.NewLine)

        For y = max + 1 To timeframe - 1

            sw.Write(myDataset3.Tables("myData3").Rows(y).Item("Date"))
            sw.Write(",")
            sw.Write(myDataset3.Tables("myData3").Rows(y).Item("TotalPL"))
            sw.Write(",")

            sw.Write(sw.NewLine)
        Next y
        sw.Close()


    End Function


    Private Sub Chart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Chart.Click

        Dim y As Integer
        Dim timeframe As Integer
        Dim timeframe2 As Integer
        Dim max As Integer
        max = Math.Max(Convert.ToInt64(txtMA1.Text), Convert.ToInt64(txtMA2.Text))
        timeframe = Convert.ToInt64(txtEndId.Text) - Convert.ToInt64(txtStartID.Text) - max
        timeframe2 = Convert.ToInt64(txtEndId.Text) - Convert.ToInt64(txtStartID.Text)
        myFormChart = New Form3

        Dim values(timeframe - 3) As Double
        Dim values2(timeframe - 3) As Double
        For y = 0 To timeframe - 3
            values(y) = myDataset3.Tables("myData3").Rows(y + Convert.ToInt64(txtStartID.Text)).Item("TotalPL")
            values2(y) = myDataset3.Tables("myData3").Rows(y + Convert.ToInt64(txtStartID.Text)).Item("LEPL")
        Next y

        myFormChart.Chart1.DataSource = values

        'myFormChart.Chart1.Value(0, 0) = 1
        'myFormChart.Chart1.Value(0, 1) = 2
        'myFormChart.Chart1.Value(1, 0) = 3
        'myFormChart.Chart1.Value(1, 1) = 4

        myFormChart.Chart1.Gallery = SoftwareFX.ChartFX.Lite.Gallery.Lines
        myFormChart.Chart1.Grid = SoftwareFX.ChartFX.Lite.ChartGrid.Horz


        myFormChart.Show()

    End Sub
End Class
