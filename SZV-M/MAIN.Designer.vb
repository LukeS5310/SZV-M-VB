<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MAIN
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
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
        Me.components = New System.ComponentModel.Container()
        Me.BTN_START = New System.Windows.Forms.Button()
        Me.PB_PROGRESS = New System.Windows.Forms.ProgressBar()
        Me.LBL_STATE = New System.Windows.Forms.Label()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.TS_MEMUSE = New System.Windows.Forms.ToolStripStatusLabel()
        Me.LBL_DBG = New System.Windows.Forms.ToolStripStatusLabel()
        Me.Mem_Timer = New System.Windows.Forms.Timer(Me.components)
        Me.BW_IndexBases = New System.ComponentModel.BackgroundWorker()
        Me.BW_OPERATION = New System.ComponentModel.BackgroundWorker()
        Me.BTN_START_PO = New System.Windows.Forms.Button()
        Me.CB_MONTH = New System.Windows.Forms.ComboBox()
        Me.BTN_SETTINGS = New System.Windows.Forms.Button()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'BTN_START
        '
        Me.BTN_START.Location = New System.Drawing.Point(13, 13)
        Me.BTN_START.Name = "BTN_START"
        Me.BTN_START.Size = New System.Drawing.Size(224, 23)
        Me.BTN_START.TabIndex = 0
        Me.BTN_START.Text = "ИМПОРТ ДО"
        Me.BTN_START.UseVisualStyleBackColor = True
        '
        'PB_PROGRESS
        '
        Me.PB_PROGRESS.Location = New System.Drawing.Point(12, 73)
        Me.PB_PROGRESS.Name = "PB_PROGRESS"
        Me.PB_PROGRESS.Size = New System.Drawing.Size(337, 23)
        Me.PB_PROGRESS.TabIndex = 1
        '
        'LBL_STATE
        '
        Me.LBL_STATE.Location = New System.Drawing.Point(12, 103)
        Me.LBL_STATE.Name = "LBL_STATE"
        Me.LBL_STATE.Size = New System.Drawing.Size(337, 25)
        Me.LBL_STATE.TabIndex = 2
        Me.LBL_STATE.Text = "Label1"
        Me.LBL_STATE.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TS_MEMUSE, Me.LBL_DBG})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 136)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(362, 22)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 3
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'TS_MEMUSE
        '
        Me.TS_MEMUSE.Name = "TS_MEMUSE"
        Me.TS_MEMUSE.Size = New System.Drawing.Size(76, 17)
        Me.TS_MEMUSE.Text = "MEM_USAGE"
        '
        'LBL_DBG
        '
        Me.LBL_DBG.Name = "LBL_DBG"
        Me.LBL_DBG.Size = New System.Drawing.Size(13, 17)
        Me.LBL_DBG.Text = "0"
        '
        'Mem_Timer
        '
        Me.Mem_Timer.Enabled = True
        Me.Mem_Timer.Interval = 1000
        '
        'BW_IndexBases
        '
        Me.BW_IndexBases.WorkerReportsProgress = True
        '
        'BW_OPERATION
        '
        Me.BW_OPERATION.WorkerReportsProgress = True
        '
        'BTN_START_PO
        '
        Me.BTN_START_PO.Location = New System.Drawing.Point(12, 44)
        Me.BTN_START_PO.Name = "BTN_START_PO"
        Me.BTN_START_PO.Size = New System.Drawing.Size(338, 23)
        Me.BTN_START_PO.TabIndex = 4
        Me.BTN_START_PO.Text = "ИМПОРТ ""ПОСЛЕ"" И РАССЧЕТ"
        Me.BTN_START_PO.UseVisualStyleBackColor = True
        '
        'CB_MONTH
        '
        Me.CB_MONTH.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_MONTH.FormattingEnabled = True
        Me.CB_MONTH.Items.AddRange(New Object() {"Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"})
        Me.CB_MONTH.Location = New System.Drawing.Point(280, 136)
        Me.CB_MONTH.MaxDropDownItems = 12
        Me.CB_MONTH.Name = "CB_MONTH"
        Me.CB_MONTH.Size = New System.Drawing.Size(70, 21)
        Me.CB_MONTH.TabIndex = 5
        '
        'BTN_SETTINGS
        '
        Me.BTN_SETTINGS.Location = New System.Drawing.Point(243, 13)
        Me.BTN_SETTINGS.Name = "BTN_SETTINGS"
        Me.BTN_SETTINGS.Size = New System.Drawing.Size(107, 23)
        Me.BTN_SETTINGS.TabIndex = 6
        Me.BTN_SETTINGS.Text = "Настройки..."
        Me.BTN_SETTINGS.UseVisualStyleBackColor = True
        '
        'MAIN
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(362, 158)
        Me.Controls.Add(Me.BTN_SETTINGS)
        Me.Controls.Add(Me.CB_MONTH)
        Me.Controls.Add(Me.BTN_START_PO)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.LBL_STATE)
        Me.Controls.Add(Me.PB_PROGRESS)
        Me.Controls.Add(Me.BTN_START)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "MAIN"
        Me.Text = "СЗВ-М 1.0 RC2"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents BTN_START As Button
    Friend WithEvents PB_PROGRESS As ProgressBar
    Friend WithEvents LBL_STATE As Label
    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents TS_MEMUSE As ToolStripStatusLabel
    Friend WithEvents Mem_Timer As Timer
    Friend WithEvents BW_IndexBases As System.ComponentModel.BackgroundWorker
    Friend WithEvents MainData As DataSet
    Friend WithEvents INDEX4 As DataTable
    Friend WithEvents RA As DataColumn
    Friend WithEvents DataColumn1 As DataColumn
    Friend WithEvents FA As DataColumn
    Friend WithEvents DataColumn2 As DataColumn
    Friend WithEvents OT As DataColumn
    Friend WithEvents NPERS As DataColumn
    Friend WithEvents RDAT As DataColumn
    Friend WithEvents NAZ As DataColumn
    Friend WithEvents FW_PENS_DO As DataColumn
    Friend WithEvents SC_PENS_DO As DataColumn
    Friend WithEvents FW_DOPL_DO As DataColumn
    Friend WithEvents SC_DOPL_DO As DataColumn
    Friend WithEvents FW_PENS_PO As DataColumn
    Friend WithEvents SC_PENS_PO As DataColumn
    Friend WithEvents FW_DOPL_PO As DataColumn
    Friend WithEvents SC_DOPL_PO As DataColumn
    Friend WithEvents RABOT_DO As DataColumn
    Friend WithEvents RABOT_PO As DataColumn
    Friend WithEvents P_DELTA As DataColumn
    Friend WithEvents D_DELTA As DataColumn
    Friend WithEvents MAN_ID As DataColumn
    Friend WithEvents BW_OPERATION As System.ComponentModel.BackgroundWorker
    Friend WithEvents BTN_START_PO As Button
    Friend WithEvents LBL_DBG As ToolStripStatusLabel
    Friend WithEvents CB_MONTH As ComboBox
    Friend WithEvents BTN_SETTINGS As Button
End Class
