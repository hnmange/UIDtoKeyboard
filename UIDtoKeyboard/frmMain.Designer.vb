﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.cbxReaderList = New System.Windows.Forms.ComboBox()
        Me.btnRefreshReader = New System.Windows.Forms.Button()
        Me.btnStartMonitor = New System.Windows.Forms.Button()
        Me.txtInputSpace = New System.Windows.Forms.TextBox()
        Me.btnStopMonitor = New System.Windows.Forms.Button()
        Me.lblInst1 = New System.Windows.Forms.Label()
        Me.lblInst2 = New System.Windows.Forms.Label()
        Me.txtReadingMode = New System.Windows.Forms.TextBox()
        Me.lblReadingMode = New System.Windows.Forms.Label()
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cbxReaderList
        '
        Me.cbxReaderList.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.cbxReaderList.FormattingEnabled = True
        Me.cbxReaderList.Location = New System.Drawing.Point(25, 22)
        Me.cbxReaderList.Name = "cbxReaderList"
        Me.cbxReaderList.Size = New System.Drawing.Size(400, 26)
        Me.cbxReaderList.TabIndex = 0
        '
        'btnRefreshReader
        '
        Me.btnRefreshReader.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.btnRefreshReader.Location = New System.Drawing.Point(431, 22)
        Me.btnRefreshReader.Name = "btnRefreshReader"
        Me.btnRefreshReader.Size = New System.Drawing.Size(87, 27)
        Me.btnRefreshReader.TabIndex = 2
        Me.btnRefreshReader.Text = "Refresh"
        Me.btnRefreshReader.UseVisualStyleBackColor = True
        '
        'btnStartMonitor
        '
        Me.btnStartMonitor.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.btnStartMonitor.Location = New System.Drawing.Point(22, 91)
        Me.btnStartMonitor.Name = "btnStartMonitor"
        Me.btnStartMonitor.Size = New System.Drawing.Size(496, 32)
        Me.btnStartMonitor.TabIndex = 3
        Me.btnStartMonitor.Text = "Start Monitor"
        Me.btnStartMonitor.UseVisualStyleBackColor = True
        '
        'txtInputSpace
        '
        Me.txtInputSpace.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.txtInputSpace.Location = New System.Drawing.Point(22, 141)
        Me.txtInputSpace.Multiline = True
        Me.txtInputSpace.Name = "txtInputSpace"
        Me.txtInputSpace.Size = New System.Drawing.Size(496, 384)
        Me.txtInputSpace.TabIndex = 4
        '
        'btnStopMonitor
        '
        Me.btnStopMonitor.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.btnStopMonitor.Location = New System.Drawing.Point(431, 55)
        Me.btnStopMonitor.Name = "btnStopMonitor"
        Me.btnStopMonitor.Size = New System.Drawing.Size(87, 27)
        Me.btnStopMonitor.TabIndex = 5
        Me.btnStopMonitor.Text = "Stop"
        Me.btnStopMonitor.UseVisualStyleBackColor = True
        '
        'lblInst1
        '
        Me.lblInst1.AutoSize = True
        Me.lblInst1.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblInst1.Location = New System.Drawing.Point(551, 30)
        Me.lblInst1.Name = "lblInst1"
        Me.lblInst1.Size = New System.Drawing.Size(104, 18)
        Me.lblInst1.TabIndex = 6
        Me.lblInst1.Text = "Reading Mode"
        '
        'lblInst2
        '
        Me.lblInst2.AutoSize = True
        Me.lblInst2.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblInst2.Location = New System.Drawing.Point(567, 59)
        Me.lblInst2.Name = "lblInst2"
        Me.lblInst2.Size = New System.Drawing.Size(217, 72)
        Me.lblInst2.TabIndex = 7
        Me.lblInst2.Text = "1- Card UID (4 Byte)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "2- Card UID (4 Byte + Reverse)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "3- Card UID (7 Byte)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "4- Ca" &
    "rd UID (7 Byte + Reverse) "
        '
        'txtReadingMode
        '
        Me.txtReadingMode.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.txtReadingMode.Location = New System.Drawing.Point(136, 56)
        Me.txtReadingMode.Name = "txtReadingMode"
        Me.txtReadingMode.Size = New System.Drawing.Size(289, 24)
        Me.txtReadingMode.TabIndex = 8
        '
        'lblReadingMode
        '
        Me.lblReadingMode.AutoSize = True
        Me.lblReadingMode.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblReadingMode.Location = New System.Drawing.Point(22, 59)
        Me.lblReadingMode.Name = "lblReadingMode"
        Me.lblReadingMode.Size = New System.Drawing.Size(108, 18)
        Me.lblReadingMode.TabIndex = 9
        Me.lblReadingMode.Text = "Reading Mode:"
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info
        Me.NotifyIcon1.BalloonTipText = "UIDtoKeyboard running"
        Me.NotifyIcon1.BalloonTipTitle = "UIDtoKeyboard"
        Me.NotifyIcon1.ContextMenuStrip = Me.ContextMenuStrip1
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "UIDtoKeyboard"
        Me.NotifyIcon1.Visible = True
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem1, Me.ToolStripMenuItem2})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(127, 48)
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.CheckOnClick = True
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(126, 22)
        Me.ToolStripMenuItem1.Text = "Auto start"
        '
        'ToolStripMenuItem2
        '
        Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
        Me.ToolStripMenuItem2.Size = New System.Drawing.Size(126, 22)
        Me.ToolStripMenuItem2.Text = "Exit"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(810, 487)
        Me.Controls.Add(Me.lblReadingMode)
        Me.Controls.Add(Me.txtReadingMode)
        Me.Controls.Add(Me.lblInst2)
        Me.Controls.Add(Me.lblInst1)
        Me.Controls.Add(Me.btnStopMonitor)
        Me.Controls.Add(Me.txtInputSpace)
        Me.Controls.Add(Me.btnStartMonitor)
        Me.Controls.Add(Me.btnRefreshReader)
        Me.Controls.Add(Me.cbxReaderList)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmMain"
        Me.Opacity = 0R
        Me.ShowInTaskbar = False
        Me.Text = "UIDtoKeyboard"
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cbxReaderList As ComboBox
    Friend WithEvents btnRefreshReader As Button
    Friend WithEvents btnStartMonitor As Button
    Friend WithEvents txtInputSpace As TextBox
    Friend WithEvents btnStopMonitor As Button
    Friend WithEvents lblInst1 As Label
    Friend WithEvents lblInst2 As Label
    Friend WithEvents txtReadingMode As TextBox
    Friend WithEvents lblReadingMode As Label
    Friend WithEvents NotifyIcon1 As NotifyIcon
    Friend WithEvents ContextMenuStrip1 As ContextMenuStrip
    Friend WithEvents ToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem2 As ToolStripMenuItem
End Class
