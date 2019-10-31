<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFilSelect
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
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

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFilSelect))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.CList1 = New System.Windows.Forms.CheckedListBox()
        Me.cmdTrue = New System.Windows.Forms.Button()
        Me.cmdFalse = New System.Windows.Forms.Button()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.CmdEnd = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.CList1)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 5)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(202, 135)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "OCRデータ読込日付・時間・件数"
        '
        'CList1
        '
        Me.CList1.FormattingEnabled = True
        Me.CList1.Location = New System.Drawing.Point(6, 13)
        Me.CList1.Name = "CList1"
        Me.CList1.Size = New System.Drawing.Size(190, 116)
        Me.CList1.TabIndex = 0
        '
        'cmdTrue
        '
        Me.cmdTrue.Location = New System.Drawing.Point(226, 10)
        Me.cmdTrue.Name = "cmdTrue"
        Me.cmdTrue.Size = New System.Drawing.Size(98, 29)
        Me.cmdTrue.TabIndex = 1
        Me.cmdTrue.Text = "全伝票選択(&A)"
        Me.cmdTrue.UseVisualStyleBackColor = True
        '
        'cmdFalse
        '
        Me.cmdFalse.Location = New System.Drawing.Point(225, 45)
        Me.cmdFalse.Name = "cmdFalse"
        Me.cmdFalse.Size = New System.Drawing.Size(99, 29)
        Me.cmdFalse.TabIndex = 2
        Me.cmdFalse.Text = "全伝票取消(&S)"
        Me.cmdFalse.UseVisualStyleBackColor = True
        '
        'cmdOK
        '
        Me.cmdOK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOK.Location = New System.Drawing.Point(225, 80)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(99, 27)
        Me.cmdOK.TabIndex = 3
        Me.cmdOK.Text = "確定(&K)"
        Me.cmdOK.UseVisualStyleBackColor = True
        '
        'CmdEnd
        '
        Me.CmdEnd.ForeColor = System.Drawing.Color.Black
        Me.CmdEnd.Location = New System.Drawing.Point(226, 113)
        Me.CmdEnd.Name = "CmdEnd"
        Me.CmdEnd.Size = New System.Drawing.Size(98, 27)
        Me.CmdEnd.TabIndex = 4
        Me.CmdEnd.Text = "終了(&X)"
        Me.CmdEnd.UseVisualStyleBackColor = True
        '
        'frmFilSelect
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(335, 152)
        Me.Controls.Add(Me.CmdEnd)
        Me.Controls.Add(Me.cmdTrue)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdFalse)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmFilSelect"
        Me.Text = "作業中断伝票の選択"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdTrue As System.Windows.Forms.Button
    Friend WithEvents cmdFalse As System.Windows.Forms.Button
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents CmdEnd As System.Windows.Forms.Button
    Friend WithEvents CList1 As System.Windows.Forms.CheckedListBox

End Class
