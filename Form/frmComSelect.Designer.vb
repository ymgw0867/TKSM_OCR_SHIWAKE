<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmComSelect
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmComSelect))
        Me.lblmsg = New System.Windows.Forms.Label()
        Me.btnOk = New System.Windows.Forms.Button()
        Me.gvComp = New System.Windows.Forms.DataGridView()
        Me.No = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.kisyu = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Kessan = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Comp = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.gvComp, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblmsg
        '
        Me.lblmsg.AutoSize = True
        Me.lblmsg.Location = New System.Drawing.Point(12, 9)
        Me.lblmsg.Name = "lblmsg"
        Me.lblmsg.Size = New System.Drawing.Size(121, 12)
        Me.lblmsg.TabIndex = 0
        Me.lblmsg.Text = "会社を選択してください。"
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(446, 168)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(95, 23)
        Me.btnOk.TabIndex = 2
        Me.btnOk.Text = "OK"
        Me.btnOk.UseVisualStyleBackColor = True
        '
        'gvComp
        '
        Me.gvComp.AllowUserToResizeColumns = False
        Me.gvComp.AllowUserToResizeRows = False
        Me.gvComp.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.gvComp.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.No, Me.kisyu, Me.Kessan, Me.Comp})
        Me.gvComp.Location = New System.Drawing.Point(14, 24)
        Me.gvComp.MultiSelect = False
        Me.gvComp.Name = "gvComp"
        Me.gvComp.ReadOnly = True
        Me.gvComp.RowHeadersVisible = False
        Me.gvComp.RowTemplate.Height = 21
        Me.gvComp.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.gvComp.Size = New System.Drawing.Size(527, 127)
        Me.gvComp.TabIndex = 3
        '
        'No
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.No.DefaultCellStyle = DataGridViewCellStyle1
        Me.No.HeaderText = "   No."
        Me.No.Name = "No"
        Me.No.ReadOnly = True
        Me.No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.No.Width = 50
        '
        'kisyu
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.kisyu.DefaultCellStyle = DataGridViewCellStyle2
        Me.kisyu.HeaderText = "　　　　期首"
        Me.kisyu.Name = "kisyu"
        Me.kisyu.ReadOnly = True
        Me.kisyu.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'Kessan
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Kessan.DefaultCellStyle = DataGridViewCellStyle3
        Me.Kessan.HeaderText = "　　　決算期"
        Me.Kessan.Name = "Kessan"
        Me.Kessan.ReadOnly = True
        Me.Kessan.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'Comp
        '
        Me.Comp.HeaderText = "会社名"
        Me.Comp.Name = "Comp"
        Me.Comp.ReadOnly = True
        Me.Comp.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Comp.Width = 285
        '
        'frmComSelect
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(552, 200)
        Me.Controls.Add(Me.gvComp)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.lblmsg)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmComSelect"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "会社選択"
        CType(Me.gvComp, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblmsg As System.Windows.Forms.Label
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents gvComp As System.Windows.Forms.DataGridView
    Friend WithEvents No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents kisyu As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Kessan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Comp As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
