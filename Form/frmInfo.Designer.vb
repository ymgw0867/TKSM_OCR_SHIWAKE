<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmInfo
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
        Me.PrintDocument = New System.Drawing.Printing.PrintDocument()
        Me.PrintDImage = New System.Drawing.Printing.PrintDocument()
        Me.RasterImageViewer1 = New Leadtools.WinForms.RasterImageViewer()
        Me.cmdPlus = New System.Windows.Forms.Button()
        Me.cmdMinus = New System.Windows.Forms.Button()
        Me.btnOk = New System.Windows.Forms.Button()
        Me.cmdImgPrn = New System.Windows.Forms.Button()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.btnDltDen = New System.Windows.Forms.Button()
        Me.cmdChudan = New System.Windows.Forms.Button()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.lblNowDen = New System.Windows.Forms.Label()
        Me.tabData = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.TabPage4 = New System.Windows.Forms.TabPage()
        Me.TabPage5 = New System.Windows.Forms.TabPage()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.DataGridView3 = New System.Windows.Forms.DataGridView()
        Me.DataGridView4 = New System.Windows.Forms.DataGridView()
        Me.DataGridView5 = New System.Windows.Forms.DataGridView()
        Me.DataGridView6 = New System.Windows.Forms.DataGridView()
        Me.DataGridView7 = New System.Windows.Forms.DataGridView()
        Me.DataGridView8 = New System.Windows.Forms.DataGridView()
        Me.tabData.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.TabPage4.SuspendLayout()
        Me.TabPage5.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView8, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PrintDocument
        '
        '
        'PrintDImage
        '
        '
        'RasterImageViewer1
        '
        Me.RasterImageViewer1.Location = New System.Drawing.Point(12, 12)
        Me.RasterImageViewer1.Name = "RasterImageViewer1"
        Me.RasterImageViewer1.Size = New System.Drawing.Size(515, 347)
        Me.RasterImageViewer1.TabIndex = 0
        '
        'cmdPlus
        '
        Me.cmdPlus.Location = New System.Drawing.Point(533, 12)
        Me.cmdPlus.Name = "cmdPlus"
        Me.cmdPlus.Size = New System.Drawing.Size(28, 26)
        Me.cmdPlus.TabIndex = 1
        Me.cmdPlus.Text = "Button1"
        Me.cmdPlus.UseVisualStyleBackColor = True
        '
        'cmdMinus
        '
        Me.cmdMinus.Location = New System.Drawing.Point(533, 44)
        Me.cmdMinus.Name = "cmdMinus"
        Me.cmdMinus.Size = New System.Drawing.Size(28, 26)
        Me.cmdMinus.TabIndex = 2
        Me.cmdMinus.Text = "Button2"
        Me.cmdMinus.UseVisualStyleBackColor = True
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(607, 14)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(91, 24)
        Me.btnOk.TabIndex = 3
        Me.btnOk.Text = "エラーチェック(&E)"
        Me.btnOk.UseVisualStyleBackColor = True
        '
        'cmdImgPrn
        '
        Me.cmdImgPrn.Location = New System.Drawing.Point(695, 14)
        Me.cmdImgPrn.Name = "cmdImgPrn"
        Me.cmdImgPrn.Size = New System.Drawing.Size(91, 24)
        Me.cmdImgPrn.TabIndex = 4
        Me.cmdImgPrn.Text = "画像印刷(&I)"
        Me.cmdImgPrn.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(782, 14)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(91, 23)
        Me.btnPrint.TabIndex = 5
        Me.btnPrint.Text = "伝票印刷(&P)"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'btnDltDen
        '
        Me.btnDltDen.Location = New System.Drawing.Point(607, 34)
        Me.btnDltDen.Name = "btnDltDen"
        Me.btnDltDen.Size = New System.Drawing.Size(91, 24)
        Me.btnDltDen.TabIndex = 6
        Me.btnDltDen.Text = "伝票削除(&D)"
        Me.btnDltDen.UseVisualStyleBackColor = True
        '
        'cmdChudan
        '
        Me.cmdChudan.Location = New System.Drawing.Point(695, 34)
        Me.cmdChudan.Name = "cmdChudan"
        Me.cmdChudan.Size = New System.Drawing.Size(91, 24)
        Me.cmdChudan.TabIndex = 7
        Me.cmdChudan.Text = "作業中断(&C)"
        Me.cmdChudan.UseVisualStyleBackColor = True
        '
        'cmdExit
        '
        Me.cmdExit.Location = New System.Drawing.Point(782, 34)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(91, 24)
        Me.cmdExit.TabIndex = 8
        Me.cmdExit.Text = "終了(&X)"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'lblNowDen
        '
        Me.lblNowDen.AutoSize = True
        Me.lblNowDen.Location = New System.Drawing.Point(894, 25)
        Me.lblNowDen.Name = "lblNowDen"
        Me.lblNowDen.Size = New System.Drawing.Size(29, 12)
        Me.lblNowDen.TabIndex = 9
        Me.lblNowDen.Text = "0000"
        '
        'tabData
        '
        Me.tabData.Controls.Add(Me.TabPage1)
        Me.tabData.Controls.Add(Me.TabPage2)
        Me.tabData.Controls.Add(Me.TabPage3)
        Me.tabData.Controls.Add(Me.TabPage4)
        Me.tabData.Controls.Add(Me.TabPage5)
        Me.tabData.Location = New System.Drawing.Point(533, 76)
        Me.tabData.Name = "tabData"
        Me.tabData.SelectedIndex = 0
        Me.tabData.Size = New System.Drawing.Size(442, 283)
        Me.tabData.TabIndex = 10
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.DataGridView4)
        Me.TabPage1.Location = New System.Drawing.Point(4, 21)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(434, 258)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "NG"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.DataGridView3)
        Me.TabPage2.Controls.Add(Me.DataGridView2)
        Me.TabPage2.Location = New System.Drawing.Point(4, 21)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(434, 258)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "科目"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.DataGridView7)
        Me.TabPage3.Controls.Add(Me.DataGridView6)
        Me.TabPage3.Controls.Add(Me.DataGridView5)
        Me.TabPage3.Location = New System.Drawing.Point(4, 21)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(434, 258)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "税・部門"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.DataGridView8)
        Me.TabPage4.Location = New System.Drawing.Point(4, 21)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage4.Size = New System.Drawing.Size(434, 258)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "摘要"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.DataGridView1)
        Me.TabPage5.Location = New System.Drawing.Point(4, 21)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage5.Size = New System.Drawing.Size(434, 258)
        Me.TabPage5.TabIndex = 4
        Me.TabPage5.Text = "会社"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(0, 0)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowTemplate.Height = 21
        Me.DataGridView1.Size = New System.Drawing.Size(431, 255)
        Me.DataGridView1.TabIndex = 0
        '
        'DataGridView2
        '
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Location = New System.Drawing.Point(5, 4)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.RowTemplate.Height = 21
        Me.DataGridView2.Size = New System.Drawing.Size(206, 251)
        Me.DataGridView2.TabIndex = 0
        '
        'DataGridView3
        '
        Me.DataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView3.Location = New System.Drawing.Point(217, 4)
        Me.DataGridView3.Name = "DataGridView3"
        Me.DataGridView3.RowTemplate.Height = 21
        Me.DataGridView3.Size = New System.Drawing.Size(214, 251)
        Me.DataGridView3.TabIndex = 1
        '
        'DataGridView4
        '
        Me.DataGridView4.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView4.Location = New System.Drawing.Point(4, 3)
        Me.DataGridView4.Name = "DataGridView4"
        Me.DataGridView4.RowTemplate.Height = 21
        Me.DataGridView4.Size = New System.Drawing.Size(427, 252)
        Me.DataGridView4.TabIndex = 0
        '
        'DataGridView5
        '
        Me.DataGridView5.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView5.Location = New System.Drawing.Point(1, 1)
        Me.DataGridView5.Name = "DataGridView5"
        Me.DataGridView5.RowTemplate.Height = 21
        Me.DataGridView5.Size = New System.Drawing.Size(202, 151)
        Me.DataGridView5.TabIndex = 0
        '
        'DataGridView6
        '
        Me.DataGridView6.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView6.Location = New System.Drawing.Point(2, 155)
        Me.DataGridView6.Name = "DataGridView6"
        Me.DataGridView6.RowTemplate.Height = 21
        Me.DataGridView6.Size = New System.Drawing.Size(200, 100)
        Me.DataGridView6.TabIndex = 1
        '
        'DataGridView7
        '
        Me.DataGridView7.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView7.Location = New System.Drawing.Point(208, 1)
        Me.DataGridView7.Name = "DataGridView7"
        Me.DataGridView7.RowTemplate.Height = 21
        Me.DataGridView7.Size = New System.Drawing.Size(223, 254)
        Me.DataGridView7.TabIndex = 2
        '
        'DataGridView8
        '
        Me.DataGridView8.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView8.Location = New System.Drawing.Point(1, 1)
        Me.DataGridView8.Name = "DataGridView8"
        Me.DataGridView8.RowTemplate.Height = 21
        Me.DataGridView8.Size = New System.Drawing.Size(430, 254)
        Me.DataGridView8.TabIndex = 0
        '
        'frmInfo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(987, 706)
        Me.Controls.Add(Me.tabData)
        Me.Controls.Add(Me.lblNowDen)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdChudan)
        Me.Controls.Add(Me.btnDltDen)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.cmdImgPrn)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.cmdMinus)
        Me.Controls.Add(Me.cmdPlus)
        Me.Controls.Add(Me.RasterImageViewer1)
        Me.Name = "frmInfo"
        Me.Text = "frmInfo"
        Me.tabData.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage4.ResumeLayout(False)
        Me.TabPage5.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView8, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents PrintDocument As System.Drawing.Printing.PrintDocument
    Friend WithEvents PrintDImage As System.Drawing.Printing.PrintDocument
    Friend WithEvents RasterImageViewer1 As Leadtools.WinForms.RasterImageViewer
    Friend WithEvents cmdPlus As System.Windows.Forms.Button
    Friend WithEvents cmdMinus As System.Windows.Forms.Button
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents cmdImgPrn As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnDltDen As System.Windows.Forms.Button
    Friend WithEvents cmdChudan As System.Windows.Forms.Button
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents lblNowDen As System.Windows.Forms.Label
    Friend WithEvents tabData As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents DataGridView4 As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents DataGridView3 As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridView2 As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents DataGridView7 As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridView6 As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridView5 As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents DataGridView8 As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
End Class
