namespace CSV_import_export
{
	partial class frmImport
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.label1 = new System.Windows.Forms.Label();
            this.txtFileToImport = new System.Windows.Forms.TextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.dataGridView_preView = new System.Windows.Forms.DataGridView();
            this.chkFirstRowColumnNames = new System.Windows.Forms.CheckBox();
            this.gpbSeparator = new System.Windows.Forms.GroupBox();
            this.txtSeparatorOtherChar = new System.Windows.Forms.TextBox();
            this.rdbSeparatorOther = new System.Windows.Forms.RadioButton();
            this.rdbTab = new System.Windows.Forms.RadioButton();
            this.rdbSemicolon = new System.Windows.Forms.RadioButton();
            this.btnPreview = new System.Windows.Forms.Button();
            this.lblProgress = new System.Windows.Forms.Label();
            this.queryTextBox = new System.Windows.Forms.TextBox();
            this.numberTextBox = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.newNumberTextBox = new System.Windows.Forms.TextBox();
            this.newQueryTextBox = new System.Windows.Forms.TextBox();
            this.lblProgressNew = new System.Windows.Forms.Label();
            this.newLoadButton = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.newTxtSeparatorOtherChar = new System.Windows.Forms.TextBox();
            this.newRdbSeparatorOther = new System.Windows.Forms.RadioButton();
            this.newRdbTab = new System.Windows.Forms.RadioButton();
            this.newRdbSemicolon = new System.Windows.Forms.RadioButton();
            this.newFileCheckBox = new System.Windows.Forms.CheckBox();
            this.dataGridView_preViewNew = new System.Windows.Forms.DataGridView();
            this.newFileButton = new System.Windows.Forms.Button();
            this.newFileTextBox = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.outputDataGridView = new System.Windows.Forms.DataGridView();
            this.differenceButton = new System.Windows.Forms.Button();
            this.filesizeLabel = new System.Windows.Forms.Label();
            this.subTractDataGridView = new System.Windows.Forms.DataGridView();
            this.subtractBtn = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_preView)).BeginInit();
            this.gpbSeparator.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_preViewNew)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.outputDataGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.subTractDataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(82, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "CSV file to load:";
            // 
            // txtFileToImport
            // 
            this.txtFileToImport.Location = new System.Drawing.Point(101, 12);
            this.txtFileToImport.Name = "txtFileToImport";
            this.txtFileToImport.Size = new System.Drawing.Size(292, 20);
            this.txtFileToImport.TabIndex = 1;
            this.txtFileToImport.TextChanged += new System.EventHandler(this.tbFile_TextChanged);
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(399, 12);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(108, 22);
            this.btnBrowse.TabIndex = 2;
            this.btnBrowse.Text = "Browse";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnFileOpen_Click);
            // 
            // dataGridView_preView
            // 
            this.dataGridView_preView.AllowUserToAddRows = false;
            this.dataGridView_preView.AllowUserToDeleteRows = false;
            this.dataGridView_preView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView_preView.Location = new System.Drawing.Point(12, 155);
            this.dataGridView_preView.Name = "dataGridView_preView";
            this.dataGridView_preView.ReadOnly = true;
            this.dataGridView_preView.Size = new System.Drawing.Size(497, 175);
            this.dataGridView_preView.TabIndex = 3;
            // 
            // chkFirstRowColumnNames
            // 
            this.chkFirstRowColumnNames.AutoSize = true;
            this.chkFirstRowColumnNames.Checked = true;
            this.chkFirstRowColumnNames.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkFirstRowColumnNames.Location = new System.Drawing.Point(169, 47);
            this.chkFirstRowColumnNames.Name = "chkFirstRowColumnNames";
            this.chkFirstRowColumnNames.Size = new System.Drawing.Size(156, 17);
            this.chkFirstRowColumnNames.TabIndex = 4;
            this.chkFirstRowColumnNames.Text = "First row has column names";
            this.chkFirstRowColumnNames.UseVisualStyleBackColor = true;
            // 
            // gpbSeparator
            // 
            this.gpbSeparator.Controls.Add(this.txtSeparatorOtherChar);
            this.gpbSeparator.Controls.Add(this.rdbSeparatorOther);
            this.gpbSeparator.Controls.Add(this.rdbTab);
            this.gpbSeparator.Controls.Add(this.rdbSemicolon);
            this.gpbSeparator.Location = new System.Drawing.Point(15, 47);
            this.gpbSeparator.Name = "gpbSeparator";
            this.gpbSeparator.Size = new System.Drawing.Size(129, 94);
            this.gpbSeparator.TabIndex = 5;
            this.gpbSeparator.TabStop = false;
            this.gpbSeparator.Text = "Separator";
            this.gpbSeparator.Visible = false;
            // 
            // txtSeparatorOtherChar
            // 
            this.txtSeparatorOtherChar.Location = new System.Drawing.Point(73, 66);
            this.txtSeparatorOtherChar.MaxLength = 1;
            this.txtSeparatorOtherChar.Name = "txtSeparatorOtherChar";
            this.txtSeparatorOtherChar.Size = new System.Drawing.Size(24, 20);
            this.txtSeparatorOtherChar.TabIndex = 3;
            this.txtSeparatorOtherChar.TextChanged += new System.EventHandler(this.txtSeparatorOtherChar_TextChanged);
            // 
            // rdbSeparatorOther
            // 
            this.rdbSeparatorOther.AutoSize = true;
            this.rdbSeparatorOther.Location = new System.Drawing.Point(6, 65);
            this.rdbSeparatorOther.Name = "rdbSeparatorOther";
            this.rdbSeparatorOther.Size = new System.Drawing.Size(54, 17);
            this.rdbSeparatorOther.TabIndex = 2;
            this.rdbSeparatorOther.Text = "Other:";
            this.rdbSeparatorOther.UseVisualStyleBackColor = true;
            // 
            // rdbTab
            // 
            this.rdbTab.AutoSize = true;
            this.rdbTab.Location = new System.Drawing.Point(6, 42);
            this.rdbTab.Name = "rdbTab";
            this.rdbTab.Size = new System.Drawing.Size(46, 17);
            this.rdbTab.TabIndex = 1;
            this.rdbTab.Text = "TAB";
            this.rdbTab.UseVisualStyleBackColor = true;
            // 
            // rdbSemicolon
            // 
            this.rdbSemicolon.AutoSize = true;
            this.rdbSemicolon.Checked = true;
            this.rdbSemicolon.Location = new System.Drawing.Point(6, 19);
            this.rdbSemicolon.Name = "rdbSemicolon";
            this.rdbSemicolon.Size = new System.Drawing.Size(74, 17);
            this.rdbSemicolon.TabIndex = 0;
            this.rdbSemicolon.TabStop = true;
            this.rdbSemicolon.Text = "Semicolon";
            this.rdbSemicolon.UseVisualStyleBackColor = true;
            // 
            // btnPreview
            // 
            this.btnPreview.Location = new System.Drawing.Point(169, 115);
            this.btnPreview.Name = "btnPreview";
            this.btnPreview.Size = new System.Drawing.Size(97, 25);
            this.btnPreview.TabIndex = 6;
            this.btnPreview.Text = "Load preview";
            this.btnPreview.UseVisualStyleBackColor = true;
            this.btnPreview.Click += new System.EventHandler(this.btnPreview_Click);
            // 
            // lblProgress
            // 
            this.lblProgress.AutoSize = true;
            this.lblProgress.Location = new System.Drawing.Point(12, 333);
            this.lblProgress.Name = "lblProgress";
            this.lblProgress.Size = new System.Drawing.Size(91, 13);
            this.lblProgress.TabIndex = 9;
            this.lblProgress.Text = "Imported: 0 row(s)";
            // 
            // queryTextBox
            // 
            this.queryTextBox.Location = new System.Drawing.Point(169, 70);
            this.queryTextBox.Multiline = true;
            this.queryTextBox.Name = "queryTextBox";
            this.queryTextBox.Size = new System.Drawing.Size(338, 39);
            this.queryTextBox.TabIndex = 15;
            // 
            // numberTextBox
            // 
            this.numberTextBox.Location = new System.Drawing.Point(304, 118);
            this.numberTextBox.Name = "numberTextBox";
            this.numberTextBox.Size = new System.Drawing.Size(74, 20);
            this.numberTextBox.TabIndex = 16;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(272, 121);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(26, 13);
            this.label2.TabIndex = 17;
            this.label2.Text = "First";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(395, 121);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(29, 13);
            this.label3.TabIndex = 18;
            this.label3.Text = "rows";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(1032, 121);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(29, 13);
            this.label4.TabIndex = 36;
            this.label4.Text = "rows";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(909, 121);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(26, 13);
            this.label5.TabIndex = 35;
            this.label5.Text = "First";
            // 
            // newNumberTextBox
            // 
            this.newNumberTextBox.Location = new System.Drawing.Point(941, 118);
            this.newNumberTextBox.Name = "newNumberTextBox";
            this.newNumberTextBox.Size = new System.Drawing.Size(74, 20);
            this.newNumberTextBox.TabIndex = 34;
            // 
            // newQueryTextBox
            // 
            this.newQueryTextBox.Location = new System.Drawing.Point(806, 70);
            this.newQueryTextBox.Multiline = true;
            this.newQueryTextBox.Name = "newQueryTextBox";
            this.newQueryTextBox.Size = new System.Drawing.Size(338, 39);
            this.newQueryTextBox.TabIndex = 33;
            // 
            // lblProgressNew
            // 
            this.lblProgressNew.AutoSize = true;
            this.lblProgressNew.Location = new System.Drawing.Point(649, 333);
            this.lblProgressNew.Name = "lblProgressNew";
            this.lblProgressNew.Size = new System.Drawing.Size(91, 13);
            this.lblProgressNew.TabIndex = 27;
            this.lblProgressNew.Text = "Imported: 0 row(s)";
            // 
            // newLoadButton
            // 
            this.newLoadButton.Location = new System.Drawing.Point(806, 115);
            this.newLoadButton.Name = "newLoadButton";
            this.newLoadButton.Size = new System.Drawing.Size(97, 25);
            this.newLoadButton.TabIndex = 25;
            this.newLoadButton.Text = "Load preview";
            this.newLoadButton.UseVisualStyleBackColor = true;
            this.newLoadButton.Click += new System.EventHandler(this.newLoadButton_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.newTxtSeparatorOtherChar);
            this.groupBox1.Controls.Add(this.newRdbSeparatorOther);
            this.groupBox1.Controls.Add(this.newRdbTab);
            this.groupBox1.Controls.Add(this.newRdbSemicolon);
            this.groupBox1.Location = new System.Drawing.Point(652, 47);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(129, 94);
            this.groupBox1.TabIndex = 24;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Separator";
            this.groupBox1.Visible = false;
            // 
            // newTxtSeparatorOtherChar
            // 
            this.newTxtSeparatorOtherChar.Location = new System.Drawing.Point(73, 66);
            this.newTxtSeparatorOtherChar.MaxLength = 1;
            this.newTxtSeparatorOtherChar.Name = "newTxtSeparatorOtherChar";
            this.newTxtSeparatorOtherChar.Size = new System.Drawing.Size(24, 20);
            this.newTxtSeparatorOtherChar.TabIndex = 3;
            this.newTxtSeparatorOtherChar.TextChanged += new System.EventHandler(this.newTxtSeparatorOtherChar_TextChanged);
            // 
            // newRdbSeparatorOther
            // 
            this.newRdbSeparatorOther.AutoSize = true;
            this.newRdbSeparatorOther.Location = new System.Drawing.Point(6, 65);
            this.newRdbSeparatorOther.Name = "newRdbSeparatorOther";
            this.newRdbSeparatorOther.Size = new System.Drawing.Size(54, 17);
            this.newRdbSeparatorOther.TabIndex = 2;
            this.newRdbSeparatorOther.Text = "Other:";
            this.newRdbSeparatorOther.UseVisualStyleBackColor = true;
            // 
            // newRdbTab
            // 
            this.newRdbTab.AutoSize = true;
            this.newRdbTab.Location = new System.Drawing.Point(6, 42);
            this.newRdbTab.Name = "newRdbTab";
            this.newRdbTab.Size = new System.Drawing.Size(46, 17);
            this.newRdbTab.TabIndex = 1;
            this.newRdbTab.Text = "TAB";
            this.newRdbTab.UseVisualStyleBackColor = true;
            // 
            // newRdbSemicolon
            // 
            this.newRdbSemicolon.AutoSize = true;
            this.newRdbSemicolon.Checked = true;
            this.newRdbSemicolon.Location = new System.Drawing.Point(6, 19);
            this.newRdbSemicolon.Name = "newRdbSemicolon";
            this.newRdbSemicolon.Size = new System.Drawing.Size(74, 17);
            this.newRdbSemicolon.TabIndex = 0;
            this.newRdbSemicolon.TabStop = true;
            this.newRdbSemicolon.Text = "Semicolon";
            this.newRdbSemicolon.UseVisualStyleBackColor = true;
            // 
            // newFileCheckBox
            // 
            this.newFileCheckBox.AutoSize = true;
            this.newFileCheckBox.Checked = true;
            this.newFileCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.newFileCheckBox.Location = new System.Drawing.Point(806, 47);
            this.newFileCheckBox.Name = "newFileCheckBox";
            this.newFileCheckBox.Size = new System.Drawing.Size(156, 17);
            this.newFileCheckBox.TabIndex = 23;
            this.newFileCheckBox.Text = "First row has column names";
            this.newFileCheckBox.UseVisualStyleBackColor = true;
            // 
            // dataGridView_preViewNew
            // 
            this.dataGridView_preViewNew.AllowUserToAddRows = false;
            this.dataGridView_preViewNew.AllowUserToDeleteRows = false;
            this.dataGridView_preViewNew.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView_preViewNew.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView_preViewNew.Location = new System.Drawing.Point(649, 155);
            this.dataGridView_preViewNew.Name = "dataGridView_preViewNew";
            this.dataGridView_preViewNew.ReadOnly = true;
            this.dataGridView_preViewNew.Size = new System.Drawing.Size(497, 175);
            this.dataGridView_preViewNew.TabIndex = 22;
            // 
            // newFileButton
            // 
            this.newFileButton.Location = new System.Drawing.Point(1036, 12);
            this.newFileButton.Name = "newFileButton";
            this.newFileButton.Size = new System.Drawing.Size(108, 22);
            this.newFileButton.TabIndex = 21;
            this.newFileButton.Text = "Browse";
            this.newFileButton.UseVisualStyleBackColor = true;
            this.newFileButton.Click += new System.EventHandler(this.newFileButton_Click);
            // 
            // newFileTextBox
            // 
            this.newFileTextBox.Location = new System.Drawing.Point(738, 12);
            this.newFileTextBox.Name = "newFileTextBox";
            this.newFileTextBox.Size = new System.Drawing.Size(292, 20);
            this.newFileTextBox.TabIndex = 20;
            this.newFileTextBox.TextChanged += new System.EventHandler(this.newFileTextBox_TextChanged);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(649, 16);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(82, 13);
            this.label9.TabIndex = 19;
            this.label9.Text = "CSV file to load:";
            // 
            // outputDataGridView
            // 
            this.outputDataGridView.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)));
            this.outputDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.outputDataGridView.Location = new System.Drawing.Point(15, 364);
            this.outputDataGridView.Name = "outputDataGridView";
            this.outputDataGridView.Size = new System.Drawing.Size(494, 105);
            this.outputDataGridView.TabIndex = 37;
            // 
            // differenceButton
            // 
            this.differenceButton.Location = new System.Drawing.Point(515, 364);
            this.differenceButton.Name = "differenceButton";
            this.differenceButton.Size = new System.Drawing.Size(75, 23);
            this.differenceButton.TabIndex = 38;
            this.differenceButton.Text = "Difference";
            this.differenceButton.UseVisualStyleBackColor = true;
            this.differenceButton.Click += new System.EventHandler(this.differenceButton_Click);
            // 
            // filesizeLabel
            // 
            this.filesizeLabel.AutoSize = true;
            this.filesizeLabel.Location = new System.Drawing.Point(232, 333);
            this.filesizeLabel.Name = "filesizeLabel";
            this.filesizeLabel.Size = new System.Drawing.Size(50, 13);
            this.filesizeLabel.TabIndex = 39;
            this.filesizeLabel.Text = "File size: ";
            // 
            // subTractDataGridView
            // 
            this.subTractDataGridView.AllowUserToAddRows = false;
            this.subTractDataGridView.AllowUserToDeleteRows = false;
            this.subTractDataGridView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.subTractDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.subTractDataGridView.Location = new System.Drawing.Point(647, 364);
            this.subTractDataGridView.Name = "subTractDataGridView";
            this.subTractDataGridView.ReadOnly = true;
            this.subTractDataGridView.Size = new System.Drawing.Size(497, 105);
            this.subTractDataGridView.TabIndex = 40;
            // 
            // subtractBtn
            // 
            this.subtractBtn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.subtractBtn.Location = new System.Drawing.Point(1151, 363);
            this.subtractBtn.Name = "subtractBtn";
            this.subtractBtn.Size = new System.Drawing.Size(75, 23);
            this.subtractBtn.TabIndex = 41;
            this.subtractBtn.Text = "Subtract";
            this.subtractBtn.UseVisualStyleBackColor = true;
            this.subtractBtn.Click += new System.EventHandler(this.subtractBtn_Click);
            // 
            // frmImport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1252, 481);
            this.Controls.Add(this.subtractBtn);
            this.Controls.Add(this.subTractDataGridView);
            this.Controls.Add(this.filesizeLabel);
            this.Controls.Add(this.differenceButton);
            this.Controls.Add(this.outputDataGridView);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.newNumberTextBox);
            this.Controls.Add(this.newQueryTextBox);
            this.Controls.Add(this.lblProgressNew);
            this.Controls.Add(this.newLoadButton);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.newFileCheckBox);
            this.Controls.Add(this.dataGridView_preViewNew);
            this.Controls.Add(this.newFileButton);
            this.Controls.Add(this.newFileTextBox);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.numberTextBox);
            this.Controls.Add(this.queryTextBox);
            this.Controls.Add(this.lblProgress);
            this.Controls.Add(this.btnPreview);
            this.Controls.Add(this.gpbSeparator);
            this.Controls.Add(this.chkFirstRowColumnNames);
            this.Controls.Add(this.dataGridView_preView);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.txtFileToImport);
            this.Controls.Add(this.label1);
            this.MinimumSize = new System.Drawing.Size(527, 515);
            this.Name = "frmImport";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Import CSV";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_preView)).EndInit();
            this.gpbSeparator.ResumeLayout(false);
            this.gpbSeparator.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_preViewNew)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.outputDataGridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.subTractDataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox txtFileToImport;
		private System.Windows.Forms.Button btnBrowse;
		private System.Windows.Forms.DataGridView dataGridView_preView;
		private System.Windows.Forms.CheckBox chkFirstRowColumnNames;
		private System.Windows.Forms.GroupBox gpbSeparator;
		private System.Windows.Forms.RadioButton rdbSeparatorOther;
		private System.Windows.Forms.RadioButton rdbTab;
		private System.Windows.Forms.RadioButton rdbSemicolon;
		private System.Windows.Forms.TextBox txtSeparatorOtherChar;
        private System.Windows.Forms.Button btnPreview;
        private System.Windows.Forms.Label lblProgress;
        private System.Windows.Forms.TextBox queryTextBox;
        private System.Windows.Forms.TextBox numberTextBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox newNumberTextBox;
        private System.Windows.Forms.TextBox newQueryTextBox;
        private System.Windows.Forms.Label lblProgressNew;
        private System.Windows.Forms.Button newLoadButton;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox newTxtSeparatorOtherChar;
        private System.Windows.Forms.RadioButton newRdbSeparatorOther;
        private System.Windows.Forms.RadioButton newRdbTab;
        private System.Windows.Forms.RadioButton newRdbSemicolon;
        private System.Windows.Forms.CheckBox newFileCheckBox;
        private System.Windows.Forms.DataGridView dataGridView_preViewNew;
        private System.Windows.Forms.Button newFileButton;
        private System.Windows.Forms.TextBox newFileTextBox;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.DataGridView outputDataGridView;
        private System.Windows.Forms.Button differenceButton;
        private System.Windows.Forms.Label filesizeLabel;
        private System.Windows.Forms.DataGridView subTractDataGridView;
        private System.Windows.Forms.Button subtractBtn;
	}
}