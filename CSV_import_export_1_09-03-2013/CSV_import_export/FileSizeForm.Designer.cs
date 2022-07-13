namespace CSV_import_export
{
    partial class FileSizeForm
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
            this.fileTextBox = new System.Windows.Forms.TextBox();
            this.selectButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.newDataGridView = new System.Windows.Forms.DataGridView();
            this.label4 = new System.Windows.Forms.Label();
            this.newFileButton = new System.Windows.Forms.Button();
            this.newFileTextBox = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.diffDataGridView = new System.Windows.Forms.DataGridView();
            this.diffButton = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.newDataGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.diffDataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // fileTextBox
            // 
            this.fileTextBox.Location = new System.Drawing.Point(81, 20);
            this.fileTextBox.Multiline = true;
            this.fileTextBox.Name = "fileTextBox";
            this.fileTextBox.Size = new System.Drawing.Size(186, 28);
            this.fileTextBox.TabIndex = 0;
            // 
            // selectButton
            // 
            this.selectButton.Location = new System.Drawing.Point(285, 20);
            this.selectButton.Name = "selectButton";
            this.selectButton.Size = new System.Drawing.Size(75, 23);
            this.selectButton.TabIndex = 1;
            this.selectButton.Text = "Browse";
            this.selectButton.UseVisualStyleBackColor = true;
            this.selectButton.Click += new System.EventHandler(this.selectButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(-1, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(76, 17);
            this.label1.TabIndex = 2;
            this.label1.Text = "Select files";
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(2, 96);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(358, 268);
            this.dataGridView1.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(-1, 63);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(116, 17);
            this.label2.TabIndex = 5;
            this.label2.Text = "Previous file Size";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(414, 63);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(108, 17);
            this.label3.TabIndex = 10;
            this.label3.Text = "Current file Size";
            // 
            // newDataGridView
            // 
            this.newDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.newDataGridView.Location = new System.Drawing.Point(417, 96);
            this.newDataGridView.Name = "newDataGridView";
            this.newDataGridView.Size = new System.Drawing.Size(358, 268);
            this.newDataGridView.TabIndex = 9;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(414, 26);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(76, 17);
            this.label4.TabIndex = 8;
            this.label4.Text = "Select files";
            // 
            // newFileButton
            // 
            this.newFileButton.Location = new System.Drawing.Point(700, 20);
            this.newFileButton.Name = "newFileButton";
            this.newFileButton.Size = new System.Drawing.Size(75, 23);
            this.newFileButton.TabIndex = 7;
            this.newFileButton.Text = "Browse";
            this.newFileButton.UseVisualStyleBackColor = true;
            this.newFileButton.Click += new System.EventHandler(this.newFileButton_Click);
            // 
            // newFileTextBox
            // 
            this.newFileTextBox.Location = new System.Drawing.Point(496, 20);
            this.newFileTextBox.Multiline = true;
            this.newFileTextBox.Name = "newFileTextBox";
            this.newFileTextBox.Size = new System.Drawing.Size(184, 28);
            this.newFileTextBox.TabIndex = 6;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(842, 63);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(244, 17);
            this.label5.TabIndex = 12;
            this.label5.Text = "Difference of previous Vs Current File";
            // 
            // diffDataGridView
            // 
            this.diffDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.diffDataGridView.Location = new System.Drawing.Point(845, 96);
            this.diffDataGridView.Name = "diffDataGridView";
            this.diffDataGridView.Size = new System.Drawing.Size(358, 268);
            this.diffDataGridView.TabIndex = 11;
            // 
            // diffButton
            // 
            this.diffButton.Location = new System.Drawing.Point(1106, 63);
            this.diffButton.Name = "diffButton";
            this.diffButton.Size = new System.Drawing.Size(97, 23);
            this.diffButton.TabIndex = 13;
            this.diffButton.Text = "Show Difference";
            this.diffButton.UseVisualStyleBackColor = true;
            this.diffButton.Click += new System.EventHandler(this.diffButton_Click);
            // 
            // FileSizeForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1267, 482);
            this.Controls.Add(this.diffButton);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.diffDataGridView);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.newDataGridView);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.newFileButton);
            this.Controls.Add(this.newFileTextBox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.selectButton);
            this.Controls.Add(this.fileTextBox);
            this.Name = "FileSizeForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FileSizeForm";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.newDataGridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.diffDataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox fileTextBox;
        private System.Windows.Forms.Button selectButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DataGridView newDataGridView;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button newFileButton;
        private System.Windows.Forms.TextBox newFileTextBox;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.DataGridView diffDataGridView;
        private System.Windows.Forms.Button diffButton;
    }
}