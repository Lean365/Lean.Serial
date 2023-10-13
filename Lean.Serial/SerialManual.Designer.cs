namespace Lean.Serial
{
    partial class SerialManual
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SerialManual));
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.dateTimePicker2 = new System.Windows.Forms.DateTimePicker();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.DgvSndata = new System.Windows.Forms.DataGridView();
            this.DgvSnDataCsv = new System.Windows.Forms.DataGridView();
            this.button1 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.DgvSndata)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DgvSnDataCsv)).BeginInit();
            this.SuspendLayout();
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(45, 67);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(125, 21);
            this.dateTimePicker1.TabIndex = 26;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.ForeColor = System.Drawing.Color.Red;
            this.label1.Location = new System.Drawing.Point(94, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(190, 34);
            this.label1.TabIndex = 28;
            this.label1.Text = "自定义上传";
            // 
            // dateTimePicker2
            // 
            this.dateTimePicker2.Location = new System.Drawing.Point(45, 97);
            this.dateTimePicker2.Name = "dateTimePicker2";
            this.dateTimePicker2.Size = new System.Drawing.Size(125, 21);
            this.dateTimePicker2.TabIndex = 29;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(6, 184);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(361, 23);
            this.progressBar1.TabIndex = 30;
            // 
            // DgvSndata
            // 
            this.DgvSndata.AccessibleDescription = "Sndata";
            this.DgvSndata.BackgroundColor = System.Drawing.SystemColors.ButtonFace;
            this.DgvSndata.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DgvSndata.Location = new System.Drawing.Point(6, 134);
            this.DgvSndata.Name = "DgvSndata";
            this.DgvSndata.ReadOnly = true;
            this.DgvSndata.RowTemplate.Height = 23;
            this.DgvSndata.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.DgvSndata.Size = new System.Drawing.Size(361, 45);
            this.DgvSndata.TabIndex = 23;
            // 
            // DgvSnDataCsv
            // 
            this.DgvSnDataCsv.AccessibleDescription = "SnDataCsv";
            this.DgvSnDataCsv.BackgroundColor = System.Drawing.SystemColors.ButtonFace;
            this.DgvSnDataCsv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DgvSnDataCsv.Location = new System.Drawing.Point(6, 209);
            this.DgvSnDataCsv.Name = "DgvSnDataCsv";
            this.DgvSnDataCsv.ReadOnly = true;
            this.DgvSnDataCsv.RowTemplate.Height = 23;
            this.DgvSnDataCsv.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.DgvSnDataCsv.Size = new System.Drawing.Size(361, 40);
            this.DgvSnDataCsv.TabIndex = 24;
            // 
            // button1
            // 
            this.button1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.button1.Image = global::Lean.Serial.Properties.Resources.confirm;
            this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button1.Location = new System.Drawing.Point(231, 75);
            this.button1.Name = "button1";
            this.button1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.button1.Size = new System.Drawing.Size(80, 40);
            this.button1.TabIndex = 27;
            this.button1.Text = "上传";
            this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // SerialManual
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(384, 261);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.dateTimePicker2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dateTimePicker1);
            this.Controls.Add(this.DgvSnDataCsv);
            this.Controls.Add(this.DgvSndata);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "SerialManual";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Serail Manual";
            this.Load += new System.EventHandler(this.SerialManual_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DgvSndata)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DgvSnDataCsv)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dateTimePicker2;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.DataGridView DgvSndata;
        private System.Windows.Forms.DataGridView DgvSnDataCsv;
    }
}