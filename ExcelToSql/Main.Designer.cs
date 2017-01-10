namespace ExcelToSql
{
    partial class Main
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
            this.txt_filepath = new System.Windows.Forms.TextBox();
            this.btn_convert = new System.Windows.Forms.Button();
            this.btn_broswerFile = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 39);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(71, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "Excel路径：";
            // 
            // txt_filepath
            // 
            this.txt_filepath.Location = new System.Drawing.Point(91, 39);
            this.txt_filepath.Name = "txt_filepath";
            this.txt_filepath.Size = new System.Drawing.Size(263, 21);
            this.txt_filepath.TabIndex = 1;
            // 
            // btn_convert
            // 
            this.btn_convert.Location = new System.Drawing.Point(177, 81);
            this.btn_convert.Name = "btn_convert";
            this.btn_convert.Size = new System.Drawing.Size(75, 23);
            this.btn_convert.TabIndex = 2;
            this.btn_convert.Text = "生成sql";
            this.btn_convert.UseVisualStyleBackColor = true;
            this.btn_convert.Click += new System.EventHandler(this.btn_convert_Click);
            // 
            // btn_broswerFile
            // 
            this.btn_broswerFile.Location = new System.Drawing.Point(377, 39);
            this.btn_broswerFile.Name = "btn_broswerFile";
            this.btn_broswerFile.Size = new System.Drawing.Size(75, 21);
            this.btn_broswerFile.TabIndex = 3;
            this.btn_broswerFile.Text = "浏览";
            this.btn_broswerFile.UseVisualStyleBackColor = true;
            this.btn_broswerFile.Click += new System.EventHandler(this.btn_broswerFile_Click);
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(471, 133);
            this.Controls.Add(this.btn_broswerFile);
            this.Controls.Add(this.btn_convert);
            this.Controls.Add(this.txt_filepath);
            this.Controls.Add(this.label1);
            this.MaximizeBox = false;
            this.Name = "Main";
            this.Text = "Main";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txt_filepath;
        private System.Windows.Forms.Button btn_convert;
        private System.Windows.Forms.Button btn_broswerFile;
    }
}