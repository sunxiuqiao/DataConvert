namespace DBConvertToExcel
{
    partial class Form1
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
            this.AttributePath = new System.Windows.Forms.TextBox();
            this.ChoosePath = new System.Windows.Forms.Button();
            this.ConverToExcel = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.ExcelPath = new System.Windows.Forms.TextBox();
            this.ChooseExcelPath = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.shppath = new System.Windows.Forms.TextBox();
            this.ChooseshpPath = new System.Windows.Forms.Button();
            this.ConvertToshp = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "属性表路径：";
            // 
            // AttributePath
            // 
            this.AttributePath.Location = new System.Drawing.Point(110, 23);
            this.AttributePath.Name = "AttributePath";
            this.AttributePath.Size = new System.Drawing.Size(225, 21);
            this.AttributePath.TabIndex = 1;
            // 
            // ChoosePath
            // 
            this.ChoosePath.Location = new System.Drawing.Point(362, 23);
            this.ChoosePath.Name = "ChoosePath";
            this.ChoosePath.Size = new System.Drawing.Size(75, 23);
            this.ChoosePath.TabIndex = 2;
            this.ChoosePath.Text = "选择路径";
            this.ChoosePath.UseVisualStyleBackColor = true;
            this.ChoosePath.Click += new System.EventHandler(this.ChoosePath_Click);
            // 
            // ConverToExcel
            // 
            this.ConverToExcel.Location = new System.Drawing.Point(271, 190);
            this.ConverToExcel.Name = "ConverToExcel";
            this.ConverToExcel.Size = new System.Drawing.Size(75, 23);
            this.ConverToExcel.TabIndex = 3;
            this.ConverToExcel.Text = "导出excel";
            this.ConverToExcel.UseVisualStyleBackColor = true;
            this.ConverToExcel.Click += new System.EventHandler(this.ConverToExcel_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(14, 77);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(101, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "选择导出表路径：";
            // 
            // ExcelPath
            // 
            this.ExcelPath.Location = new System.Drawing.Point(110, 69);
            this.ExcelPath.Name = "ExcelPath";
            this.ExcelPath.Size = new System.Drawing.Size(225, 21);
            this.ExcelPath.TabIndex = 5;
            // 
            // ChooseExcelPath
            // 
            this.ChooseExcelPath.Location = new System.Drawing.Point(362, 69);
            this.ChooseExcelPath.Name = "ChooseExcelPath";
            this.ChooseExcelPath.Size = new System.Drawing.Size(75, 23);
            this.ChooseExcelPath.TabIndex = 6;
            this.ChooseExcelPath.Text = "选择导出路径";
            this.ChooseExcelPath.UseVisualStyleBackColor = true;
            this.ChooseExcelPath.Click += new System.EventHandler(this.ChooseExcelPath_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(8, 122);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(107, 12);
            this.label3.TabIndex = 7;
            this.label3.Text = "选择导出shp路径：";
            // 
            // shppath
            // 
            this.shppath.Location = new System.Drawing.Point(110, 113);
            this.shppath.Name = "shppath";
            this.shppath.Size = new System.Drawing.Size(225, 21);
            this.shppath.TabIndex = 8;
            // 
            // ChooseshpPath
            // 
            this.ChooseshpPath.Location = new System.Drawing.Point(362, 122);
            this.ChooseshpPath.Name = "ChooseshpPath";
            this.ChooseshpPath.Size = new System.Drawing.Size(75, 23);
            this.ChooseshpPath.TabIndex = 9;
            this.ChooseshpPath.Text = "选择shp路径";
            this.ChooseshpPath.UseVisualStyleBackColor = true;
            this.ChooseshpPath.Click += new System.EventHandler(this.ChooseshpPath_Click);
            // 
            // ConvertToshp
            // 
            this.ConvertToshp.Location = new System.Drawing.Point(362, 190);
            this.ConvertToshp.Name = "ConvertToshp";
            this.ConvertToshp.Size = new System.Drawing.Size(75, 23);
            this.ConvertToshp.TabIndex = 10;
            this.ConvertToshp.Text = "导出shp";
            this.ConvertToshp.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(487, 348);
            this.Controls.Add(this.ConvertToshp);
            this.Controls.Add(this.ChooseshpPath);
            this.Controls.Add(this.shppath);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.ChooseExcelPath);
            this.Controls.Add(this.ExcelPath);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.ConverToExcel);
            this.Controls.Add(this.ChoosePath);
            this.Controls.Add(this.AttributePath);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox AttributePath;
        private System.Windows.Forms.Button ChoosePath;
        private System.Windows.Forms.Button ConverToExcel;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox ExcelPath;
        private System.Windows.Forms.Button ChooseExcelPath;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox shppath;
        private System.Windows.Forms.Button ChooseshpPath;
        private System.Windows.Forms.Button ConvertToshp;
    }
}

