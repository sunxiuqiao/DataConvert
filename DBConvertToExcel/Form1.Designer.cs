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
            this.ConvertToshp = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.spatialdbPath = new System.Windows.Forms.TextBox();
            this.ChooseSpatialPath = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(24, 17);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "属性表路径：";
            // 
            // AttributePath
            // 
            this.AttributePath.Location = new System.Drawing.Point(97, 12);
            this.AttributePath.Name = "AttributePath";
            this.AttributePath.Size = new System.Drawing.Size(238, 21);
            this.AttributePath.TabIndex = 1;
            // 
            // ChoosePath
            // 
            this.ChoosePath.Location = new System.Drawing.Point(362, 12);
            this.ChoosePath.Name = "ChoosePath";
            this.ChoosePath.Size = new System.Drawing.Size(75, 23);
            this.ChoosePath.TabIndex = 2;
            this.ChoosePath.Text = "选择路径";
            this.ChoosePath.UseVisualStyleBackColor = true;
            this.ChoosePath.Click += new System.EventHandler(this.ChoosePath_Click);
            // 
            // ConverToExcel
            // 
            this.ConverToExcel.Location = new System.Drawing.Point(260, 116);
            this.ConverToExcel.Name = "ConverToExcel";
            this.ConverToExcel.Size = new System.Drawing.Size(75, 23);
            this.ConverToExcel.TabIndex = 3;
            this.ConverToExcel.Text = "导出excel";
            this.ConverToExcel.UseVisualStyleBackColor = true;
            this.ConverToExcel.Click += new System.EventHandler(this.ConverToExcel_Click);
            // 
            // ConvertToshp
            // 
            this.ConvertToshp.Location = new System.Drawing.Point(350, 116);
            this.ConvertToshp.Name = "ConvertToshp";
            this.ConvertToshp.Size = new System.Drawing.Size(75, 23);
            this.ConvertToshp.TabIndex = 10;
            this.ConvertToshp.Text = "导出shp";
            this.ConvertToshp.UseVisualStyleBackColor = true;
            this.ConvertToshp.Click += new System.EventHandler(this.ConvertToshp_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 65);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(89, 12);
            this.label4.TabIndex = 11;
            this.label4.Text = "空间数据路径：";
            // 
            // spatialdbPath
            // 
            this.spatialdbPath.Location = new System.Drawing.Point(97, 57);
            this.spatialdbPath.Name = "spatialdbPath";
            this.spatialdbPath.Size = new System.Drawing.Size(238, 21);
            this.spatialdbPath.TabIndex = 12;
            // 
            // ChooseSpatialPath
            // 
            this.ChooseSpatialPath.Location = new System.Drawing.Point(362, 54);
            this.ChooseSpatialPath.Name = "ChooseSpatialPath";
            this.ChooseSpatialPath.Size = new System.Drawing.Size(75, 23);
            this.ChooseSpatialPath.TabIndex = 13;
            this.ChooseSpatialPath.Text = "选择路径";
            this.ChooseSpatialPath.UseVisualStyleBackColor = true;
            this.ChooseSpatialPath.Click += new System.EventHandler(this.ChooseSpatialPath_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(452, 164);
            this.Controls.Add(this.ChooseSpatialPath);
            this.Controls.Add(this.spatialdbPath);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.ConvertToshp);
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
        private System.Windows.Forms.Button ConvertToshp;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox spatialdbPath;
        private System.Windows.Forms.Button ChooseSpatialPath;
    }
}

