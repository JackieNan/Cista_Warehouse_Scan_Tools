namespace Warehouse_Scan_Tools
{
    partial class Form2
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
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(166, 117);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(535, 28);
            this.textBox1.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 16F);
            this.label1.Location = new System.Drawing.Point(33, 112);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(111, 33);
            this.label1.TabIndex = 1;
            this.label1.Text = "文件名";
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("宋体", 16F);
            this.button1.Location = new System.Drawing.Point(117, 329);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(180, 67);
            this.button1.TabIndex = 2;
            this.button1.Text = "选择文件";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("宋体", 16F);
            this.button2.Location = new System.Drawing.Point(521, 329);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(180, 67);
            this.button2.TabIndex = 3;
            this.button2.Text = "继续装单";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("宋体", 13F);
            this.button3.Location = new System.Drawing.Point(653, 23);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(110, 52);
            this.button3.TabIndex = 4;
            this.button3.Text = "结单";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox1);
            this.Name = "Form2";
            this.Text = "Form2";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
    }
}