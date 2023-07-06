using System;
using System.Drawing;
using System.Windows.Forms;

namespace friendly_bpa2excel
{
    partial class Form1
    {
        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpar os recursos que estão sendo usados.
        /// </summary>
        /// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código gerado pelo Windows Form Designer

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.button5 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(18, 32);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(28, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "BPA";
            this.label1.Visible = false;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(58, 29);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(227, 20);
            this.textBox1.TabIndex = 3;
            this.textBox1.Visible = false;
            this.textBox1.TextChanged += textBox1_textChange;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(290, 29);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(26, 20);
            this.button3.TabIndex = 4;
            this.button3.Text = "...";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Visible = false;
            this.button3.Click += button3_Click;
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(123, 327);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(64, 20);
            this.button4.TabIndex = 5;
            this.button4.Text = "Converter";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Visible = false;
            this.button4.Click += button4_Click;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(51, 170);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(79, 17);
            this.checkBox1.TabIndex = 7;
            this.checkBox1.Text = "BPA>Excel";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.CheckedChanged += checkBox1_CheckedChanged;
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.Location = new System.Drawing.Point(51, 208);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(79, 17);
            this.checkBox2.TabIndex = 8;
            this.checkBox2.Text = "Excel>BPA";
            this.checkBox2.UseVisualStyleBackColor = true;
            this.checkBox2.CheckedChanged += checkBox2_CheckedChanged;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(58, 55);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(227, 20);
            this.textBox2.TabIndex = 9;
            this.textBox2.Visible = false;
            this.textBox2.TextChanged += textBox2_textChange;
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(58, 80);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(227, 20);
            this.textBox3.TabIndex = 10;
            this.textBox3.Visible = false;
            this.textBox3.TextChanged += textBox3_textChange;
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(290, 55);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(26, 20);
            this.button5.TabIndex = 11;
            this.button5.Text = "...";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Visible = false;
            this.button5.Click += button5_Click;
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(290, 80);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(26, 20);
            this.button6.TabIndex = 12;
            this.button6.Text = "...";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Visible = false;
            this.button6.Click += button6_Click;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 58);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(35, 13);
            this.label3.TabIndex = 13;
            this.label3.Text = "BPAC";
            this.label3.Visible = false;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(10, 83);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(31, 13);
            this.label4.TabIndex = 14;
            this.label4.Text = "BPAI";
            this.label4.Visible = false;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(9, 274);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(305, 20);
            this.progressBar1.TabIndex = 15;
            this.progressBar1.Visible = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(332, 390);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.button6);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.checkBox2);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "Form1";
            this.Text = "Friendly_bpa2excel";
            this.ResumeLayout(false);
            this.PerformLayout();
            this.ShowIcon = false;

        }

        #endregion
        private Label label1;
        private TextBox textBox1;
        private Button button3;
        private Button button4;
        private CheckBox checkBox1;
        private CheckBox checkBox2;
        private TextBox textBox2;
        private TextBox textBox3;
        private Button button5;
        private Button button6;
        private Label label3;
        private Label label4;
        private ProgressBar progressBar1;
    }
}

