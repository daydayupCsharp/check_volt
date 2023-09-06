namespace check_volt
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            richTextBox1 = new RichTextBox();
            textBox1 = new TextBox();
            richTextBox2 = new RichTextBox();
            label1 = new Label();
            label2 = new Label();
            richTextBox3 = new RichTextBox();
            button1 = new Button();
            textBox2 = new TextBox();
            checkBox1 = new CheckBox();
            checkBox2 = new CheckBox();
            checkBox3 = new CheckBox();
            checkBox4 = new CheckBox();
            checkBox5 = new CheckBox();
            checkBox6 = new CheckBox();
            checkBox7 = new CheckBox();
            SuspendLayout();
            // 
            // richTextBox1
            // 
            richTextBox1.Location = new Point(2, 341);
            richTextBox1.Name = "richTextBox1";
            richTextBox1.ReadOnly = true;
            richTextBox1.Size = new Size(479, 71);
            richTextBox1.TabIndex = 1;
            richTextBox1.Text = "";
            // 
            // textBox1
            // 
            textBox1.Location = new Point(389, 52);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(439, 23);
            textBox1.TabIndex = 2;
            textBox1.KeyPress += textpack_KeyPress;
            // 
            // richTextBox2
            // 
            richTextBox2.Location = new Point(554, 341);
            richTextBox2.Name = "richTextBox2";
            richTextBox2.ReadOnly = true;
            richTextBox2.Size = new Size(447, 71);
            richTextBox2.TabIndex = 3;
            richTextBox2.Text = "";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(293, 52);
            label1.Name = "label1";
            label1.Size = new Size(56, 17);
            label1.TabIndex = 4;
            label1.Text = "包体码：";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(293, 81);
            label2.Name = "label2";
            label2.Size = new Size(68, 17);
            label2.TabIndex = 5;
            label2.Text = "包体结果：";
            // 
            // richTextBox3
            // 
            richTextBox3.Location = new Point(389, 81);
            richTextBox3.Name = "richTextBox3";
            richTextBox3.ReadOnly = true;
            richTextBox3.Size = new Size(477, 236);
            richTextBox3.TabIndex = 6;
            richTextBox3.Text = "";
            // 
            // button1
            // 
            button1.Location = new Point(19, 10);
            button1.Name = "button1";
            button1.Size = new Size(104, 23);
            button1.TabIndex = 7;
            button1.Text = "临时更改设置";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // textBox2
            // 
            textBox2.Location = new Point(129, 10);
            textBox2.Name = "textBox2";
            textBox2.Size = new Size(632, 23);
            textBox2.TabIndex = 8;
            textBox2.Text = "D:\\ATETest-20230524-V2.00-Debug-30-5\\TestData";
            textBox2.Visible = false;
            // 
            // checkBox1
            // 
            checkBox1.AutoSize = true;
            checkBox1.Checked = true;
            checkBox1.CheckState = CheckState.Checked;
            checkBox1.Location = new Point(2, 106);
            checkBox1.Name = "checkBox1";
            checkBox1.Size = new Size(135, 21);
            checkBox1.TabIndex = 9;
            checkBox1.Text = "包体档位和最小容量";
            checkBox1.UseVisualStyleBackColor = true;
            checkBox1.Visible = false;
            // 
            // checkBox2
            // 
            checkBox2.AutoSize = true;
            checkBox2.Checked = true;
            checkBox2.CheckState = CheckState.Checked;
            checkBox2.Location = new Point(2, 133);
            checkBox2.Name = "checkBox2";
            checkBox2.Size = new Size(135, 21);
            checkBox2.TabIndex = 10;
            checkBox2.Text = "包内电芯是否被隔离";
            checkBox2.UseVisualStyleBackColor = true;
            checkBox2.Visible = false;
            // 
            // checkBox3
            // 
            checkBox3.AutoSize = true;
            checkBox3.Checked = true;
            checkBox3.CheckState = CheckState.Checked;
            checkBox3.Location = new Point(2, 160);
            checkBox3.Name = "checkBox3";
            checkBox3.Size = new Size(99, 21);
            checkBox3.TabIndex = 11;
            checkBox3.Text = "包是否被隔离";
            checkBox3.UseVisualStyleBackColor = true;
            checkBox3.Visible = false;
            // 
            // checkBox4
            // 
            checkBox4.AutoSize = true;
            checkBox4.Checked = true;
            checkBox4.CheckState = CheckState.Checked;
            checkBox4.Location = new Point(2, 187);
            checkBox4.Name = "checkBox4";
            checkBox4.Size = new Size(113, 21);
            checkBox4.TabIndex = 12;
            checkBox4.Text = "校验静测Bv文件";
            checkBox4.UseVisualStyleBackColor = true;
            checkBox4.Visible = false;
            // 
            // checkBox5
            // 
            checkBox5.AutoSize = true;
            checkBox5.Location = new Point(2, 227);
            checkBox5.Name = "checkBox5";
            checkBox5.Size = new Size(135, 21);
            checkBox5.TabIndex = 13;
            checkBox5.Text = "模组档位和最小容量";
            checkBox5.UseVisualStyleBackColor = true;
            checkBox5.Visible = false;
            // 
            // checkBox6
            // 
            checkBox6.AutoSize = true;
            checkBox6.Location = new Point(2, 254);
            checkBox6.Name = "checkBox6";
            checkBox6.Size = new Size(147, 21);
            checkBox6.TabIndex = 14;
            checkBox6.Text = "模组内电芯是否被隔离";
            checkBox6.UseVisualStyleBackColor = true;
            checkBox6.Visible = false;
            // 
            // checkBox7
            // 
            checkBox7.AutoSize = true;
            checkBox7.Location = new Point(2, 281);
            checkBox7.Name = "checkBox7";
            checkBox7.Size = new Size(173, 21);
            checkBox7.TabIndex = 15;
            checkBox7.Text = "校验包体或模组采样Bv文件";
            checkBox7.UseVisualStyleBackColor = true;
            checkBox7.Visible = false;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 17F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1013, 450);
            Controls.Add(checkBox7);
            Controls.Add(checkBox6);
            Controls.Add(checkBox5);
            Controls.Add(checkBox4);
            Controls.Add(checkBox3);
            Controls.Add(checkBox2);
            Controls.Add(checkBox1);
            Controls.Add(textBox2);
            Controls.Add(button1);
            Controls.Add(richTextBox3);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(richTextBox2);
            Controls.Add(textBox1);
            Controls.Add(richTextBox1);
            Name = "Form1";
            Text = "静态包体电压检测防呆";
            Load += Form1_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private RichTextBox richTextBox1;
        private TextBox textBox1;
        private RichTextBox richTextBox2;
        private Label label1;
        private Label label2;
        private RichTextBox richTextBox3;
        private Button button1;
        private TextBox textBox2;
        private CheckBox checkBox1;
        private CheckBox checkBox2;
        private CheckBox checkBox3;
        private CheckBox checkBox4;
        private CheckBox checkBox5;
        private CheckBox checkBox6;
        private CheckBox checkBox7;
    }
}