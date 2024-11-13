namespace SaljPartOrderForms
{
    partial class inputdate
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
            this.button1 = new System.Windows.Forms.Button();
            this.dtLevdate = new System.Windows.Forms.DateTimePicker();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(89, 81);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 1;
            this.button1.Text = "OK";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // dtLevdate
            // 
            this.dtLevdate.CustomFormat = "yyMMdd";
            this.dtLevdate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtLevdate.Location = new System.Drawing.Point(63, 29);
            this.dtLevdate.Name = "dtLevdate";
            this.dtLevdate.Size = new System.Drawing.Size(124, 20);
            this.dtLevdate.TabIndex = 2;
            // 
            // inputdate
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(256, 116);
            this.Controls.Add(this.dtLevdate);
            this.Controls.Add(this.button1);
            this.Name = "inputdate";
            this.Text = "Ange leveransdatum";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.DateTimePicker dtLevdate;
    }
}