namespace PilotOCR
{
    partial class ProgressDialog
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
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.Comment = new System.Windows.Forms.Label();
            this.cancelBtn = new System.Windows.Forms.Button();
            this.Comment2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(12, 78);
            this.progressBar1.Maximum = 10;
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(598, 23);
            this.progressBar1.Step = 1;
            this.progressBar1.TabIndex = 0;
            // 
            // Comment
            // 
            this.Comment.AutoSize = true;
            this.Comment.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.Comment.Location = new System.Drawing.Point(12, 9);
            this.Comment.Margin = new System.Windows.Forms.Padding(3, 0, 3, 1);
            this.Comment.Name = "Comment";
            this.Comment.Size = new System.Drawing.Size(68, 13);
            this.Comment.TabIndex = 1;
            this.Comment.Text = "Распознано";
            this.Comment.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // cancelBtn
            // 
            this.cancelBtn.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelBtn.Location = new System.Drawing.Point(537, 107);
            this.cancelBtn.Name = "cancelBtn";
            this.cancelBtn.Size = new System.Drawing.Size(75, 23);
            this.cancelBtn.TabIndex = 2;
            this.cancelBtn.Text = "Cancel";
            this.cancelBtn.UseVisualStyleBackColor = true;
            this.cancelBtn.Click += new System.EventHandler(this.cancelBtn_Click);
            // 
            // Comment2
            // 
            this.Comment2.AutoSize = true;
            this.Comment2.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.Comment2.Location = new System.Drawing.Point(12, 32);
            this.Comment2.Margin = new System.Windows.Forms.Padding(3, 0, 3, 1);
            this.Comment2.MaximumSize = new System.Drawing.Size(598, 0);
            this.Comment2.Name = "Comment2";
            this.Comment2.Size = new System.Drawing.Size(0, 13);
            this.Comment2.TabIndex = 1;
            // 
            // ProgressDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(622, 142);
            this.Controls.Add(this.cancelBtn);
            this.Controls.Add(this.Comment2);
            this.Controls.Add(this.Comment);
            this.Controls.Add(this.progressBar1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "ProgressDialog";
            this.ShowIcon = false;
            this.Text = "Распознавание документов";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.ProgressDialog_FormClosed);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label Comment;
        private System.Windows.Forms.Button cancelBtn;
        private System.Windows.Forms.Label Comment2;
    }
}