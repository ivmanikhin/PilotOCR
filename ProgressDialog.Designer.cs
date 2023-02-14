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
            this.SuspendLayout();
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(12, 33);
            this.progressBar1.Maximum = 10;
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(600, 23);
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
            // ProgressDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(624, 68);
            this.Controls.Add(this.Comment);
            this.Controls.Add(this.progressBar1);
            this.Name = "ProgressDialog";
            this.Text = "Распознавание документов";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label Comment;
    }
}