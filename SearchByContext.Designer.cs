namespace PilotOCR
{
    partial class SearchByContext
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
            this.SearchTextBox = new System.Windows.Forms.TextBox();
            this.RadioButtonAnd = new System.Windows.Forms.RadioButton();
            this.RadioButtonOr = new System.Windows.Forms.RadioButton();
            this.ResultTextBox = new System.Windows.Forms.TextBox();
            this.SearchButton = new System.Windows.Forms.Button();
            this.СheckBoxInbox = new System.Windows.Forms.CheckBox();
            this.CheckBoxSent = new System.Windows.Forms.CheckBox();
            this.GroupBoxSearchMethod = new System.Windows.Forms.GroupBox();
            this.GroupBoxSearchMethod.SuspendLayout();
            this.SuspendLayout();
            // 
            // SearchTextBox
            // 
            this.SearchTextBox.AcceptsReturn = true;
            this.SearchTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.SearchTextBox.Location = new System.Drawing.Point(17, 16);
            this.SearchTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.SearchTextBox.Multiline = true;
            this.SearchTextBox.Name = "SearchTextBox";
            this.SearchTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.SearchTextBox.Size = new System.Drawing.Size(480, 517);
            this.SearchTextBox.TabIndex = 0;
            // 
            // RadioButtonAnd
            // 
            this.RadioButtonAnd.AutoSize = true;
            this.RadioButtonAnd.Location = new System.Drawing.Point(6, 21);
            this.RadioButtonAnd.Name = "RadioButtonAnd";
            this.RadioButtonAnd.Size = new System.Drawing.Size(88, 20);
            this.RadioButtonAnd.TabIndex = 1;
            this.RadioButtonAnd.Text = "Поиск \"И\"";
            this.RadioButtonAnd.UseVisualStyleBackColor = true;
            // 
            // RadioButtonOr
            // 
            this.RadioButtonOr.AutoSize = true;
            this.RadioButtonOr.Checked = true;
            this.RadioButtonOr.Location = new System.Drawing.Point(100, 21);
            this.RadioButtonOr.Name = "RadioButtonOr";
            this.RadioButtonOr.Size = new System.Drawing.Size(107, 20);
            this.RadioButtonOr.TabIndex = 1;
            this.RadioButtonOr.TabStop = true;
            this.RadioButtonOr.Text = "Поиск \"ИЛИ\"";
            this.RadioButtonOr.UseVisualStyleBackColor = true;
            // 
            // ResultTextBox
            // 
            this.ResultTextBox.AcceptsReturn = true;
            this.ResultTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.ResultTextBox.Location = new System.Drawing.Point(505, 16);
            this.ResultTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.ResultTextBox.Multiline = true;
            this.ResultTextBox.Name = "ResultTextBox";
            this.ResultTextBox.ReadOnly = true;
            this.ResultTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.ResultTextBox.Size = new System.Drawing.Size(666, 570);
            this.ResultTextBox.TabIndex = 0;
            // 
            // SearchButton
            // 
            this.SearchButton.Location = new System.Drawing.Point(422, 558);
            this.SearchButton.Name = "SearchButton";
            this.SearchButton.Size = new System.Drawing.Size(75, 28);
            this.SearchButton.TabIndex = 2;
            this.SearchButton.Text = "Поиск";
            this.SearchButton.UseVisualStyleBackColor = true;
            this.SearchButton.Click += new System.EventHandler(this.SearchButton_Click);
            // 
            // СheckBoxInbox
            // 
            this.СheckBoxInbox.AutoSize = true;
            this.СheckBoxInbox.Checked = true;
            this.СheckBoxInbox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.СheckBoxInbox.Location = new System.Drawing.Point(242, 566);
            this.СheckBoxInbox.Name = "СheckBoxInbox";
            this.СheckBoxInbox.Size = new System.Drawing.Size(89, 20);
            this.СheckBoxInbox.TabIndex = 3;
            this.СheckBoxInbox.Text = "Входящие";
            this.СheckBoxInbox.UseVisualStyleBackColor = true;
            // 
            // CheckBoxSent
            // 
            this.CheckBoxSent.AutoSize = true;
            this.CheckBoxSent.Checked = true;
            this.CheckBoxSent.CheckState = System.Windows.Forms.CheckState.Checked;
            this.CheckBoxSent.Location = new System.Drawing.Point(242, 540);
            this.CheckBoxSent.Name = "CheckBoxSent";
            this.CheckBoxSent.Size = new System.Drawing.Size(97, 20);
            this.CheckBoxSent.TabIndex = 3;
            this.CheckBoxSent.Text = "Исходящие";
            this.CheckBoxSent.UseVisualStyleBackColor = true;
            // 
            // GroupBoxSearchMethod
            // 
            this.GroupBoxSearchMethod.Controls.Add(this.RadioButtonAnd);
            this.GroupBoxSearchMethod.Controls.Add(this.RadioButtonOr);
            this.GroupBoxSearchMethod.Location = new System.Drawing.Point(12, 537);
            this.GroupBoxSearchMethod.Name = "GroupBoxSearchMethod";
            this.GroupBoxSearchMethod.Size = new System.Drawing.Size(213, 49);
            this.GroupBoxSearchMethod.TabIndex = 4;
            this.GroupBoxSearchMethod.TabStop = false;
            this.GroupBoxSearchMethod.Text = "Тип поиска:";
            // 
            // SearchByContext
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1180, 589);
            this.Controls.Add(this.GroupBoxSearchMethod);
            this.Controls.Add(this.CheckBoxSent);
            this.Controls.Add(this.СheckBoxInbox);
            this.Controls.Add(this.SearchButton);
            this.Controls.Add(this.ResultTextBox);
            this.Controls.Add(this.SearchTextBox);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.Name = "SearchByContext";
            this.ShowIcon = false;
            this.Text = "Поиск по тексту писем";
            this.GroupBoxSearchMethod.ResumeLayout(false);
            this.GroupBoxSearchMethod.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox SearchTextBox;
        private System.Windows.Forms.RadioButton RadioButtonAnd;
        private System.Windows.Forms.RadioButton RadioButtonOr;
        private System.Windows.Forms.TextBox ResultTextBox;
        private System.Windows.Forms.Button SearchButton;
        private System.Windows.Forms.CheckBox СheckBoxInbox;
        private System.Windows.Forms.CheckBox CheckBoxSent;
        private System.Windows.Forms.GroupBox GroupBoxSearchMethod;
    }
}