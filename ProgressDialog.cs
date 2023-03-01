using System;
using System.Threading;
using System.Windows.Forms;

namespace PilotOCR
{
    public partial class ProgressDialog : Form
    {
        private static ModifyObjectsPlugin _mainClass;

        public int progressValue = 0;
        public ProgressDialog(ModifyObjectsPlugin mainClass)
        {
            _mainClass = mainClass;
            InitializeComponent();
        }
        public void UpdateProgress()
        {
            new Thread(() =>
            {
                Invoke((Action)(() =>
                {
                    progressBar1.Value++;
                    Comment.Text = "Распознано " + progressBar1.Value.ToString() + " документов из " + progressBar1.Maximum.ToString();
                }));
            }).Start();
        }



        public void SetMax(int maxValue)
        {
            progressBar1.Maximum = maxValue;
            Comment.Text = "Распознано 0 документов из " + maxValue.ToString();
        }

        private void cancelBtn_Click(object sender, EventArgs e)
        {
            _mainClass.KillThemAll();
            this.Close();
        }

        public void CloseRemotely()
        {
            new Thread(() =>
            {
                Invoke((Action)(() => Close()));
            }).Start();
        }
    }
}
