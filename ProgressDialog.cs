using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;

namespace PilotOCR
{
    public partial class ProgressDialog : Form
    {
        public int progressValue = 0;
        public ProgressDialog()
        {
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
    }
}
