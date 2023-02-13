using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PilotOCR
{
    public partial class ProgressDialog : Form
    {
        public ProgressDialog()
        {
            InitializeComponent();
        }

        public void RefreshProgress()
        {
            progressBar1.Value++;
            Comment.Text = "Распознано " + progressBar1.Value.ToString() + " документов из " + progressBar1.Maximum.ToString();
        }
        public void SetMax(int maxValue)
        {
            progressBar1.Maximum = maxValue;
            Comment.Text = "Распознано 0 документов из " + maxValue.ToString();
        }

    }
}
