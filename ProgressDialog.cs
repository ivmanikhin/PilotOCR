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

        //обновление прогрессбара и информации о ходе распознавания писем:
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

        //обновление информации о распознаваемом в данный момент письме
        public void SetCurrentDocName(string currentDocName)
        {
            new Thread(() =>
            {
                Invoke((Action)(() =>
                {
                    Comment2.Text = currentDocName;
                }));
            }).Start();
        }


        //установка максимального значение прогрессбара = кол-ву писем:
        public void SetMax(int maxValue)
        {
            progressBar1.Maximum = maxValue;
            Comment.Text = "Распознано 0 документов из " + maxValue.ToString();
        }

        //по нажатию отмены остановить процесс распознавания:
        private void cancelBtn_Click(object sender, EventArgs e)
        {
            _mainClass.KillThemAll();
            this.Close();
        }

        //закрывание окна по сигналу основной программы:
        public void CloseRemotely()
        {
            new Thread(() =>
            {
                Invoke((Action)(() => Close()));
            }).Start();
        }

        //отмена распознавания по закрытию окна:
        private void ProgressDialog_FormClosed(object sender, FormClosedEventArgs e)
        {
            _mainClass.KillThemAll();
        }
    }
}
