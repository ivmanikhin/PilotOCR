using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace PilotOCR
{
    public partial class SearchByContext : Form
    {
        private const string CONNECTION_PARAMETERS = "datasource=localhost;port=3306;username=root;password=C@L0P$Ck;charset=utf8";
        public SearchByContext()
        {
            InitializeComponent();
        }

        private void SearchButton_Click(object sender, EventArgs e)
        {
            if (SearchTextBox.Text == null)
                return;
            bool searchInInbox = true;
            bool searchInSent = true;
            string searchMethod = "OR";
            List<string> results = new List<string>();
            string searchConditions = "";
            MySqlConnection connection = new MySqlConnection(CONNECTION_PARAMETERS);
            List<string> searchList = SearchTextBox.Text.Split("\r\n".ToCharArray()).ToList();
            foreach (string s in searchList)
            {
                if (s.Length > 0)
                    searchConditions += $"(text like '%{MySqlHelper.EscapeString(s)}%') {searchMethod} ";
            }
            string commandText = $"select letter_counter, out_no, date, subject from pilotsql.inbox where {searchConditions.Remove(searchConditions.Length - 3)}";
            connection.Open();
            if (searchInInbox)
            {
                using (var command = new MySqlCommand(commandText, connection))
                {
                    //Debug.WriteLine(command.CommandText);
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            results.Add($"{reader.GetString(0)} - {reader.GetString(1)} - {reader.GetString(2)} - {reader.GetString(3)}");
                        }
                    }
                };
            }
            if (searchInSent)
            {
                using (var command = new MySqlCommand($"select letter_counter, out_no, date, subject from pilotsql.sent where {searchConditions.Remove(searchConditions.Length - 3)}", connection))
                {
                    //Debug.WriteLine(command.CommandText);
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            results.Add($"{reader.GetString(0)} - {reader.GetString(1)} - {reader.GetString(2)} - {reader.GetString(3)}");
                        }
                    }
                }
            }
            connection.Close();
            ResultTextBox.Text = "";
            foreach (string s in results)
                ResultTextBox.Text += (s + "\r\n\r\n");
        }
    }
}
