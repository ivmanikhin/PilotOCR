using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace PilotOCR
{
    public partial class SearchByContext : Form
    {
        private readonly string connectionParameters = System.IO.File.ReadAllText($"{Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)}\\ASCON\\Pilot-ICE Enterprise\\PilotOCR\\connection_settings.txt");
        public SearchByContext()
        {
            InitializeComponent();
        }

        private void SearchButton_Click(object sender, EventArgs e)
        {
            if (SearchTextBox.Text == null)
                return;
            bool searchInInbox = this.СheckBoxInbox.Checked;
            bool searchInSent = this.CheckBoxSent.Checked;
            string searchMethod;
            List<string> results = new List<string>();
            string searchConditions = "";
            MySqlConnection connection = new MySqlConnection(connectionParameters);
            List<string> searchList = SearchTextBox.Text.Split("\r\n".ToCharArray()).ToList();
            if (RadioButtonOr.Checked)
                searchMethod = "OR";
            else
                searchMethod = "AND";
            foreach (string s in searchList)
            {
                if (s.Length > 0)
                    searchConditions += $"(text like '%{MySqlHelper.EscapeString(s)}%') {searchMethod} ";
            }
            string commandText = $"select letter_counter, out_no, date, subject from pilotsql.inbox where {searchConditions.Remove(searchConditions.Length - 4)}";
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
                using (var command = new MySqlCommand($"select letter_counter, out_no, date, subject from pilotsql.sent where {searchConditions.Remove(searchConditions.Length - 4)}", connection))
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
