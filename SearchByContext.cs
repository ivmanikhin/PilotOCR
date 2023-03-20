using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
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

        private string ConvertToRegex(string input)
        {
            var replacements = new Dictionary<char, string> { { 'б', "[бБ6]" },
                                                              { 'Б', "[бБ6]" },
                                                              { '6', "[бБ6]" },
                                                              { 'В', "[В8]" },
                                                              { '8', "[В8]" },
                                                              { 'З', "[З3]" },
                                                              { '3', "[З3]" },
                                                              { 'и', "[инп]" },
                                                              { 'н', "[инп]" },
                                                              { 'п', "[инпл]" },
                                                              { 'л', "[лп]" },
                                                              { 'О', "[О0@]" },
                                                              { '0', "[О0@]" },
                                                              { '@', "[О0@]" } };
            string result = "";
            foreach (char c in input)
            {
                if (replacements.Keys.ToList().Contains(c))
                    result += replacements[c];
                else result += c;
            }


            Debug.WriteLine(result);
            return result;
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
                    //searchConditions += $"(text like '%{MySqlHelper.EscapeString(s)}%') {searchMethod} ";
                    searchConditions += $"(text regexp '{ConvertToRegex(s)}') {searchMethod} ";
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
