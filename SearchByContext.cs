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
        //GUI для поиска по тексту писем и вложений
    {
        //словарь замен для метода ConvertToRegex:
        private readonly Dictionary<char, string> REPLACEMENTS = new Dictionary<char, string> { { 'б', "[бБ6]" },
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
                                                                                                { '@', "[О0@]" },
                                                                                                { '(', "[(]" },
                                                                                                { ')', "[)]" } };

        //чтение настроек подключения к базе данных:
        private readonly string connectionParameters = System.IO.File.ReadAllText($"{Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)}\\ASCON\\Pilot-ICE Enterprise\\PilotOCR\\connection_settings.txt");
        
        public SearchByContext()
        {
            InitializeComponent();
        }

        private string ConvertToRegex(string input, Dictionary<char, string> replacements)
            //метод, конвертирующий поисковый запрос в regex выражение для компенсации погрешностей распознавания
        {
            string safeInput = MySqlHelper.EscapeString(input);
            if (replacements == null) return safeInput;
            if (input == null) return "";
            string result = "";
            foreach (char c in safeInput)
            {
                if (replacements.Keys.ToList().Contains(c))
                    result += replacements[c];
                else result += c;
            }
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
            //составленеие списка искомых фраз:
            List<string> searchList = SearchTextBox.Text.Split("\r\n".ToCharArray()).ToList();
            //определение типа поиска – "И" / "ИЛИ":
            if (RadioButtonOr.Checked)
                searchMethod = "OR";
            else
                searchMethod = "AND";
            //составление списка условий для SQL запроса: 
            foreach (string s in searchList)
            {
                if (s.Length > 0)
                    searchConditions += $"(text regexp '{ConvertToRegex(s, REPLACEMENTS)}') {searchMethod} ";
            }
            //составление SQL запроса:
            string commandText = $"select letter_counter, out_no, date, subject from pilotsql.inbox where {searchConditions.Remove(searchConditions.Length - 4)}";
            connection.Open();
            //поиск во входящих:
            if (searchInInbox)
            {
                using (var command = new MySqlCommand(commandText, connection))
                {
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            results.Add($"{reader.GetString(0)} - {reader.GetString(1)} - {reader.GetString(2)} - {reader.GetString(3)}");
                        }
                    }
                };
            }
            //поиск в исходящих:
            if (searchInSent)
            {
                using (var command = new MySqlCommand($"select letter_counter, out_no, date, subject from pilotsql.sent where {searchConditions.Remove(searchConditions.Length - 4)}", connection))
                {
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
            //отображение результатов поиска:
            foreach (string s in results)
                ResultTextBox.Text += (s + "\r\n\r\n");

            //TODO 1:
            //Показывать, в каком именно вложении искомая фраза.
            
            //TODO 2:
            //Вариант 1 (простой): сделать результаты поиска ссылками;
            //Вариант 2 (сложный): сделать результаты поиска кнопками, открывающими письмо в новой вкладке и конктретное вложение, содержащее искомую фразу.

        }

    }
}
