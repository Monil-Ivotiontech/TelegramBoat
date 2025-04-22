using System;
using System.IO;
using System.Text.Json;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace Json_Creator_Tool
{

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            cmbType.SelectedIndex = 0;
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            string fileName = "feed.json"; // File placed in bin\Debug\net8.0 or bin\Release\net8.0
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);

            if (File.Exists(filePath))
            {
                string content = File.ReadAllText(filePath);
                string source = cmbType.Text;
                string Date = Convert.ToDateTime(dt.Text).Day.ToString() + Convert.ToDateTime(dt.Text).ToString("MMM").ToUpper() + Convert.ToDateTime(dt.Text).Year.ToString();
                int from = Convert.ToInt32(txtFrom.Value);
                int to = Convert.ToInt32(txtTo.Value);
                int add = Convert.ToInt32(txtAdd.Value);
                string Translates = string.Empty;
                //if (source == "SENSEX")
                //{
                //Translates += "\"Source\" : \"SENSEX 15APR2025 70100 CE\",";

                for (int i = from; i <= to; i++)
                {
                    if (i != from)
                        i = i - 1;
                    Translates += "{";
                    Translates += "\"Source\":\"" + source + " " + Date + " " + i + " " + "CE\",";
                    Translates += "\"Symbol\":\"" + source + Date.Replace("2025", "").Replace("2026", "") + "\\/" + i + "\\/" + "CALL\",";
                    Translates += "\"BidMarkup\":" + "\"0\",";
                    Translates += "\"AskMarkup\":" + "\"0\",";
                    Translates += "\"Digits\":" + "\"2\"";
                    Translates += "},";

                    Translates += "{";
                    Translates += "\"Source\":\"" + source + " " + Date + " " + i + " " + "PT\",";
                    Translates += "\"Symbol\":\"" + source + Date.Replace("2025", "").Replace("2026", "") + "\\/" + i + "\\/" + "PUT\",";
                    Translates += "\"BidMarkup\":" + "\"0\",";
                    Translates += "\"AskMarkup\":" + "\"0\",";
                    Translates += "\"Digits\":" + "\"2\"";
                    Translates += "},";

                    i += add;
                }

                Translates += "{";
                Translates += "\"Source\":\"" + source + " " + Date + " " + to + " " + "CE\",";
                Translates += "\"Symbol\":\"" + source + Date.Replace("2025", "").Replace("2026", "") + "\\/" + to + "\\/" + "CALL\",";
                Translates += "\"BidMarkup\":" + "\"0\",";
                Translates += "\"AskMarkup\":" + "\"0\",";
                Translates += "\"Digits\":" + "\"2\"";
                Translates += "},";

                Translates += "{";
                Translates += "\"Source\":\"" + source + " " + Date + " " + to + " " + "PT\",";
                Translates += "\"Symbol\":\"" + source + Date.Replace("2025", "").Replace("2026", "") + "\\/" + to + "\\/" + "PUT\",";
                Translates += "\"BidMarkup\":" + "\"0\",";
                Translates += "\"AskMarkup\":" + "\"0\",";
                Translates += "\"Digits\":" + "\"2\"";
                Translates += "},";

                // }
                Translates = Translates.Remove(Translates.Length - 1);
                content = content.Replace("###", Translates);
                //string jsonString = JsonSerializer.Serialize(content, new JsonSerializerOptions { WriteIndented = true });

                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Title = "Save JSON File";
                    saveFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*";
                    saveFileDialog.FileName = "data.json";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        File.WriteAllText(saveFileDialog.FileName, content);
                        MessageBox.Show("File saved successfully at:\n" + saveFileDialog.FileName);
                    }
                }

            }
            else
            {
                Console.WriteLine("File not found: " + filePath);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            cmbType.SelectedIndex = 0;
            txtAdd.Value = 0m;
            txtFrom.Value = 0m;
            txtTo.Value = 0m;
            dt.Value = DateTime.Now.Date;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
