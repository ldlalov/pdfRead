using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;


namespace pdfRead
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        int tax = 0;
        string[] settings = File.ReadAllLines("Settings.txt");
        public static string pdfText(string path)
        {
            PdfReader reader = new PdfReader(path);
            string text = string.Empty;
            for (int page = 1; page <= reader.NumberOfPages; page++)
            {
                text += PdfTextExtractor.GetTextFromPage(reader, page);
            }
            reader.Close();
            return text;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string currentLot = "";
            richTextBox1.Text = "";
            string[] files = Directory.GetFiles(settings[1]);
            int filesCount = files.Length;
            for (int i = 0; i < filesCount; i++)
            {
                int invoiceFoundindex = 0;
                string invoiceNum = "";
                string currentFile = files[i];
                string fileType = currentFile.Substring(7, 7);


                //Проверка на фактура от Елтрейд
                if (fileType == "Елтрейд")
                {
                    richTextBox1.Text += Environment.NewLine;
                    richTextBox1.Text += "------------------";
                    richTextBox1.Text += "SIM КАРТИ ЕЛТРЕЙД";
                    richTextBox1.Text += "------------------";
                    richTextBox1.Text += Environment.NewLine;
                    string readedText = pdfText(currentFile);

                    invoiceFoundindex = readedText.IndexOf("Номер на фактурата", 0);
                    if (invoiceFoundindex != -1)
                    {
                        invoiceNum = readedText.Substring(invoiceFoundindex, 29);
                    }
                    invoiceFoundindex = readedText.IndexOf("Фактура No. : ", 0);
                    if (invoiceFoundindex != -1)
                    {
                        invoiceNum = readedText.Substring(invoiceFoundindex, 24);
                    }
                    int Start = 0;
                    //Извлича номера на фактурата
                    richTextBox1.Text += invoiceNum;
                    richTextBox1.Text += Environment.NewLine;
                    for (int index = 0; index < readedText.Length; index++)
                    {
                        Start = readedText.IndexOf("ED", index);
                        if (Start != -1)
                        {
                            currentLot = readedText.Substring(Start, 8);
                            richTextBox1.Text += readedText.Substring(Start, 8);
                            richTextBox1.Text += Read_From_Database($"{currentLot}");
                            richTextBox1.Text += Environment.NewLine;
                            index = Start;
                        }
                        else
                        {
                            break;
                        }
                    }
                }


                //CSV файл от Датекс
                string ext = currentFile.Substring(currentFile.Length-4, 4);
                if (ext == ".csv")
                {
                    richTextBox1.Text += Environment.NewLine;
                    richTextBox1.Text += "------------------";
                    richTextBox1.Text += "SIM КАРТИ ОТ ДАТЕКС";
                    richTextBox1.Text += "------------------";
                    richTextBox1.Text += Environment.NewLine;
                    using (var reader = new StreamReader(currentFile))
                    {
                        //List<string> listA = new List<string>();
                        //List<string> listB = new List<string>();
                        reader.ReadLine();
                        while (!reader.EndOfStream)
                        {
                            var line = reader.ReadLine();
                            var values = line.Split(';');
                            tax = int.Parse(values[6].Substring(1,1));
                            currentLot= values[2].Substring(1, 8);
                            richTextBox1.Text += currentLot;
                            richTextBox1.Text += Read_From_Database($"{currentLot}");
                            richTextBox1.Text += Environment.NewLine;
                        }
                    }
                }


                //CSV файл от Тремол
                ext = currentFile.Substring(currentFile.Length-4, 4);
                if (ext == ".xls")
                {
                    richTextBox1.Text += Environment.NewLine;
                    richTextBox1.Text += "------------------";
                    richTextBox1.Text += "SIM КАРТИ ОТ ТРЕМОЛ";
                    richTextBox1.Text += "------------------";
                    richTextBox1.Text += Environment.NewLine;
                    using (var reader = new StreamReader(currentFile))
                    {
                        //List<string> listA = new List<string>();
                        //List<string> listB = new List<string>();
                        reader.ReadLine();
                        while (!reader.EndOfStream)
                        {
                            var line = reader.ReadLine();
                            var values = line.Split('\t');
                            currentLot= values[0].Substring(0, 8);
                            richTextBox1.Text += currentLot;
                            richTextBox1.Text += Read_From_Database($"{currentLot}");
                            richTextBox1.Text += Environment.NewLine;
                        }
                    }
                }
            }
        }
        public string Read_From_Database(string currentLot)
        {
            string lastDeliveryDate = "";
            string taxa = "";
            string Acct = "";
            string AcctTax = "";
            string sqlString = "";
            string sqlString1 = "";
            using (SqlConnection connection = new SqlConnection(
           settings[0]))
            {
                // Взема последната дата на доставка от таблица Operations
                    sqlString = ($"SELECT top(1) [Date],[Acct],[GoodID] FROM[dbo].[Operations] WHERE[Lot] = '{currentLot}' AND [OperType] = 1 AND ([GoodId] = 1330 or [GoodId] = 1823 or [GoodId] = 2443 or [GoodId] = 1179 or [GoodId] = 1329 or [GoodId] = 1351 or [GoodId] = 1374 or [GoodId] = 1375 or [GoodId] = 1376 or [GoodId] = 1377 or [GoodId] = 1393 or [GoodId] = 1401 or [GoodId] = 2098 or [GoodId] = 2268 or [GoodId] = 2269 or [GoodId] = 2270 or [GoodId] = 2271) order by ID desc");
                    sqlString1 = ($"SELECT top(1) [Date],[Acct],[GoodID] FROM[dbo].[Operations] WHERE[Lot] = '{currentLot}' AND [OperType] = 1 AND ([GoodId] = 1340 or [GoodId] = 1609) order by ID desc");
                    SqlCommand command = new SqlCommand(
                    sqlString, connection);
                try
                {
                    connection.Open();
                }
                catch
                {
                    MessageBox.Show("Няма връзка с базата данни!"/*ex.ToString()*/);
                    //connectionError = true;
                }
                finally
                {
                    
                }
                Acct = "Няма доставен договор!";
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        reader.Read();
                        lastDeliveryDate = reader[0].ToString();//Дата на последна доставка
                        Acct = $"Последна доставка: {reader[1]} / { lastDeliveryDate} ";
                    }
                }
                //Ако има такса
                if (tax > 0)
                {
                    taxa = "Има недоставена такса!";
                    SqlCommand command1 = new SqlCommand(
                    sqlString1, connection);

                    using (SqlDataReader reader = command1.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            reader.Read();
                            AcctTax = reader[1].ToString();
                            taxa = $"    такса: {AcctTax} / {reader[0]}";//Дата на последна доставка на таксата
                        }
                    }
                }
                return $"   {Acct}    {taxa}";
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Process.Start("notepad.exe", "Help.txt");
        }
    }
}
