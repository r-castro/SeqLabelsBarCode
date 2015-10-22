using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace SeqLabelsBarCode
{
    public partial class Form1 : Form
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern SafeFileHandle CreateFile(string lpFileName, FileAccess dwDesiredAccess,
            uint dwShareMode, IntPtr lpSecurityAttributes, FileMode dwCreationDisposition,
                uint dwFlagsAndAttributes, IntPtr hTemplateFile);

        private char _STX = '\x02';
        private string _AUTO_GAP = "C0000"; //Gap Automático
        //private string _MANUAL_GAP = "KI503"; //Gap Manual
        private string _AUTO_BACKFEED = "f340"; //BackFeed Automático
        private string _PARAMETER_1 = "00220";
        private string _PARAMETER_2 = "KI7";
        private string _PARAMETER_3 = "V0";
        private string _PARAMETER_4 = "L";
        private string _TEMPERATURE = "H12"; //Parametro de ajuste de temperatura (H10 --> H20)
        private string _PARAMETER_5 = "PA";
        private string _PARAMETER_6 = "A2";
        private string _PARAMETER_7 = "D11";

        public Form1()
        {
            InitializeComponent();
            PrinterList();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void PrinterList()
        {
            string hostName = System.Net.Dns.GetHostName();
            comboBox1.Items.Add(@"\\" + hostName + @"\Argox");

            foreach (string printersList in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            {
                comboBox1.Items.Add(printersList);
            }
            comboBox1.SelectedIndex = 0;           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string file = @"PrintText.txt";
            string printerValue = comboBox1.Text;
            

            if ((textBox1.Text != "") && (textBox2.Text != ""))
            {
                Int32 initial = Int32.Parse(textBox1.Text);
                Int32 final = Int32.Parse(textBox2.Text);

                for (int i = initial; i <= final; i++)
                {
                    List<string> listValues = new List<string>();

                    listValues.Add(_STX + _AUTO_GAP);
                    listValues.Add(_STX + _AUTO_BACKFEED);
                    listValues.Add(_STX + _PARAMETER_1);
                    listValues.Add(_STX + _PARAMETER_2);
                    listValues.Add(_STX + _PARAMETER_3);
                    listValues.Add(_STX + _PARAMETER_4);
                    listValues.Add(_TEMPERATURE);
                    listValues.Add(_PARAMETER_5);
                    listValues.Add(_PARAMETER_6);
                    listValues.Add(_PARAMETER_7);
                    listValues.Add("1E0003000010010A" + i.ToString());
                    listValues.Add("Q1");
                    listValues.Add("E");

                    try
                    {
                        StreamWriter writer = new StreamWriter(file);
                        foreach (string values in listValues)
                        {
                            writer.WriteLine(values);
                        }

                        writer.Flush();
                        writer.Close();

                        File.Copy(file, printerValue);
                    }
                    catch (Exception ex)
                    {
                        ex.ToString();
                    }
                }
            }
            else
            {
                MessageBox.Show("Os valores não podem ser nulos", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
