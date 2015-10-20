using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SeqLabelsBarCode
{
    public partial class Form1 : Form
    {

        private char _STX = '\x02';
        private string _AUTO_GAP = "C0000"; //Gap Automático
        private string _MANUAL_GAP = "KI503"; //Gap Manual
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
            foreach (string printersList in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            {
                comboBox1.Items.Add(printersList);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Int32 initial = Int32.Parse(textBox1.Text);
            Int32 final = Int32.Parse(textBox2.Text);

            if ((textBox1.Text != "") && (textBox2.Text != ""))
            {
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
                    listValues.Add("1E0002000010020" + i.ToString());
                    listValues.Add("Q1");
                    listValues.Add("E");

                    using (System.IO.StreamWriter file =
                    new System.IO.StreamWriter(@"C:\Users\Public\WriteLines2.txt", true))
                    {
                        foreach (string values in listValues)
                        {
                            file.WriteLine(values);
                            System.Console.WriteLine(values);
                        }
                    }
                }

                MessageBox.Show("Ok", "Teste", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Os valores não podem ser nulos", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


    }
}
