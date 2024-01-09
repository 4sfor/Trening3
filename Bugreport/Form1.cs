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
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace Bugreport
{
    public partial class Form1 : Form
    {
        private string Id="";
        private string Header = "";
        private string Step = "";
        private string ResultExpected = "";
        private string ResultActual = "";
        private string VerProduct = "";
        private string VerBroswer = "";
        private string Os = "";
        private string Device = "";
        private string Model = "";
        ToolStripLabel dateLabel;
        ToolStripLabel timeLabel;
        ToolStripLabel infoLabel;
        Timer timer;

        public Form1()
        {
            InitializeComponent();
            infoLabel = new ToolStripLabel();
            infoLabel.Text = "Текущие дата и время:";
            dateLabel = new ToolStripLabel();
            timeLabel = new ToolStripLabel();

            statusStrip1.Items.Add(infoLabel);
            statusStrip1.Items.Add(dateLabel);
            statusStrip1.Items.Add(timeLabel);

            timer = new Timer() { Interval = 1000 };
            timer.Tick += timer_Tick;
            timer.Start();
        }
        void timer_Tick(object sender, EventArgs e)
        {
            dateLabel.Text = DateTime.Now.ToLongDateString();
            timeLabel.Text = DateTime.Now.ToLongTimeString();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            // amount of padding to add
            const int padding = 3;
            // get number of lines (first line is 0, so add 1)
            //int numLines = this.textBox2.GetLineFromCharIndex(this.textBox2.TextLength) + 1;
            // get border thickness
            //int border = this.textBox2.Height - this.textBox2.ClientSize.Height;
            // set height (height of one line * number of lines + spacing)
            //this.textBox2.Height = this.textBox2.Font.Height * numLines + padding + border;
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            // amount of padding to add
            const int padding = 3;

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            // amount of padding to add
            const int padding = 3;

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            // amount of padding to add
            const int padding = 3;

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            // amount of padding to add
            const int padding = 3;

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            // amount of padding to add
            const int padding = 3;

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            // amount of padding to add
            const int padding = 3;

        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            // amount of padding to add
            const int padding = 3;

        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            // amount of padding to add
            const int padding = 3;

        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            // amount of padding to add
            const int padding = 3;

        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {

        }


        private void textBox12_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

            if (textBox1.Text == "")
            {
                MessageBox.Show("Заполните обязательное поле 'ID'");
                return;
            }
            if (textBox2.Text == "")
            {
                MessageBox.Show("Заполните обязательное поле 'Заголовок'");
                return;
            }
            if (textBox13.Text == "")
            {
                MessageBox.Show("Заполните обязательное поле 'Шаги воспроизведения'");
                return;
            }
            if (textBox5.Text == "")
            {
                MessageBox.Show("Заполните обязательное поле 'Ожидаемый результат'");
                return;
            }
            if (textBox6.Text == "")
            {
                MessageBox.Show("Заполните обязательное поле 'Фактический результат'");
                return;
            }
            if (textBox7.Text == "")
            {
                MessageBox.Show("Заполните обязательное поле 'Версия продукта'");
                return;
            }
            if (textBox10.Text == "")
            {
                MessageBox.Show("Заполните обязательное поле 'ОС'");
                return;
            }
            Id = textBox1.Text;
            Header = textBox2.Text;
            Step = textBox13.Text;
            ResultExpected = textBox5.Text;
            ResultActual = textBox6.Text;
            VerProduct = textBox7.Text;
            VerBroswer = textBox8.Text;
            Os = textBox10.Text;
            Device = textBox9.Text;
            Model = textBox11.Text;
            Form2 ft=new Form2(Id, Header, Step, ResultExpected,ResultActual, VerProduct,VerBroswer,Os,Device,Model);
            ft.Show();
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void statusStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            Form3 ft3 = new Form3();
            ft3.Show();
        }

        private void создатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (System.Windows.Forms.Control control  in this.Controls)
                if (control is TextBox)
                    ((TextBox)control).Text = null;
           // Form2 form2 = new Form2();
            

        }

        private void закрытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Вы действительно хотите выйти", "Закрытие программы", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                Application.Exit();
            }


        }

        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("Заполните обязательное поле 'ID'");
                return;
            }
            if (textBox2.Text == "")
            {
                MessageBox.Show("Заполните обязательное поле 'Заголовок'");
                return;
            }
            if (textBox13.Text == "")
            {
                MessageBox.Show("Заполните обязательное поле 'Шаги воспроизведения'");
                return;
            }
            if (textBox5.Text == "")
            {
                MessageBox.Show("Заполните обязательное поле 'Ожидаемый результат'");
                return;
            }
            if (textBox6.Text == "")
            {
                MessageBox.Show("Заполните обязательное поле 'Фактический результат'");
                return;
            }
            if (textBox7.Text == "")
            {
                MessageBox.Show("Заполните обязательное поле 'Версия продукта'");
                return;
            }
            if (textBox10.Text == "")
            {
                MessageBox.Show("Заполните обязательное поле 'ОС'");
                return;
            }
            Id = textBox1.Text;
            Header = textBox2.Text;
            Step = textBox13.Text;
            ResultExpected = textBox5.Text;
            ResultActual = textBox6.Text;
            VerProduct = textBox7.Text;
            VerBroswer = textBox8.Text;
            Os = textBox10.Text;
            Device = textBox9.Text;
            Model = textBox11.Text;
            try
            {
                using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(@"\Bugreport\отчеты\" + Id + "_" + Header + ".docx", WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());
                    string[] lines = new string[]
                    {
                        Id,
                        Header,
                        "Шаги воспроизведения: " + Step,
                        "Ожидаемый результат: " + ResultExpected,
                        "Фактический результат: " + ResultActual,
                        "Версия продукта: " + VerProduct +"    " +  " Версия браузера" + VerBroswer + "    " + "ОС: " + Os,
                        "Устройство: " + Device + "    " + Model,
                    };
                    foreach (string line in lines)
                    {
                        Paragraph para = body.AppendChild(new Paragraph());
                        Run run = para.AppendChild(new Run());
                        run.AppendChild(new Text(line));
                    }

                }

                MessageBox.Show("Отчет сформирован");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при сохранении результатов: " + ex.Message);
            }

        }
            
            
        
    }
}