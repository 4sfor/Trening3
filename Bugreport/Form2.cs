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
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Office2010.Excel;

namespace Bugreport
{
    public partial class Form2 : Form
    {
        private string IdForm1 = "";
        private string HeaderForm1 = "";
        private string About = "";
        private string PreCondition = "";
        private string AfterCondition = "";
        private string StepForm1 = "";
        private string ResultExpectedForm1 = "";
        private string ResultActualForm1 = "";
        private string PathScreen = "";
        private string PathLog = "";
        private string VerProductForm1 = "";
        private string VerBroswerForm1 = "";
        private string OsForm1 = "";
        private string DeviceForm1 = "";
        private string ModelForm1 = "";
        ToolStripLabel dateLabel;
        ToolStripLabel timeLabel;
        ToolStripLabel infoLabel;
        Timer timer;
        public Form2(String Id, String Header, String Step, String ResultExpected, String ResultActual, String VerProduct, String VerBroswer, String Os, String Device, String Model)
        {
            InitializeComponent();
            IdForm1 = Id;
            HeaderForm1 = Header;
            StepForm1 = Step;
            ResultExpectedForm1 = ResultExpected;
            ResultActualForm1 = ResultActual;
            VerProductForm1 = VerProduct;
            VerBroswerForm1 = VerBroswer;
            OsForm1 = Os;
            DeviceForm1 = Device;
            ModelForm1=Model;
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
        public Form2()
        {
            InitializeComponent();
        }
        void timer_Tick(object sender, EventArgs e)
        {
            dateLabel.Text = DateTime.Now.ToLongDateString();
            timeLabel.Text = DateTime.Now.ToLongTimeString();
        }
        private static void AddImageToDocument(WordprocessingDocument wordDoc, string relationshipId)
        {
            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = 5999000L, Cy = 5792900L },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = 5990000L, Cy = 5792000L }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            // Append the reference to body, the element should be in a Run.
            wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {


            About = textBox3.Text;
            PreCondition = textBox4.Text;
            AfterCondition = textBox12.Text;

            try
            {
                using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(@"\Bugreport\отчеты\" + IdForm1 + "_" + HeaderForm1 + ".docx", WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());
                    string[] lines = new string[]
                    {
                        IdForm1,
                        HeaderForm1,
                        "Описание: " + About,
                        "Предусловие: " + PreCondition,
                        "Постусловие: " + AfterCondition,
                        "Шаги воспроизведения: " + StepForm1,
                        "Ожидаемый результат: " + ResultExpectedForm1,
                        "Фактический результат: " + ResultActualForm1,
                        "Версия продукта: " + VerProductForm1 +"    " +  " Версия браузера" + VerBroswerForm1 + "    " + "ОС: " + OsForm1,
                        "Устройство: " + DeviceForm1 + "    " + ModelForm1,
                        "-------------------------------------------------",
                        "Дополнительный материалы"
                    };
                    foreach (string line in lines)
                    {
                        Paragraph para = body.AppendChild(new Paragraph());
                        Run run = para.AppendChild(new Run());
                        run.AppendChild(new Text(line));
                    }
                    try
                    {
                        ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                        using (FileStream fs = new FileStream(PathScreen, FileMode.Open))
                        {
                            imagePart.FeedData(fs);
                        }
                        AddImageToDocument(wordDocument, mainPart.GetIdOfPart(imagePart));
                    }catch(Exception a)
                    {
                        Console.WriteLine("");
                    }

                }

                MessageBox.Show("Отчет сформирован");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при сохранении результатов: " + ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                PathScreen = openFileDialog1.FileName;

            }
            button2.Text = PathScreen;
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            Form3 ft3 = new Form3();
            ft3.Show();
        }

        private void создатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (System.Windows.Forms.Control control in this.Controls)
                if (control is TextBox)
                    ((TextBox)control).Text = null;
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

            About = textBox3.Text;
            PreCondition = textBox4.Text;
            AfterCondition = textBox12.Text;

            try
            {
                using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(@"\Bugreport\отчеты\" + IdForm1 + "_" + HeaderForm1 + ".docx", WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());
                    string[] lines = new string[]
                    {
                        IdForm1,
                        HeaderForm1,
                        "Описание: " + About,
                        "Предусловие: " + PreCondition,
                        "Постусловие: " + AfterCondition,
                        "Шаги воспроизведения: " + StepForm1,
                        "Ожидаемый результат: " + ResultExpectedForm1,
                        "Фактический результат: " + ResultActualForm1,
                        "Версия продукта: " + VerProductForm1 +"    " +  " Версия браузера" + VerBroswerForm1 + "    " + "ОС: " + OsForm1,
                        "Устройство: " + DeviceForm1 + "    " + ModelForm1,
                        "-------------------------------------------------",
                        "Дополнительный материалы"
                    };
                    foreach (string line in lines)
                    {
                        Paragraph para = body.AppendChild(new Paragraph());
                        Run run = para.AppendChild(new Run());
                        run.AppendChild(new Text(line));
                    }

                    ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                    using (FileStream fs = new FileStream(PathScreen, FileMode.Open))
                    {
                        imagePart.FeedData(fs);
                    }
                    AddImageToDocument(wordDocument, mainPart.GetIdOfPart(imagePart));
                }

                MessageBox.Show("Отчет сформирован");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при сохранении результатов: " + ex.Message);
            }
        }

        /*       private void button3_Click(object sender, EventArgs e)
                {
                    if (openFileDialog2.ShowDialog() == DialogResult.OK)
                    {
                        PathScreen = openFileDialog2.FileName;

                    }
                    button3.Text = PathScreen.Substring(0, 14) + "...";
                }*/
    }
}

