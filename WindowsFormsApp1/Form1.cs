using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace WindowsFormsApp1
{
    public partial class InterviewerHelper : Form
    {
        public int sumpoint = 0;

        public int level1 = 0;
        public int level2 = 0;
        public int level3 = 0;
        public int level4 = 0;

        public List<string> results1 = new List<string>();
        public List<string> results2 = new List<string>();
        public List<string> results3 = new List<string>();

        public InterviewerHelper()
        {
            InitializeComponent();
            Createfolder();
            CreateQuestionsFile();
            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
             
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadComboBoxQuestions();
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text != null)
            {
                if (radioButton1.Checked || radioButton2.Checked || radioButton3.Checked || radioButton4.Checked)
                {
                    AddToReviewedQuestionList();
                    CleanRadioButtons();
                    CleanAdditionalNotes();
                    BarExample();
                }
            }
        }

        private void CleanAdditionalNotes()
        {
            richTextBox1.Text = string.Empty;
        }

        private void CleanRadioButtons()
        {
            radioButton1.Checked = false;
            radioButton2.Checked = false;
            radioButton3.Checked = false;
            radioButton4.Checked = false;
        }

        private void AddToReviewedQuestionList()
        {
                    CalculateLevel();
                    SaveResults();
                    sumpoint = sumpoint + CalculatePoint();
                    listBox1.Items.Add(comboBox1.Text);
                    comboBox1.Items.Remove(comboBox1.SelectedItem);
                    comboBox1.Text = "";               
        }

        private void CalculateLevel()
        {
            if (radioButton1.Checked)
            {
                level1 = level1 + 1;
            }
            else if (radioButton2.Checked)
            {
                level2 = level2 + 1;
            }
            else if (radioButton3.Checked)
            {
                level3 = level3 + 1;
            }
            else 
            {
                level4 = level4 + 1;
            }
        }

        private void SaveResults()
        {
            results1.Add(comboBox1.Text);
            results2.Add(CalculatePoint().ToString());
            results3.Add(richTextBox1.Text);
        }

        private int CalculatePoint()
        {
            if (radioButton1.Checked)
            {
                return  0;
            }
            else if (radioButton2.Checked)
            {
                return  30;
            }
            else if (radioButton3.Checked)
            {
                return 60;
            }
            else
            {
                return 100;
            }
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            CreatePDF();
        }

        private void CreatePDF()
        {
            string subPath = Path.Combine("C:\\", "INTERVIEW","Report_From_"+textBox1.Text+".pdf");
            System.IO.FileStream fs = new FileStream(subPath, FileMode.Create);
            // Create an instance of the document class which represents the PDF document itself.  
            Document document = new Document(PageSize.A4, 25, 25, 30, 30);
            // Create an instance to the PDF file by creating an instance of the PDF   
            // Writer class using the document and the filestrem in the constructor.  

            PdfWriter writer = PdfWriter.GetInstance(document, fs);
            
            document.Open();
            document.AddTitle("Interview Report");

            Paragraph title = new Paragraph("Interview Summary");
            document.Add(title);

            Paragraph Name = new Paragraph("Candidate name: "+textBox1.Text);
            document.Add(Name);

            Paragraph Date = new Paragraph("Interview Date: " + textBox2.Text);
            document.Add(Date);

            Paragraph Interviewer = new Paragraph("Interviewer name: " + textBox3.Text);
            document.Add(Interviewer);

            Paragraph space1 = new Paragraph("");
            document.Add(space1);

            AddGraphToPDF(document);
            PrintReportContet(document);
            // Close the document  
            document.Close();
            // Close the writer instance  
            writer.Close();
            // Always close open filehandles explicity  
            fs.Close();

            PresentReport(subPath);
        }

        private void PresentReport(string subPath)
        {
            System.Diagnostics.Process.Start(subPath);
        }

        private void AddGraphToPDF(Document doc)
        {
            using (MemoryStream memoryStream = new MemoryStream())

            {

                chart1.SaveImage(memoryStream, ChartImageFormat.Png);

                iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(memoryStream.GetBuffer());

                img.ScalePercent(100f);
                img.Alignment = iTextSharp.text.Image.ALIGN_CENTER;

                doc.Add(img);                
            }
        }

        private void PrintReportContet(Document document)
        {
            PdfPTable table = new PdfPTable(3);

            table.AddCell("Question");
            table.AddCell("Comments");
            PdfPCell cell = new PdfPCell(new Phrase("Points"));
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            table.AddCell(cell);            

            for (int i=0; i<=listBox1.Items.Count-1;i++)
            {
                table.AddCell(results1[i]);
                table.AddCell(results3[i]);
                PdfPCell cell1 = new PdfPCell(new Phrase(results2[i]));
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                table.AddCell(cell1);                                                                            
            }
            document.Add(table);
        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private void chart1_Load(object sender, EventArgs e)
        {
            BarExample(); //Show bar chart
                          
        }

        public void BarExample()
        {
            //reset your chart series and legends
            chart1.Series.Clear();
            chart1.Legends.Clear();

            //Add a new Legend(if needed) and do some formating
            chart1.Legends.Add("MyLegend");
            chart1.Legends[0].LegendStyle = LegendStyle.Table;
            chart1.Legends[0].Docking = Docking.Bottom;
            chart1.Legends[0].Alignment = StringAlignment.Center;
            chart1.Legends[0].Title = "Candidate Levels";
            chart1.Legends[0].BorderColor = Color.Black;

            //Add a new chart-series
            string seriesname = "MySeriesName";
            chart1.Series.Add(seriesname);
            //set the chart-type to "Pie"
            chart1.Series[seriesname].ChartType = SeriesChartType.Pie;

            //Add some datapoints so the series. in this case you can pass the values to this method
            chart1.Series[seriesname].Points.AddXY("Experience", level3);
            chart1.Series[seriesname].Points.AddXY("Knowledge", level2);
            chart1.Series[seriesname].Points.AddXY("Unknown", level1);
            chart1.Series[seriesname].Points.AddXY("Expertice", level4);
            //chart1.Series[seriesname].Points.AddXY("MyPointName4", 10);
        }
    }
}
