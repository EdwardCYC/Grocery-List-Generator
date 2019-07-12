using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Configuration;
using System.IO;
using Font = System.Drawing.Font;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Web;
using Spire.Pdf;
using PdfDocument = Spire.Pdf.PdfDocument;
using System.Drawing.Printing;

namespace GroceryListGenerator2
{
    public partial class Form1 : Form
    {
        string[] commandStrings = new string[6];
        List<DataGridView> dgvList = new List<DataGridView>();
        string connectionString = ConfigurationManager.ConnectionStrings["FoodDB"].ConnectionString;
        DataSet foodDataSet = new DataSet();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Part of the command strings to import data from all six sheets
            commandStrings[0] = "Proteins";
            commandStrings[1] = "Vegetables (P)";
            commandStrings[2] = "Vegetables (LP)";
            commandStrings[3] = "Fruits";
            commandStrings[4] = "Condiments, Spices, Herbs";
            commandStrings[5] = "Misc";

            // Add all datagridview to a list so that they can be iterated using foreach
            dgvList.Add(dataGridView1);
            dgvList.Add(dataGridView2);
            dgvList.Add(dataGridView3);
            dgvList.Add(dataGridView4);
            dgvList.Add(dataGridView5);
            dgvList.Add(dataGridView6);

            // Get the data from Excel and display it
            GetDataFromExcel();
            FormatDataGridViews();
        }

        private void GetDataFromExcel()
        {
            for (int i = 0; i < commandStrings.Length; i++)
            {
                string commandString = "SELECT [Name (English)], [Name (Chinese)], [Add to List], Quantity, Quantifier " +
                        "FROM [" + commandStrings[i] + "$]";

                using (OleDbConnection conn = new OleDbConnection(connectionString))
                using (OleDbCommand command = new OleDbCommand(commandString, conn))
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                {
                    // Fill the DataSet with the data from the Excel spreadsheet
                    adapter.Fill(foodDataSet, "Table " + i);
                    // Display the data from the DataSet on the UI
                    dgvList[i].DataSource = foodDataSet.Tables[i];
                }
            }
        }

        private void FormatDataGridViews()
        {
            foreach (DataGridView dgv in dgvList)
            {
                // Hide the column that shows which row is being selected, and the column that shows the quantifier
                dgv.RowHeadersVisible = false;
                dgv.Columns["Quantifier"].Visible = false;
                // Centralise all text in every column, and change the font to Century 11 Point
                foreach (DataGridViewColumn column in dgv.Columns)
                {
                    column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    column.DefaultCellStyle.Font = new Font("Century", 11.25F, GraphicsUnit.Pixel);
                }
                // Change the width of the columns
                dgv.Columns[1].Width = 95;
                dgv.Columns[2].Width = 75;
                dgv.Columns[3].Width = 70;
            }
        }

        private void createListBttn_Click(object sender, EventArgs e)
        {
            GenerateGroceryList(dgvList);
        }

        private void GenerateGroceryList(List<DataGridView> dgvList)
        {
            // Empty the textboxes
            groceryListTB.Text = "";
            englishListTB.Text = "";

            // Get data from all six datagridviews and display the selected items in Chinese
            foreach (DataGridView dataGridView in dgvList)
            {
                foreach (DataGridViewRow row in dataGridView.Rows)
                {
                    if (Convert.ToBoolean(row.Cells[2].Value))
                    {
                        if (row.Cells[3].Value != System.DBNull.Value)
                        {
                            groceryListTB.Text += row.Cells[3].Value.ToString() + row.Cells[4].Value.ToString() + " "
                                                + row.Cells[1].Value.ToString() + Environment.NewLine;
                        }
                        else
                        {
                            groceryListTB.Text += "      " + row.Cells[4].Value.ToString() + " "
                                                + row.Cells[1].Value.ToString() + Environment.NewLine;
                        }

                    }
                }
            }

            // Get data from all six datagridviews and display the selected items in English
            foreach (DataGridView dataGridView in dgvList)
            {
                foreach (DataGridViewRow row in dataGridView.Rows)
                {
                    if (Convert.ToBoolean(row.Cells[2].Value))
                    {
                        if (row.Cells[3].Value != System.DBNull.Value)
                        {
                            englishListTB.Text += row.Cells[3].Value.ToString() + " "
                                                + row.Cells[0].Value.ToString() + Environment.NewLine;
                        }
                        else
                        {
                            englishListTB.Text += "    " + row.Cells[0].Value.ToString() + Environment.NewLine;
                        }
                    }
                }
            }

            if (groceryListTB.Text != String.Empty)
                groceryListTB.Text = groceryListTB.Text.Remove(groceryListTB.Text.LastIndexOf(Environment.NewLine));
            if (englishListTB.Text != String.Empty)
                englishListTB.Text = englishListTB.Text.Remove(englishListTB.Text.LastIndexOf(Environment.NewLine));
        }

        private void saveBttn_Click(object sender, EventArgs e)
        {
            SaveToPdf();

        }

        private void SaveToPdf()
        {
            if (groceryListTB.Text == String.Empty)
            {
                MessageBox.Show("No items in grocery list", "Empty List");
                return;
            }

            // Create a Document object to write to the PDF
            using (Document doc = new Document(iTextSharp.text.PageSize.LETTER, 28, 28, 16, 16))
            {
                // Create an instance of PDF Writer, write the doc object to a file using a FileStream
                if (File.Exists("Grocery List.pdf"))
                    File.Delete("Grocery List.pdf");
                PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream("Grocery List.pdf", FileMode.Create));

                // Use a font that can write Chinese text
                //string fontpath = @"C:\Users\User\source\repos\GroceryListGenerator2\GroceryListGenerator2\";
                BaseFont customFont = BaseFont.CreateFont("msjh_0.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                iTextSharp.text.Font font = new iTextSharp.text.Font(customFont, 14);

                // Write the text to the PDF document
                Paragraph listToPrint = new Paragraph(groceryListTB.Text, font);
                doc.Open();
                doc.Add(listToPrint);
            }

            CreatePrintPreview();
            printBttn.Enabled = true;
            printBttn.Visible = true;
            infoBttn.Enabled = true;
            infoBttn.Visible = true;
        }

        private void CreatePrintPreview()
        {
            // Generate a preview of the saved file
            PdfDocument pdf = new PdfDocument();
            pdf.LoadFromFile("Grocery List.pdf");
            this.printPreviewControl1.Rows = 1;
            this.printPreviewControl1.Columns = 1;
            pdf.Preview(printPreviewControl1);
        }

        private void printBttn_Click(object sender, EventArgs e)
        {
            PdfDocument pdf = new PdfDocument();
            pdf.LoadFromFile("Grocery List.pdf");

            PrintDialog dialogPrint = new PrintDialog();
            dialogPrint.AllowPrintToFile = true;
            dialogPrint.AllowSomePages = true;
            dialogPrint.PrinterSettings.MinimumPage = 1;
            dialogPrint.PrinterSettings.MaximumPage = pdf.Pages.Count;
            dialogPrint.PrinterSettings.FromPage = 1;
            dialogPrint.PrinterSettings.ToPage = pdf.Pages.Count;

            if (dialogPrint.ShowDialog() == DialogResult.OK)
            {
                pdf.PrintSettings.SelectPageRange(dialogPrint.PrinterSettings.FromPage, dialogPrint.PrinterSettings.ToPage);
                pdf.PrintSettings.PrinterName = dialogPrint.PrinterSettings.PrinterName;

                pdf.Print();
            }
        }

        private void infoBttn_Click(object sender, EventArgs e)
        {
            SavePrintingData(CountLines());
        }

        private void SavePrintingData(int lineCount)
        {
            string connectionString2 = ConfigurationManager.ConnectionStrings["PrintingData"].ConnectionString;
            string commandString = "INSERT INTO [Sheet1$] VALUES (@DateTime, @LineCount)";

            using (OleDbConnection conn = new OleDbConnection(connectionString2))
            using (OleDbCommand command = new OleDbCommand(commandString, conn))
            using (OleDbDataAdapter adapter = new OleDbDataAdapter())
            {
                command.Parameters.AddWithValue("@DateTime", DateTime.Now);
                command.Parameters.AddWithValue("@LineCount", lineCount);
                adapter.InsertCommand = command;

                conn.Open();
                adapter.InsertCommand.ExecuteNonQuery();
            }
        }

        private int CountLines()
        {
            PdfDocument document = new PdfDocument();
            document.LoadFromFile("Grocery List.pdf");

            StringBuilder builder = new StringBuilder();
            // It is assumed that we will only print single page grocery lists per errand
            builder.Append(document.Pages[0].ExtractText());
            string content = builder.ToString();
            int lineCount = content.Split('\n').Length - 1;

            //MessageBox.Show(lineCount.ToString());
            return lineCount;
        }

        private void clearBttn_Click(object sender, EventArgs e)
        {
            foreach (DataGridView dgv in dgvList)
            {
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    row.Cells[2].Value = false;
                    row.Cells[3].Value = String.Empty;
                }
            }
            groceryListTB.Text = String.Empty;
            englishListTB.Text = String.Empty;
            printPreviewControl1.Document = null;
            printPreviewControl1.Refresh();
            printBttn.Enabled = false;
            printBttn.Visible = false;
            infoBttn.Enabled = false;
            infoBttn.Visible = false;
        }
    }
}
