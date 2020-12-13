using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using System.Text.RegularExpressions;

namespace Beadando_projekt
{
    public partial class Form1 : Form
    {
        Excel.Application xlApp; 
        Excel.Workbook xlWB; 
        Excel.Worksheet xlSheet;

        List<Fire> _fires = new List<Fire>();

       

        public Form1()
        {
            InitializeComponent();            
            
            
            
        }

        private string GetCell(int x, int y)
        {
            string ExcelCoordinate = "";
            int dividend = y;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                ExcelCoordinate = Convert.ToChar(65 + modulo).ToString() + ExcelCoordinate;
                dividend = (int)((dividend - modulo) / 26);
            }
            ExcelCoordinate += x.ToString();

            return ExcelCoordinate;
        }

        private void LoadFires()
        {
            _fires.Clear();

            using(StreamReader sr = new StreamReader("top_20_CA_wildfires.csv", Encoding.Default))
            {
                sr.ReadLine();
                while (!sr.EndOfStream)
                {
                    string[] line = sr.ReadLine().Split(',');

                    Fire f = new Fire();

                    f.Name = line[0];
                    f.Cause = line[1];
                    f.Month = line[2];
                    f.Year = int.Parse(line[3]);
                    f.Country = line[4];
                    f.Acres = int.Parse(line[5]);
                    f.Structures = int.Parse(line[6]);
                    f.Deaths = int.Parse(line[7]);
                    _fires.Add(f);                    
                }
            }
        }

        private void CreateExcel()
        {
            try
            {
                
                xlApp = new Excel.Application();
                
                xlWB = xlApp.Workbooks.Add(Missing.Value);
               
                xlSheet = xlWB.ActiveSheet;
               
               CreateTable(); 

                
                xlApp.Visible = true;
                xlApp.UserControl = true;
            }
            catch (Exception ex) 
            {
                string errMsg = string.Format("Error: {0}\nLine: {1}", ex.Message, ex.Source);
                MessageBox.Show(errMsg, "Error");

               
                xlWB.Close(false, Type.Missing, Type.Missing);
                xlApp.Quit();
                xlWB = null;
                xlApp = null;
            }
        }

        private void CreateTable()
        {
            string[] headers = new string[]
            {
                "Tűz neve",
                "Oka",
                "Hónap",
                "Év",
                "Város",
                "Terület (holdban)",
                "Lerombolt épületek (db)",
                "Halálok száma (db)",
            };

            for (int i = 0; i < headers.Length; i++)
            {
                xlSheet.Cells[1, i + 1] = headers[i];
            }

            object[,] values = new object[_fires.Count, headers.Length];

            int counter = 0;

            foreach (Fire f in _fires)
            {
                values[counter, 0] = f.Name;
                values[counter, 1] = f.Cause;
                values[counter, 2] = f.Month;
                values[counter, 3] = f.Year;
                values[counter, 4] = f.Country;
                values[counter, 5] = f.Acres;
                values[counter, 6] = f.Structures;
                values[counter, 7] = f.Deaths;
                counter++;

                xlSheet.get_Range(
               GetCell(2, 1),
               GetCell(1 + values.GetLength(0), values.GetLength(1))).Value2 = values;
               FormatTable(headers.Length);
            }            
        }

        private void FormatTable(int header)
        {
            Excel.Range headerRange = xlSheet.get_Range(GetCell(1, 1), GetCell(1, header));
            headerRange.Font.Bold = true;
            headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            headerRange.EntireColumn.AutoFit();
            headerRange.RowHeight = 40;
            headerRange.Interior.Color = Color.LightBlue;
            headerRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);
            headerRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            int lastRowID = xlSheet.UsedRange.Rows.Count;
            Excel.Range tableRange = xlSheet.get_Range(GetCell(1, 1), GetCell(lastRowID, header));
            tableRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);
            tableRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            Excel.Range firstColumn = xlSheet.get_Range(GetCell(1, 1), GetCell(lastRowID, 1));
            firstColumn.Font.Bold = true;
            firstColumn.Interior.Color = Color.Orange;
            Excel.Range lastColumn = xlSheet.get_Range(GetCell(1, header), GetCell(lastRowID, header));
            lastColumn.Interior.Color = Color.Red;
            lastColumn.NumberFormat = "#,##0.00";


        }

        private void btnDone_Click(object sender, EventArgs e)
        {

            LoadFires();
            CreateExcel();
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            CleanForm(this.Controls);
            
        }
        
        private void CleanForm(Control.ControlCollection ctrl)
        {
            foreach (Control c in ctrl)
            {
                if (c is TextBox)
                {
                    c.ResetText();
                    c.BackColor = Color.Empty;
                }
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            string[] datas = new string[typeof(Fire).GetProperties().Count()];

            datas[0] = txtName.Text;
            datas[1] = txtCause.Text;
            datas[2] = txtMonth.Text;
            datas[3] = txtYear.Text;
            datas[4] = txtCountry.Text;
            datas[5] = txtAcres.Text;
            datas[6] = txtStructures.Text;
            datas[7] = txtDeaths.Text;

            using (StreamWriter sw = new StreamWriter("top_20_CA_wildfires.csv", true, Encoding.UTF8))
            {
                sw.WriteLine();

                for (int i = 0; i < datas.Length; i++)
                {
                    sw.Write(datas[i]);
                    sw.Write(',');
                }
            }

            CleanForm(this.Controls);
        }

        private bool ValidField(string txt)
        {
            return !string.IsNullOrEmpty(txt);
        }
        private bool ValidString(string txt)
        {
            Regex r = new Regex(@"^[A-Za-záéiíoóöőuúüűÁÉIÍOÓÖŐUÚÜŰ\s]*$");
            return r.IsMatch(txt);
        }
        private bool ValidNumber(string nb)
        {
            Regex r = new Regex(@"^[0-9]*$");
            return r.IsMatch(nb);
        }

        private void txtName_Validating(object sender, CancelEventArgs e)
        {
            if (!ValidField(txtName.Text))
            {
                e.Cancel = true;
                errorProvider1.SetError(txtName, "Nem lehet üres a mező!");
            }
            if (!ValidString(txtName.Text))
            {
                e.Cancel = true;
                errorProvider1.SetError(txtName, "Csak betűket adjon meg!");
            }
        }

        private void txtName_Validated(object sender, EventArgs e)
        {
            errorProvider1.SetError(txtName, "");
            txtName.BackColor = Color.Green;
        }

        private void txtCause_Validating(object sender, CancelEventArgs e)
        {
            if (!ValidField(txtCause.Text))
            {
                e.Cancel = true;
                errorProvider1.SetError(txtCause, "Nem lehet üres a mező!");
            }
            if (!ValidString(txtCause.Text))
            {
                e.Cancel = true;
                errorProvider1.SetError(txtCause, "Csak betűket adjon meg!");
            }
        }

        private void txtCause_Validated(object sender, EventArgs e)
        {
            errorProvider1.SetError(txtCause, "");
            txtCause.BackColor = Color.Green;
        }

        private void txtMonth_Validating(object sender, CancelEventArgs e)
        {
            if (!ValidField(txtMonth.Text))
            {
                e.Cancel = true;
                errorProvider1.SetError(txtMonth, "Nem lehet üres a mező!");
            }
        }

        private void txtMonth_Validated(object sender, EventArgs e)
        {
            errorProvider1.SetError(txtMonth, "");
            txtMonth.BackColor = Color.Green;
        }

        private void txtYear_Validating(object sender, CancelEventArgs e)
        {
           

            if (!ValidField(txtYear.Text))
            {
                e.Cancel = true;
                errorProvider1.SetError(txtYear, "Nem lehet üres a mező!");
            }
            if (!ValidNumber(txtYear.Text))
            {
                e.Cancel = true;
                errorProvider1.SetError(txtYear, "Csak számokat adjon meg!");               
                return;
                
            }                   
            if (int.Parse(txtYear.Text) > DateTime.Now.Year)
            {
                e.Cancel = true;
                errorProvider1.SetError(txtYear, "Ne jövőbeli dátumot adjon meg!");
            }
                               
        }

        private void txtYear_Validated(object sender, EventArgs e)
        {
            errorProvider1.SetError(txtYear, "");
            txtYear.BackColor = Color.Green;
        }

        private void txtCountry_Validating(object sender, CancelEventArgs e)
        {
            if (!ValidField(txtCountry.Text))
            {
                e.Cancel = true;
                errorProvider1.SetError(txtCountry, "Nem lehet üres a mező!");
            }
            if (!ValidString(txtCountry.Text))
            {
                e.Cancel = true;
                errorProvider1.SetError(txtCountry, "Csak betűket adjon meg!");
            }
        }

        private void txtCountry_Validated(object sender, EventArgs e)
        {
            errorProvider1.SetError(txtCountry, "");
            txtCountry.BackColor = Color.Green;
        }

        private void txtAcres_Validating(object sender, CancelEventArgs e)
        {
            if (!ValidField(txtAcres.Text))
            {
                e.Cancel = true;
                errorProvider1.SetError(txtAcres, "Nem lehet üres a mező!");
            }
            if (!ValidNumber(txtAcres.Text))
            {
                e.Cancel = true;
                errorProvider1.SetError(txtAcres, "Csak számokat adjon meg!");

            }
        }

        private void txtAcres_Validated(object sender, EventArgs e)
        {
            errorProvider1.SetError(txtAcres, "");
            txtAcres.BackColor = Color.Green;
        }

        private void txtStructures_Validating(object sender, CancelEventArgs e)
        {
            if (!ValidField(txtStructures.Text))
            {
                e.Cancel = true;
                errorProvider1.SetError(txtStructures, "Nem lehet üres a mező!");
            }
            if (!ValidNumber(txtStructures.Text))
            {
                e.Cancel = true;
                errorProvider1.SetError(txtStructures, "Csak számokat adjon meg!");

            }
        }

        private void txtStructures_Validated(object sender, EventArgs e)
        {
            errorProvider1.SetError(txtStructures, "");
            txtStructures.BackColor = Color.Green;
        }

        private void txtDeaths_Validating(object sender, CancelEventArgs e)
        {
            if (!ValidField(txtDeaths.Text))
            {
                e.Cancel = true;
                errorProvider1.SetError(txtDeaths, "Nem lehet üres a mező!");
            }
            if (!ValidNumber(txtDeaths.Text))
            {
                e.Cancel = true;
                errorProvider1.SetError(txtDeaths, "Csak számokat adjon meg!");

            }
        }

        private void txtDeaths_Validated(object sender, EventArgs e)
        {
            errorProvider1.SetError(txtDeaths, "");
            txtDeaths.BackColor = Color.Green;
        }
    }
}
