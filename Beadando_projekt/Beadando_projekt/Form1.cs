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
            LoadFires();
            CreateExcel();
            
            
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
                    f.Acres = line[5];
                    f.Structures = line[6];
                    f.Deaths = line[7];
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

        public void CreateTable()
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

               
            }
            xlSheet.get_Range(
               GetCell(2, 1),
               GetCell(1 + values.GetLength(0), values.GetLength(1))).Value2 = values;


        }
    }
}
