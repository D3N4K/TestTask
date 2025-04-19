using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace TestTskApp.Classes
{   
    internal class InfoObject
    {
        public string Name { get; set; }
        public float Distance { get; set; }
        public float Angle { get; set; }
        public float Height { get; set; }
        public float Width { get; set; }
        public string IsDefect { get; set; }
    public InfoObject(string name, float distance, float angle, float height, float width, string isDefect)
        {
            Name = name;
            Distance = distance;
            Angle = angle;
            Height = height;
            Width = width;
            IsDefect = isDefect;
        }
        public InfoObject()
        {
        }
        public List<InfoObject> ImportExcel(string filePath)
        {
            List<InfoObject> infoObjects = new List<InfoObject>();
            var app = new Excel.Application();
            try
            {
                Excel.Workbook workbook = app.Workbooks.Open(filePath);
                Excel.Worksheet worksheet = workbook.Sheets[1];
                int maxRow = worksheet.UsedRange.Rows.Count;
                for(int row = 2; row <= maxRow; row++)
                {
                    infoObjects.Add(new InfoObject(worksheet.Cells[row, 1].Text, float.Parse(worksheet.Cells[row, 2].Text), float.Parse(worksheet.Cells[row, 3].Text), float.Parse(worksheet.Cells[row, 4].Text), float.Parse(worksheet.Cells[row, 5].Text), worksheet.Cells[row, 6].Text));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                app.Quit();
            }
            return infoObjects;
        }
        public List<InfoObject> ImportCSV(string filePath)
        {
            List<InfoObject> infoObjects = new List<InfoObject>();
            try
            {
                StreamReader reader = new StreamReader(filePath);
                var headers = reader.ReadLine();
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(';');
                    infoObjects.Add(new InfoObject(values[0].ToString(), float.Parse(values[1].ToString()), float.Parse(values[2].ToString()), float.Parse(values[3].ToString()), float.Parse(values[4].ToString()), values[5].ToString()));
                }
                reader.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return infoObjects;
        }
    }
}