using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace BicycleRace
{
    class ExcelClass
    {
        private int wsNumber;
        private string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public ExcelClass()
        { }

        public ExcelClass(string path, int sheet)
        {
            this.path = path;
            this.wb = excel.Workbooks.Open(path);
            this.ws = wb.Worksheets[sheet];
        }

        public void CreateNewFile()
        {
            this.wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            this.ws = wb.Worksheets[1];
        }

        public void CreateNewSheet()
        {
            wsNumber = wb.Worksheets.Count;
            this.ws = wb.Worksheets[wsNumber];
            Worksheet sheet = wb.Worksheets.Add(After: ws);
            this.ws = wb.Worksheets[1];
        }

        public void SelectWorksheet(int SheetNumber)
        {
            this.ws = wb.Worksheets[SheetNumber];
        }

        public string ReadCell(int i, int j)
        {
            i++;
            j++;
            if (ws.Cells[i, j].Value2 != null)
                return ws.Cells[i, j].Value2.ToString();
            else
                return "Empty cell";
         }

        public void WriteToExcel(int i, int j, string text) 
        {
            i++;
            j++;
            ws.Cells[i, j].Value2 = text;
        }

        /*public void WriteToExcel(int j)
        {
            j++;
            for(int i = 2; i < 10; i++)
            {
                ws.Cells[i, j].Value2= i + j;
            }
        }*/

        public void SaveAs(string path)
        {
            wb.SaveAs(path);
        }


        public void Save()
        {
            wb.Save();
        }

        public void Close()
        {
            wb.Close();
        }

        public void ProtectSheet(string Password)       //az Excel mentés után csak jelszóval feloldás után módosítható
        {
            ws.Protect(Password);
        }
    }
}
