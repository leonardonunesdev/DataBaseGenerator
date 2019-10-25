using DataBaseGenerator.Model;
using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace DataBaseGenerator.Controller
{
    class Functions
    {
        public bool randomBoolean()
        {
            System.Threading.Thread.Sleep(1);
            Random random = new Random();
            bool binario = random.Next(100) < 50;

            return binario;
        }

        public int randomInt()
        {
            System.Threading.Thread.Sleep(1);
            Random random = new Random();

            return random.Next(100);
        }

        public double randomDouble()
        {
            Random random = new Random();
            var dif = Math.Abs(36.5 - 40.0);

            return Math.Round(random.NextDouble() * dif + 36.5, 2);
        }

        public void generateExcel(List<Row> rows)
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            

            xlWorkSheet.Cells[1, 1] = "Nome";
            xlWorkSheet.Cells[1, 2] = "Febre";
            xlWorkSheet.Cells[1, 3] = "Tosse";
            xlWorkSheet.Cells[1, 4] = "Falta ar e dificuldade respirar";
            xlWorkSheet.Cells[1, 5] = "Dor";
            xlWorkSheet.Cells[1, 6] = "Mal-estar generalizado";
            xlWorkSheet.Cells[1, 7] = "Fraqueza";
            xlWorkSheet.Cells[1, 8] = "Suor intenso";
            xlWorkSheet.Cells[1, 9] = "Nausea e Vomito";
            xlWorkSheet.Cells[1, 10] = "Pneumonia";

            int numRow = 2;

            foreach(Row row in rows)
            {
                xlWorkSheet.Cells[numRow, 1] = row.nome;
                xlWorkSheet.Cells[numRow, 2] = row.febre;
                xlWorkSheet.Cells[numRow, 3] = row.tosse;
                xlWorkSheet.Cells[numRow, 4] = row.faltaArEDificuldadeRespirar;
                xlWorkSheet.Cells[numRow, 5] = row.dor;
                xlWorkSheet.Cells[numRow, 6] = row.malEstarGeneralizado;
                xlWorkSheet.Cells[numRow, 7] = row.fraqueza;
                xlWorkSheet.Cells[numRow, 8] = row.suorInteso;
                xlWorkSheet.Cells[numRow, 9] = row.nauseaEVomito;
                xlWorkSheet.Cells[numRow, 10] = row.pneumonia;

                numRow ++;
                Console.WriteLine("Número rows:" + numRow.ToString());
            }

            xlWorkBook.SaveAs("C:\\Users\\Leonardo dos Santos\\Desktop\\Minerador de dados\\Base de dados.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
