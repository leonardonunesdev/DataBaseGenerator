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
            xlWorkSheet.Cells[1, 3] = "Tosse seca";
            xlWorkSheet.Cells[1, 4] = "Tosse catarro amarelado";
            xlWorkSheet.Cells[1, 5] = "Tosse catarro esverdeado";
            xlWorkSheet.Cells[1, 6] = "Falta ar e dificuldade respirar";
            xlWorkSheet.Cells[1, 7] = "Dor peito";
            xlWorkSheet.Cells[1, 8] = "Dor torax";
            xlWorkSheet.Cells[1, 9] = "Mal-estar generalizado";
            xlWorkSheet.Cells[1, 10] = "Fraqueza";
            xlWorkSheet.Cells[1, 11] = "Suor intenso";
            xlWorkSheet.Cells[1, 12] = "Pneumonia";

            int numRow = 2;

            foreach(Row row in rows)
            {
                xlWorkSheet.Cells[numRow, 1] = row.nome;
                xlWorkSheet.Cells[numRow, 2] = row.febre;
                xlWorkSheet.Cells[numRow, 3] = row.tosseSeca;
                xlWorkSheet.Cells[numRow, 4] = row.tosseCatarroAmarelo;
                xlWorkSheet.Cells[numRow, 5] = row.tosseCatarroEsverdeado;
                xlWorkSheet.Cells[numRow, 6] = row.faltaArEDificuldadeRespirar;
                xlWorkSheet.Cells[numRow, 7] = row.dorPeito;
                xlWorkSheet.Cells[numRow, 8] = row.dorTorax;
                xlWorkSheet.Cells[numRow, 9] = row.malEstarGeneralizado;
                xlWorkSheet.Cells[numRow, 10] = row.fraqueza;
                xlWorkSheet.Cells[numRow, 11] = row.suorInteso;
                xlWorkSheet.Cells[numRow, 12] = row.pneumonia;

                numRow ++;
            }

            xlWorkBook.SaveAs("c:\\Base de dados.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }


        //public void generateExcel(List<Row> rows)
        //{
        //    System.Data.DataTable table = new System.Data.DataTable("Base de Dados");
        //    StringBuilder sb = new StringBuilder();

        //    table.Columns.Add("Nome", typeof(string));
        //    table.Columns.Add("Febre", typeof(double));
        //    table.Columns.Add("Tosse seca", typeof(string));
        //    table.Columns.Add("Tosse catarro amarelado", typeof(string));
        //    table.Columns.Add("Tosse catarro esverdeado", typeof(string));
        //    table.Columns.Add("Falta ar e dificuldade respirar", typeof(string));
        //    table.Columns.Add("Dor peito", typeof(string));
        //    table.Columns.Add("Dor torax", typeof(string));
        //    table.Columns.Add("Mal-estar generalizado", typeof(string));
        //    table.Columns.Add("Fraqueza", typeof(string));
        //    table.Columns.Add("Suor intenso", typeof(string));
        //    table.Columns.Add("Pneumonia", typeof(string));

        //    foreach (Row row in rows)
        //    {
        //        table.Rows.Add(row.nome, row.febre, row.tosseSeca, row.tosseCatarroAmarelo, row.tosseCatarroEsverdeado, row.faltaArEDificuldadeRespirar, row.dorPeito, row.dorTorax, row.malEstarGeneralizado, row.fraqueza, row.suorInteso, row.nauseaEVomito);
        //    }

        //    StringBuilder data = convertDataTableToCsvFile(table);

        //    using (StreamWriter objWriter = new StreamWriter(@"D:\Base de dados.csv"))
        //    {
        //        objWriter.WriteLine(data);
        //    }
        //}

        //public StringBuilder convertDataTableToCsvFile(System.Data.DataTable dtData)
        //{
        //    StringBuilder data = new StringBuilder();

        //    //Taking the column names.
        //    for (int column = 0; column < dtData.Columns.Count; column++)
        //    {
        //        if (column == dtData.Columns.Count - 1)//Remove os delimitadores de linha
        //            data.Append(dtData.Columns[column].ColumnName.ToString().Replace(",", ";"));
        //        else
        //            data.Append(dtData.Columns[column].ColumnName.ToString().Replace(",", ";") + ",");
        //    }

        //    data.Append(Environment.NewLine);//New line after appending columns.

        //    for (int row = 1; row < dtData.Rows.Count; row++)
        //    {
        //        for (int column = 0; column < dtData.Columns.Count; column++)
        //        {
        //            ////Making sure that end of the line, shoould not have comma delimiter.
        //            if (column == dtData.Columns.Count - 1)
        //                data.Append(dtData.Rows[row][column].ToString().Replace(", ", ";"));
        //            else
        //                data.Append(dtData.Rows[row][column].ToString().Replace(", ", ";") + ",");

        //        }

        //        //Making sure that end of the file, should not have a new line.
        //        if (row != dtData.Rows.Count - 1)
        //            data.Append(Environment.NewLine);
        //    }
        //    return data;
        //}
    }
}
