using DataBaseGenerator.Model;
using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace DataBaseGenerator.Controller
{
    class Functions
    {
        public bool randomBoolean() //Função para gerar os valores da colunas que possuem a penas dois valores possíveis
        {
            System.Threading.Thread.Sleep(1);
            Random random = new Random();
            bool binario = random.Next(100) < 50;

            return binario;
        }

        public int randomInt() //Função para gerar os valores das colunas que possuem mais de dois valores possíveis
        {
            System.Threading.Thread.Sleep(1);
            Random random = new Random();

            return random.Next(100);
        }

        public double randomDouble() //Função para gerar os valores da coluna "Febre" 
        {
            Random random = new Random();
            var dif = Math.Abs(36.5 - 40.0);

            return Math.Round(random.NextDouble() * dif + 36.5, 2);
        }

        public List<Row> generateRows(int numRows) //Função responsável por gerar as linhas de ambos os Excel's (Banco de dados.xls e Entrada usuários.xls)
        {
            List<Row> rows = new List<Row>();
            Functions functions = new Functions();

            for (int numRow = 0; numRow < numRows; numRow++) //Loop para gerar o número de linhas de acordo com o parâmetro de entrada "numRows"
            {
                Row row = new Row();
                var intTosse = functions.randomInt();
                var intDor = functions.randomInt();
                var intFebre = functions.randomInt();

                if (intFebre <= 33) //Gerar o valor da coluna "Febre"
                    row.febre = "37,5 -";
                else if (intFebre > 33 && intFebre <= 66)
                    row.febre = "Normal";
                else
                    row.febre = "37,5 +";

                if (intTosse <= 25) //Gerar o valor da coluna "Tosse"
                    row.tosse = "Sem tosse";
                else if (intTosse > 25 && intTosse < 50)
                    row.tosse = "Tosse seca";
                else if (intTosse >= 50 && intTosse < 75)
                    row.tosse = "Tosse catarro amarelado";
                else if (intTosse >= 75)
                    row.tosse = "Tosse catarro esverdeado";

                row.faltaArEDificuldadeRespirar = functions.randomBoolean() ? "Falta ar" : "Repiracao normal";  //Gerar o valor da coluna "Falta ar e dificuldade respirar"

                if (intDor <= 25) //Gerar o valor da coluna "Dor"
                    row.dor = "Sem dor";
                else if (intDor > 25 && intDor < 50)
                    row.dor = "Torax";
                else if (intDor >= 50 && intDor < 75)
                    row.dor = "Peito";
                else if (intDor >= 75)
                    row.dor = "Torax e peito";

                row.malEstarGeneralizado = functions.randomBoolean() ? "Mal estar" : "Sem mal estar"; //Gerar o valor da coluna "Mal-estar generalizado"
                row.fraqueza = functions.randomBoolean() ? "Sim" : "Nao"; //Gerar o valor da coluna "Fraqueza"
                row.suorInteso = functions.randomBoolean() ? "Normal" : "Intenso"; //Gerar o valor da coluna "Suor intenso"
                row.nauseaEVomito = functions.randomBoolean() ? "Nausea" : "Sem nausea"; // //Gerar o valor da coluna "Nausea e Vomito"

                if ((row.febre.Equals("37,5 +") || row.febre.Equals("37,5 -")) && //Gerar o valor da coluna "Pneumonia". Caso todos os sintômas estejam positivos para Pnêumonia então o valor da coluna "Pneumonia" será 1
                    !row.tosse.Equals("Sem tosse") &&
                    row.faltaArEDificuldadeRespirar.Equals("Sim") &&
                    !row.dor.Equals("Sem dor") &&
                    row.malEstarGeneralizado.Equals("Sim") &&
                    row.fraqueza.Equals("Sim") &&
                    row.suorInteso.Equals("Sim") &&
                    row.nauseaEVomito.Equals("Sim"))
                    row.pneumonia = "1";
                else if (row.febre.Equals("Normal") && //Caso todos os sintômas estejam negativos para Pnêumonia então o valor da coluna "Pneumonia" será 0
                    row.tosse.Equals("Sem tosse") &&
                    row.faltaArEDificuldadeRespirar.Equals("Nao") &&
                    row.dor.Equals("Sem dor") &&
                    row.malEstarGeneralizado.Equals("Nao") &&
                    row.fraqueza.Equals("Nao") &&
                    row.suorInteso.Equals("Nao") &&
                    row.nauseaEVomito.Equals("Nao"))
                    row.pneumonia = "0";
                else
                    row.pneumonia = functions.randomBoolean() ? "1" : "0";

                rows.Add(row);
            }

            return rows;
        }

        public void generateExcel(List<Row> rows, string tipoExcel) //Gera os Excel's tanto do Banco de dados quanto da entrada dos usuários
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            
            //Criar o cabeçário das colunas do Excel
            xlWorkSheet.Cells[1, 1] = "Febre";
            xlWorkSheet.Cells[1, 2] = "Tosse";
            xlWorkSheet.Cells[1, 3] = "Falta ar e dificuldade respirar";
            xlWorkSheet.Cells[1, 4] = "Dor";
            xlWorkSheet.Cells[1, 5] = "Mal-estar generalizado";
            xlWorkSheet.Cells[1, 6] = "Fraqueza";
            xlWorkSheet.Cells[1, 7] = "Suor intenso";
            xlWorkSheet.Cells[1, 8] = "Nausea e Vomito";
            xlWorkSheet.Cells[1, 9] = "Pneumonia";

            int numRow = 2;

            //Preenche as linhas do Excel com as informações de cada coluna que foi gerado na função "generateRows(int numRows)"
            foreach (Row row in rows)
            {
                xlWorkSheet.Cells[numRow, 1] = row.febre;
                xlWorkSheet.Cells[numRow, 2] = row.tosse;
                xlWorkSheet.Cells[numRow, 3] = row.faltaArEDificuldadeRespirar;
                xlWorkSheet.Cells[numRow, 4] = row.dor;
                xlWorkSheet.Cells[numRow, 5] = row.malEstarGeneralizado;
                xlWorkSheet.Cells[numRow, 6] = row.fraqueza;
                xlWorkSheet.Cells[numRow, 7] = row.suorInteso;
                xlWorkSheet.Cells[numRow, 8] = row.nauseaEVomito;

                if (tipoExcel.Equals("Banco de dados"))
                    xlWorkSheet.Cells[numRow, 9] = row.pneumonia;
                else
                    xlWorkSheet.Cells[numRow, 9] = "?";

                numRow ++;
                Console.WriteLine("Número rows:" + numRow.ToString());
            }

            //Criar os arquivos Excel's no caminho C:\\Users\\Leonardo dos Santos\\Desktop\\
            if (tipoExcel.Equals("Banco de dados"))
                xlWorkBook.SaveAs("C:\\Users\\Leonardo dos Santos\\Desktop\\Base de dados.csv", Excel.XlFileFormat.xlCSV, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            else
                xlWorkBook.SaveAs("C:\\Users\\Leonardo dos Santos\\Desktop\\Entrada usuários.csv", Excel.XlFileFormat.xlCSV, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
