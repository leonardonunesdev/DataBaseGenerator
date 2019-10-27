using DataBaseGenerator.Controller;
using DataBaseGenerator.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataBaseGenerator
{
    class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {

            List<Row> rows = new List<Row>();
            Functions functions = new Functions();

            for (int numRow = 0; numRow < 1000; numRow++)
            {
                Row row = new Row();
                var intTosse = functions.randomInt();
                var intDor = functions.randomInt();
                var intFebre = functions.randomInt();

                row.nome = "Paciente " + (numRow + 1);

                if (intFebre <= 33)
                    row.febre = "37,5 -";
                else if (intFebre > 33 && intFebre <= 66)
                    row.febre = "Normal";
                else
                    row.febre = "37,5 +";

                if (intTosse <= 25)
                    row.tosse = "Sem tosse";
                else if (intTosse > 25 && intTosse < 50)
                    row.tosse = "Tosse seca";
                else if(intTosse >= 50 && intTosse < 75)
                    row.tosse = "Tosse catarro amarelado";
                else if(intTosse >= 75)
                    row.tosse = "Tosse catarro esverdeado";

                row.faltaArEDificuldadeRespirar = functions.randomBoolean() ? "Falta ar" : "Repiracao normal";

                if (intDor <= 25)
                    row.dor = "Sem dor";
                else if (intDor > 25 && intDor < 50)
                    row.dor = "Torax";
                else if (intDor >= 50 && intDor < 75)
                    row.dor = "Peito";
                else if (intDor >= 75)
                    row.dor = "Torax e peito";

                row.malEstarGeneralizado = functions.randomBoolean() ? "Mal estar" : "Sem mal estar";
                row.fraqueza = functions.randomBoolean() ? "Sim" : "Nao";
                row.suorInteso = functions.randomBoolean() ? "Normal" : "Intenso";
                row.nauseaEVomito = functions.randomBoolean() ? "Nausea" : "Sem nausea";

                if ((row.febre.Equals("37,5 +") || row.febre.Equals("37,5 -")) &&
                    !row.tosse.Equals("Sem tosse") &&
                    row.faltaArEDificuldadeRespirar.Equals("Sim") &&
                    !row.dor.Equals("Sem dor") &&
                    row.malEstarGeneralizado.Equals("Sim") &&
                    row.fraqueza.Equals("Sim") &&
                    row.suorInteso.Equals("Sim") &&
                    row.nauseaEVomito.Equals("Sim"))
                    row.pneumonia = "1";
                else if(row.febre.Equals("Normal") &&
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

            functions.generateExcel(rows);
        }
    }

}
