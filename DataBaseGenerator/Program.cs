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

                row.nome = "Paciente " + (numRow + 1);
                row.febre = functions.randomDouble();
                row.tosseSeca = functions.randomBoolean() ? "Sim" : "Nao";
                row.tosseCatarroAmarelo = functions.randomBoolean() ? "Sim" : "Nao";
                row.tosseCatarroEsverdeado = functions.randomBoolean() ? "Sim" : "Nao";
                row.faltaArEDificuldadeRespirar = functions.randomBoolean() ? "Sim" : "Nao";
                row.dorPeito = functions.randomBoolean() ? "Sim" : "Nao";
                row.dorTorax = functions.randomBoolean() ? "Sim" : "Nao";
                row.malEstarGeneralizado = functions.randomBoolean() ? "Sim" : "Nao";
                row.fraqueza = functions.randomBoolean() ? "Sim" : "Nao";
                row.suorInteso = functions.randomBoolean() ? "Sim" : "Nao";
                row.nauseaEVomito = functions.randomBoolean() ? "Sim" : "Nao";
                row.pneumonia = functions.randomBoolean() ? "1" : "0";

                rows.Add(row);
            }

            functions.generateExcel(rows);
        }
    }

}
