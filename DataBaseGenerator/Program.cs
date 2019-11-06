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

            List<Row> rowsBancoDados = new List<Row>();
            List<Row> rowsEntradaUsuario = new List<Row>();
            Functions functions = new Functions();

            rowsBancoDados = functions.generateRows(1000); //Gera as linhas do Excel refente ao Banco de dados  
            rowsEntradaUsuario = functions.generateRows(50); //Gera as linhas do Excel referente a Entrada de dados do usuário

            functions.generateExcel(rowsBancoDados, "Banco de dados"); //Gera o arquio Excel referente ao Banco de dados
            functions.generateExcel(rowsEntradaUsuario, "Entrada de usuários"); //Gera o arquivo Excel referente a Entrada de usuários
        }
    }

}
