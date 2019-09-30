using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataBaseGenerator.Model
{
    class Row
    {
        public string nome { get; set; }
        public double febre { get; set; }
        public string tosseSeca{ get; set; }
        public string tosseCatarroAmarelo{ get; set; }
        public string tosseCatarroEsverdeado{ get; set; }
        public string faltaArEDificuldadeRespirar{ get; set; }
        public string dorPeito{ get; set; }
        public string dorTorax{ get; set; }
        public string malEstarGeneralizado{ get; set; }
        public string fraqueza{ get; set; }
        public string suorInteso{ get; set; }
        public string nauseaEVomito{ get; set; }
        public string pneumonia { get; set; }
    }
}
