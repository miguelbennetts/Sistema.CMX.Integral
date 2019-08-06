using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CMX360.Comunes.Clases
{
  public  class DatosSuma
    {
        public DatosSuma()
        {
            Filas = new List<int>();
        }

        public DatosSuma(int columna)
        {
            Filas = new List<int>();
            Columna = columna;
        }
        public int Columna { get; set; }
        public List<int> Filas { get; set; }
    }
}
