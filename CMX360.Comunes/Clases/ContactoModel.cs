using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CMX360.Comunes.Clases
{
    public class ContactoModel
    {
        public string Name { get; set; }
        public string Email { get; set; }
        public string Mensaje { get; set; }
        public string Telefono { get; set; }
        public string Direccion { get; set; }
        public string CP { get; set; }
    }
}