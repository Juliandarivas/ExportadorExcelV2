using System;
using System.Collections.Generic;
using System.Linq;

namespace ExportadorExcel.Interfaces
{
    public class Fuente
    {
        public static Fuente Crear<T>(string nombreHoja, IEnumerable<T> datos)
        {
            return new Fuente(nombreHoja, typeof(T), datos.Cast<object>());
        }

        private Fuente(string nombreHoja, Type tipo, IEnumerable<object> datos)
        {
            NombreHoja = nombreHoja;
            Tipo = tipo;
            Datos = datos;
        }

        public string NombreHoja { get; set; }
        public Type Tipo { get; set; }
        public IEnumerable<object> Datos { get; set; }
    }
}
