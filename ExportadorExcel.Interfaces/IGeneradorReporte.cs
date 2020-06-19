using System.Collections.Generic;

namespace ExportadorExcel.Interfaces
{
    public interface IGeneradorReporte
    {
        byte[] GenerarReporte<T>(byte[] recurso, IEnumerable<T> datos, string nombreHojaDatos = "Datos");
        byte[] GenerarReporte(byte[] recurso, List<Fuente> fuentes);
    }
}