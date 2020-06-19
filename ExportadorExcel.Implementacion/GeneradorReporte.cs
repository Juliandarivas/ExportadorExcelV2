using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using ExportadorExcel.Interfaces;
using OfficeOpenXml;

namespace ExportadorExcel.Implementacion
{
    public class GeneradorReporte : IGeneradorReporte
    {
        public byte[] GenerarReporte<T>(byte[] recurso, IEnumerable<T> datos, string nombreHojaDatos = "Datos")
        {
            Stream stream = new MemoryStream(recurso);
            return GenerarExcel(stream, datos, nombreHojaDatos);
        }

        public byte[] GenerarReporte(byte[] recurso, List<Fuente> fuentes)
        {
            using Stream excelPlantilla = new MemoryStream(recurso);

            ExcelPackage excel = new ExcelPackage(excelPlantilla);

            foreach (var fuente in fuentes)
            {
                ExcelWorksheet hojaExcel = excel.Workbook.Worksheets[fuente.NombreHoja];
                CargarDatos(hojaExcel, fuente);
                ConvertirPropiedadesEnColumnas(hojaExcel, fuente.Tipo);
            }

            return excel.GetAsByteArray();
        }

        private void CargarDatos(ExcelWorksheet hojaExcel, Fuente fuente)
        {
            hojaExcel.InsertRow(2, fuente.Datos.Count() -1);
            ExcelRange rangoExcel = hojaExcel.Cells["A2"];

            MethodInfo informacionMetodo = typeof(ExcelRange).GetMethods()
                .First(x => x.Name == nameof(rangoExcel.LoadFromCollection));

            MethodInfo metodoGenerico = informacionMetodo.MakeGenericMethod(fuente.Tipo);

            metodoGenerico.Invoke(rangoExcel, new[] {fuente.Datos});
        }

        private static byte[] GenerarExcel<T>(Stream excelPlantilla, IEnumerable<T> datos, string nombreHojaDatos)
        {
            ExcelPackage excel = new ExcelPackage(excelPlantilla);
            ExcelWorksheet sheet = excel.Workbook.Worksheets[nombreHojaDatos];

            CargarDatos(datos, sheet);
            ConvertirPropiedadesEnColumnas<T>(sheet);

            return excel.GetAsByteArray();
        }

        private static void CargarDatos<T>(IEnumerable<T> datos, ExcelWorksheet sheet)
        {
            sheet.InsertRow(2, datos.Count() - 1);
            sheet.Cells["A2"].LoadFromCollection(datos);
        }

        private static void ConvertirPropiedadesEnColumnas(ExcelWorksheet sheet, Type tipo)
        {
            var indiceColumna = 1;
            foreach (var propiedad in tipo.GetProperties())
            {
                if (propiedad.PropertyType == typeof(DateTime))
                    sheet.Column(indiceColumna).Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

                sheet.SetValue(1, indiceColumna++, propiedad.Name);
            }
        }

        private static void ConvertirPropiedadesEnColumnas<T>(ExcelWorksheet sheet)
        {
            var indiceColumna = 1;
            foreach (var propiedad in typeof(T).GetProperties())
            {
                if (propiedad.PropertyType == typeof(DateTime))
                    sheet.Column(indiceColumna).Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

                sheet.SetValue(1, indiceColumna++, propiedad.Name);
            }
        }
    }

}
