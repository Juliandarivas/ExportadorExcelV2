using System;
using System.Collections.Generic;
using ExportadorExcel.API.Properties;
using ExportadorExcel.Interfaces;
using ExportadorExcel.Modelos;
using Microsoft.AspNetCore.Mvc;

namespace ExportadorExcel.API.Controllers
{
    [ApiController]
    [Route("[controller]")]

    public class GeneradorDocumentosContablesController : ControllerBase
    {
        private readonly IGeneradorReporte _generadorReporte;

        public GeneradorDocumentosContablesController(IGeneradorReporte generadorReporte)
        {
            _generadorReporte = generadorReporte;
        }


        [HttpGet]
        public ActionResult Generar()
        {
            List<Documento> documentos = CrearDocumentos();

            byte[] bytes = _generadorReporte.GenerarReporte(Resources.PlantillaDocumentoContable, documentos);

            return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "docuentoContable.xlsx");
        }

        public List<Documento> CrearDocumentos()
        {
            return new List<Documento>
            {
                new Documento(1, "CP", 12345, new DateTime(2020, 5, 1), "Documento prueba 1", "Movimiento prueba",
                    "010101", "Administrativo", "1005", "1005", "CuentaContable", "165435413", "Sincosoft", 5_000_000,
                    5_000_000),
                new Documento(2, "CP", 12345, new DateTime(2020, 5, 1), "Documento prueba 2", "Movimiento prueba",
                    "010101", "Administrativo", "1005", "1005", "CuentaContable", "165435413", "Sincosoft", 5_000_000,
                    5_000_000),
                new Documento(3, "CP", 12345, new DateTime(2020, 5, 1), "Documento prueba 3", "Movimiento prueba",
                    "010101", "Administrativo", "1005", "1005", "CuentaContable", "165435413", "Sincosoft", 5_000_000,
                    5_000_000),
                new Documento(4, "CP", 12345, new DateTime(2020, 5, 1), "Documento prueba 4", "Movimiento prueba",
                    "010101", "Administrativo", "1005", "1005", "CuentaContable", "165435413", "Sincosoft", 5_000_000,
                    5_000_000),
                new Documento(1, "CP", 12345, new DateTime(2020, 5, 1), "Documento prueba", "Movimiento prueba",
                    "010101", "Administrativo", "1005", "1005", "CuentaContable", "165435413", "Sincosoft", 5_000_000,
                    5_000_000),
                new Documento(1, "CP", 12345, new DateTime(2020, 5, 1), "Documento prueba", "Movimiento prueba",
                    "010101", "Administrativo", "1005", "1005", "CuentaContable", "165435413", "Sincosoft", 5_000_000,
                    5_000_000),
                new Documento(1, "CP", 12345, new DateTime(2020, 5, 1), "Documento prueba", "Movimiento prueba",
                    "010101", "Administrativo", "1005", "1005", "CuentaContable", "165435413", "Sincosoft", 5_000_000,
                    5_000_000),
                new Documento(1, "CP", 12345, new DateTime(2020, 5, 1), "Documento prueba", "Movimiento prueba",
                    "010101", "Administrativo", "1005", "1005", "CuentaContable", "165435413", "Sincosoft", 5_000_000,
                    5_000_000),
                new Documento(1, "CP", 12345, new DateTime(2020, 5, 1), "Documento prueba", "Movimiento prueba",
                    "010101", "Administrativo", "1005", "1005", "CuentaContable", "165435413", "Sincosoft", 5_000_000,
                    5_000_000),
                new Documento(1, "CP", 12345, new DateTime(2020, 5, 1), "Documento prueba", "Movimiento prueba",
                    "010101", "Administrativo", "1005", "1005", "CuentaContable", "165435413", "Sincosoft", 5_000_000,
                    5_000_000),
                new Documento(1, "CP", 12345, new DateTime(2020, 5, 1), "Documento prueba", "Movimiento prueba",
                    "010101", "Administrativo", "1005", "1005", "CuentaContable", "165435413", "Sincosoft", 5_000_000,
                    5_000_000),
                new Documento(1, "CP", 12345, new DateTime(2020, 5, 1), "Documento prueba", "Movimiento prueba",
                    "010101", "Administrativo", "1005", "1005", "CuentaContable", "165435413", "Sincosoft", 5_000_000,
                    5_000_000),
                new Documento(1, "CP", 12345, new DateTime(2020, 5, 1), "Documento prueba", "Movimiento prueba",
                    "010101", "Administrativo", "1005", "1005", "CuentaContable", "165435413", "Sincosoft", 5_000_000,
                    5_000_000),
                new Documento(1, "CP", 12345, new DateTime(2020, 5, 1), "Documento prueba", "Movimiento prueba",
                    "010101", "Administrativo", "1005", "1005", "CuentaContable", "165435413", "Sincosoft", 5_000_000,
                    5_000_000),
                new Documento(1, "CP", 12345, new DateTime(2020, 5, 1), "Documento prueba", "Movimiento prueba",
                    "010101", "Administrativo", "1005", "1005", "CuentaContable", "165435413", "Sincosoft", 5_000_000,
                    5_000_000),
                new Documento(1, "CP", 12345, new DateTime(2020, 5, 1), "Documento prueba", "Movimiento prueba",
                    "010101", "Administrativo", "1005", "1005", "CuentaContable", "165435413", "Sincosoft", 5_000_000,
                    5_000_000),
                new Documento(1, "CP", 12345, new DateTime(2020, 5, 1), "Documento prueba", "Movimiento prueba",
                    "010101", "Administrativo", "1005", "1005", "CuentaContable", "165435413", "Sincosoft", 5_000_000,
                    5_000_000),
                new Documento(1, "CP", 12345, new DateTime(2020, 5, 1), "Documento prueba", "Movimiento prueba",
                    "010101", "Administrativo", "1005", "1005", "CuentaContable", "165435413", "Sincosoft", 5_000_000,
                    5_000_000),
            };
        }
    }
}
