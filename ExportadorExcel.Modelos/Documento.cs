using System;

namespace ExportadorExcel.Modelos
{
    public class Documento
    {
        public Documento(int id, string docTipo, int docNo, DateTime docFecha, string docDesc, 
            string docObservaciones, string movDesc, string movCc, string centroCostosDesc,
            string movCuentaC, string cuentaDesc, string movNit, string terceroDesc, 
            decimal debito, decimal credito)
        {
            Id = id;
            DocTipo = docTipo;
            DocNo = docNo;
            DocFecha = docFecha;
            DocDesc = docDesc;
            DocObservaciones = docObservaciones;
            MovDesc = movDesc;
            MovCC = movCc;
            CentroCostosDesc = centroCostosDesc;
            MovCuentaC = movCuentaC;
            CuentaDesc = cuentaDesc;
            MovNit = movNit;
            TerceroDesc = terceroDesc;
            Debito = debito;
            Credito = credito;
        }

        public int Id { get; }
        public string DocTipo { get; }
        public int DocNo { get; }
        public DateTime DocFecha { get; }
        public string DocDesc { get; }
        public string DocObservaciones { get; }
        public string MovDesc { get; }
        public string MovCC { get; }
        public string CentroCostosDesc { get; }
        public string MovCuentaC { get; }
        public string CuentaDesc { get; }
        public string MovNit { get; }
        public string TerceroDesc { get; }
        public decimal Debito { get; }
        public decimal Credito { get; }
    }
}
