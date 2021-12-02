using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDomain.Excel
{
    public class RespExcelAll:BaseResponse
    {
        public List<ExcelEntity> ListaExcel { get; set; }
        public DateTime DataImportacao { get; set; }
        public int NumeroDeItens { get; set; }
        public DateTime MenorDataEntrega { get; set; }
        public double ValorTotalImportacao { get; set; }
    }
}
