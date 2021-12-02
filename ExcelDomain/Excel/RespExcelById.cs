using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDomain.Excel
{
    public class RespExcelById: BaseResponse
    {
        public DateTime DataEntrega { get; set; }
        public string NomeDoProduto { get; set; }
        public double Quantidade { get; set; }
        public double ValorUnitario { get; set; }
        public double? ValorTotal { get; set; }
    }
}
