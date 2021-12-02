using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDomain.Excel
{
    public class RespExcelInsert: BaseResponse
    {
        public List<ExcelEntity> ListaExcel { get; set; }

        public List<string> ListaResultado { get; set; }
    }
}
