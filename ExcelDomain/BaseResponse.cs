using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDomain
{
    public class BaseResponse
    {
        public bool Resultado { get; set; }
        public string Mensagem { get; set; }
        public static DateTime HoraEnvio {
            get 
            {
                return DateTime.UtcNow;            
            }
        }
    }
}
