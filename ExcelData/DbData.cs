using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelData
{
    public abstract class DbData
    {
        public readonly string _connectionString;

        protected DbData(IConfiguration configuration) 
        {
            _connectionString = configuration.GetConnectionString("defaultConnection");
        }
    }
}
