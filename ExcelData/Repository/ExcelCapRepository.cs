using ExcelDomain;
using ExcelDomain.Excel;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelData.Repository
{
    public class ExcelCapRepository : DbData
    {
        public ExcelCapRepository(IConfiguration configuration) : base(configuration) { }

        public void UploadExcel(ExcelEntity excel)
        {
            string sql = " INSERT INTO [DB_EXCELCAP].[dbo].[Excel] (DataEntrega,NomeDoProduto,Quantidade,ValorUnitario)" +
                         " VALUES (@DataEntrega,@NomeDoProduto,@Quantidade,@ValorUnitario)";
            SqlTransaction transaction = null;
            try
            {
                using (var conn = new SqlConnection(_connectionString))
                {
                    conn.Open();
                    transaction = conn.BeginTransaction(IsolationLevel.ReadCommitted);


                    using (var command = new SqlCommand(sql, conn, transaction))
                    {
                        command.Parameters.AddWithValue("@DataEntrega", excel.DataEntrega);
                        command.Parameters.AddWithValue("@NomeDoProduto", excel.NomeDoProduto);
                        command.Parameters.AddWithValue("@Quantidade", excel.Quantidade);
                        command.Parameters.AddWithValue("@ValorUnitario", excel.ValorUnitario);

                        command.ExecuteNonQuery();
                        transaction.Commit();
                    }
                    conn.Close();
                }
            }
            catch (SqlException)
            {
                transaction.Rollback();
            }
        }
    
        public DataTable GetAllImports()
        {
            string select = "SELECT * FROM [DB_EXCELCAP].[dbo].[Excel]";
            var dt = new DataTable();

            try
            {
                using (var conn = new SqlConnection(_connectionString))
                {
                    conn.Open();
                    using (var command = new SqlCommand(select, conn))
                    {
                        SqlDataReader reader = command.ExecuteReader();
                        dt.Load(reader);
                        reader.Close();
                    }
                    conn.Close();
                }
                return dt;

            }
            catch (SqlException)
            {
                return null;
            }
        }

        public DataTable GetImportById(int id)
        {
            string select = "SELECT * FROM [DB_EXCELCAP].[dbo].[Excel] WHERE ID =" + id;
            var dt = new DataTable();

            try
            {
                using (var conn = new SqlConnection(_connectionString))
                {
                    conn.Open();
                    using (var command = new SqlCommand(select, conn))
                    {
                        SqlDataReader reader = command.ExecuteReader();
                        dt.Load(reader);
                        reader.Close();
                    }
                    conn.Close();
                }
                return dt;

            }
            catch (SqlException)
            {
                return null;
            }
        }
    }
}