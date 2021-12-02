using ExcelData.Repository;
using ExcelDomain;
using ExcelDomain.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;


namespace ExcelService
{
    public class ExcelCapService
    {
        private readonly IConfiguration _config;
        public ExcelCapService(IConfiguration configuration) { this._config = configuration; }

        public RespExcelInsert UploadExcel(IFormFile file)
        {
            RespExcelInsert resp = new RespExcelInsert();
            try
            {
                resp = VerificaExcelArquivo(file);
                if (resp.Resultado)
                {
                    resp = ExtrairExceltoEntity(file);
                    if (resp.Resultado)
                    {
                        resp = VerificaEInseriExcelDados(resp);
                    }
                }

            }
            catch (Exception ex)
            {
                resp.Mensagem = ex.Message;
                resp.Resultado = false;
                return resp;
            }

            return resp;
        }

        public RespExcelAll GetAllImports()
        {
            RespExcelAll resp = new();
            try
            {
                DataTable retorno = new ExcelCapRepository(_config).GetAllImports();
                List<ExcelEntity> lstretorno = DataTableToList(retorno);
                List<ExcelEntity> lstretornoValorTotal = AdicionaValorTotal(lstretorno);
                resp = PopulaRetornoAll(lstretornoValorTotal);
            }
            catch (Exception ex)
            {
                resp.Mensagem = ex.Message;
                resp.Resultado = false;
                return resp;
            }

            return resp;
        }



        public RespExcelById GetImportById(int id)
        {
            RespExcelById resp = new();
            try
            {
                DataTable retorno = new ExcelCapRepository(_config).GetImportById(id);
                if (VerificaIDEncontrado(retorno))
                {
                    List<ExcelEntity> lstretorno = DataTableToList(retorno);
                    List<ExcelEntity> lstretornoValorTotal = AdicionaValorTotal(lstretorno);
                    resp = PopulaRetornoById(lstretornoValorTotal);
                    resp.Resultado = true;
                }
                else
                {
                    resp.Resultado = false;
                    resp.Mensagem = "Id não encontrado";
                }
             }
            catch (Exception ex)
            {
                resp.Mensagem = ex.Message;
                resp.Resultado = false;
                return resp;
            }

            return resp;
        }

        #region Métodos Auxiliares

        public RespExcelInsert VerificaEInseriExcelDados(RespExcelInsert resp)
        {
            RespExcelInsert respExcel = new();
            List<string> resultado = new();
            foreach (var itemExcel in resp.ListaExcel)
            {
                if (VerificaExistenciaCampos(itemExcel))
                {
                    if (VerificaDataEntrega(itemExcel))
                    {
                        if (VerificaCampoDescricao(itemExcel))
                        {
                            if (VerificaQtdeZerada(itemExcel))
                            {
                                if (VerificaValorUnitario(itemExcel))
                                {
                                    respExcel = InsertExcel(itemExcel);
                                }
                                else
                                {
                                    respExcel.Mensagem = "Valor Unitário Zerado ou fora de Formação";
                                    respExcel.Resultado = false;
                                }
                            }
                            else
                            {
                                respExcel.Mensagem = "Qtde se encontra Zerada";
                                respExcel.Resultado = false;
                            }
                        }
                        else
                        {
                            respExcel.Mensagem = "Campo Descrição com mais de 50 caracteres";
                            respExcel.Resultado = false;
                       }
                    }
                    else
                    {
                        respExcel.Mensagem = "Data de Entrega deve ser anterior a data de hoje.";
                        respExcel.Resultado = false;
                    }
                }
                else
                {
                    respExcel.Mensagem = "Não existem campos obrigatórios";
                    respExcel.Resultado = false;
                }

                if (respExcel.Resultado)
                {
                    resultado.Add(itemExcel.NomeDoProduto + "Resultou em Sucesso");
                }
                else
                {
                    resultado.Add(itemExcel.NomeDoProduto + "Resultou em " + respExcel.Mensagem);
                }
              
            }
            resp.ListaResultado = resultado;
            return resp;
        }




        public bool VerificaExistenciaCampos(ExcelEntity itemExcel)
        {
            if (String.IsNullOrEmpty(itemExcel.NomeDoProduto) || itemExcel.DataEntrega == null || itemExcel.Quantidade == null || itemExcel.ValorUnitario == null)
            {
                return false;
            }
            return true;
        }

        public bool VerificaDataEntrega(ExcelEntity itemExcel)
        {
            if (itemExcel.DataEntrega > DateTime.Now)
            {
                return false;
            }
            return true;
        }

        public bool VerificaCampoDescricao(ExcelEntity itemExcel)
        {
            if (itemExcel.NomeDoProduto.Length > 50)
            {
                return false;
            }
            return true;
        }

        public bool VerificaQtdeZerada(ExcelEntity itemExcel)
        {
            if (itemExcel.Quantidade <= 0)
            {
                return false;
            }
            return true;
        }

        public bool VerificaValorUnitario(ExcelEntity itemExcel)
        {
            if (itemExcel.ValorUnitario <= 0)
            {
                return false;
            }
            return true;
        }


        public RespExcelInsert InsertExcel(ExcelEntity itemExcel)
        {
            RespExcelInsert insert = new();
            try
            {
                new ExcelCapRepository(_config).UploadExcel(itemExcel);
                insert.Resultado = true;
                insert.Mensagem = "Registro de Id " + itemExcel.Id.ToString() + "e Nome:" + itemExcel.NomeDoProduto + "Salvo no Banco.";

                return insert;
            }
            catch (Exception ex)
            {
                insert.Mensagem = ex.Message;
                insert.Resultado = false;
                return insert;
            }

        }

        public RespExcelInsert ExtrairExceltoEntity(IFormFile file)
        {
            RespExcelInsert respExtrair = new();
            List<ExcelEntity> lstexcelrow = new();
         
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = null;
                using (var stream = new MemoryStream())
                {
                    file.CopyToAsync(stream);
                    using (var package = new ExcelPackage(stream))
                    {
                        worksheet = package.Workbook.Worksheets.FirstOrDefault();
                        for (int i = 2; i < worksheet.Dimension.Rows + 1; i++)
                        {
                            ExcelEntity excelrow = new ExcelEntity
                            {
                                DataEntrega = (DateTime)(worksheet.Cells[i, 1]).Value,
                                NomeDoProduto = (string)(worksheet.Cells[i, 2]).Value,
                                Quantidade = (double)(worksheet.Cells[i, 3]).Value,
                                ValorUnitario = (double)(worksheet.Cells[i, 4]).Value
                            };
                            lstexcelrow.Add(excelrow);
                        }
                    }
                }
                respExtrair.Resultado = true;
                respExtrair.ListaExcel = lstexcelrow.ToList();
                return respExtrair;
            }
            catch (Exception ex)
            {
                respExtrair.Resultado = false;
                respExtrair.Mensagem = "Falha ao extrair Excel." + ex.Message;
                return respExtrair;
            }
        }

        public static RespExcelInsert VerificaExcelArquivo(IFormFile file)
        {
            RespExcelInsert respostaVerifica = new RespExcelInsert();
            if (Path.GetExtension(file.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                if (file.Length <= 0)
                {
                    respostaVerifica.Resultado = false;
                    respostaVerifica.Mensagem = "Arquivo Nulo";
                }
                else
                {
                    respostaVerifica.Resultado = true;
                }
            }
            else
            {
                respostaVerifica.Resultado = false;
                respostaVerifica.Mensagem = "Extensão de arquivo não suportado, apenas suporta xlsx";
            }
            return respostaVerifica;
        }


        public RespExcelById PopulaRetornoById(List<ExcelEntity> lstretornoValorTotal)
        {
            ExcelEntity excelEntity = lstretornoValorTotal.FirstOrDefault();
            RespExcelById respExcelById = new RespExcelById
            {
                DataEntrega = excelEntity.DataEntrega,
                NomeDoProduto = excelEntity.NomeDoProduto,
                Quantidade = excelEntity.Quantidade,
                ValorUnitario = excelEntity.ValorUnitario,
                ValorTotal = excelEntity.ValorTotal
            };
            respExcelById.Resultado = true;
            return respExcelById;
        }

        public bool VerificaIDEncontrado(DataTable retorno)
        {
            if (retorno.Rows.Count != 0)
            {
                return true;
            }
            return false;
        }

        private List<ExcelEntity> AdicionaValorTotal(List<ExcelEntity> lstretorno)
        {
            List<ExcelEntity> lstValorTotal = new();
            foreach (var itemExcel in lstretorno)
            {
                itemExcel.ValorTotal = itemExcel.Quantidade * itemExcel.ValorUnitario;
                lstValorTotal.Add(itemExcel);
            }

            return lstValorTotal;
        }

        public List<ExcelEntity> DataTableToList(DataTable retornoDt)
        {
            List<ExcelEntity> excelList = new();
            excelList = (from DataRow dr in retornoDt.Rows
                         select new ExcelEntity()
                         {
                             Id = Convert.ToInt32(dr["Quantidade"]),
                             DataEntrega = Convert.ToDateTime(dr["DataEntrega"]),
                             NomeDoProduto = dr["NomeDoProduto"].ToString(),
                             Quantidade = Convert.ToDouble(dr["Quantidade"]),
                             ValorUnitario = Convert.ToDouble(dr["Quantidade"]),

                         }).ToList();
            return excelList;
        }

        private RespExcelAll PopulaRetornoAll(List<ExcelEntity> lstretornoValorTotal)
        {
            RespExcelAll respExcelAll = new RespExcelAll
            {
                ListaExcel = lstretornoValorTotal,
                DataImportacao = DateTime.Now,
                NumeroDeItens = (int)lstretornoValorTotal.Sum(x => x.Quantidade),
                MenorDataEntrega = ObterMenorData(lstretornoValorTotal),
                ValorTotalImportacao = (double)lstretornoValorTotal.Sum(x => x.ValorTotal),
            };
            respExcelAll.Resultado = true;
            return respExcelAll;
        }

        public DateTime ObterMenorData(List<ExcelEntity> lstretornoValorTotal)
        {
            DateTime recente = new();
            foreach (var item in lstretornoValorTotal)
            {
                if (item.DataEntrega > recente)
                {
                    recente = item.DataEntrega;
                }
            }
            return recente;
        }


        #endregion

    }
}
