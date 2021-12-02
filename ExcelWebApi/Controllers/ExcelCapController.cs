using ExcelDomain.Excel;
using ExcelService;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelWebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelCapController : ControllerBase
    {
        #region Propriedades
        private readonly ExcelCapService _excelService;
        #endregion

        #region Construtor
        public ExcelCapController(ExcelCapService excelService)
        {
            this._excelService = excelService ?? throw new ArgumentException(nameof(excelService));
        }

        #endregion
       
        [HttpPost("UploadExcel")]
        [ProducesResponseType(typeof(RespExcelInsert),StatusCodes.Status200OK)]
        [ProducesResponseType(typeof(RespExcelInsert), StatusCodes.Status400BadRequest)]
        [ProducesResponseType(typeof(RespExcelInsert), StatusCodes.Status404NotFound)]
        [ProducesResponseType(typeof(RespExcelInsert), StatusCodes.Status500InternalServerError)]
        public IActionResult UploadExcel(IFormFile file)
        {
            RespExcelInsert resp = new RespExcelInsert();
            try
            {
                resp = _excelService.UploadExcel(file);
                if (resp.Resultado)
                {
                    return Ok(resp);
                }
                return StatusCode(400, resp);
            }
            catch (UnauthorizedAccessException ex)
            {
                resp.Mensagem = "Acesso não autorizado" + ex.Message;
                resp.Resultado = false;
                return StatusCode(401, resp);
            }
            catch (Exception ex)
            {
                resp.Mensagem = ex.Message;
                resp.Resultado = false;
                return StatusCode(500, resp);
            }
        }

        [HttpGet("GetAllImports")]
        [ProducesResponseType(typeof(RespExcelAll), StatusCodes.Status200OK)]
        [ProducesResponseType(typeof(RespExcelAll), StatusCodes.Status400BadRequest)]
        [ProducesResponseType(typeof(RespExcelAll), StatusCodes.Status404NotFound)]
        [ProducesResponseType(typeof(RespExcelAll), StatusCodes.Status500InternalServerError)]
        public IActionResult GetAllImports()
        {
            RespExcelAll resp = new();
            try
            {
                resp = _excelService.GetAllImports();
                if (resp.Resultado)
                {
                    return Ok(resp);
                }
                return StatusCode(400, resp);
            }
            catch (UnauthorizedAccessException ex)
            {
                resp.Mensagem = "Acesso não autorizado" + ex.Message;
                resp.Resultado = false;
                return StatusCode(401, resp);
            }
            catch (Exception ex)
            {
                resp.Mensagem = ex.Message;
                resp.Resultado = false;
                return StatusCode(500, resp);
            }
        }

        [HttpGet("GetImportById")]
        [ProducesResponseType(typeof(RespExcelById), StatusCodes.Status200OK)]
        [ProducesResponseType(typeof(RespExcelById), StatusCodes.Status400BadRequest)]
        [ProducesResponseType(typeof(RespExcelById), StatusCodes.Status404NotFound)]
        [ProducesResponseType(typeof(RespExcelById), StatusCodes.Status500InternalServerError)]
        public IActionResult GetImportById(int id)
        {
            RespExcelById resp = new();
            try
            {
                resp = _excelService.GetImportById(id);
                if (resp.Resultado)
                    return Ok(resp);

                if (resp.Mensagem == "Id não encontrado")
                    return StatusCode(404, resp);


                return StatusCode(400, resp);
            }
            catch (UnauthorizedAccessException ex)
            {
                resp.Mensagem = "Acesso não autorizado" + ex.Message;
                resp.Resultado = false;
                return StatusCode(401, resp);
            }
            catch (Exception ex)
            {
                resp.Mensagem = ex.Message;
                resp.Resultado = false;
                return StatusCode(500, resp);
            }
        }
    }
}
