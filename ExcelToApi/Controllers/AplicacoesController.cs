using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using ExcelToApi.Data;
using ExcelToApi.DTOs;
using ExcelToApi.Entities;

namespace ExcelToApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class AplicacoesController : ControllerBase
    {
        [HttpGet]
        public List<Aplicacao> Index()
        {
            return new AplicacaoData().GetAllFromSpreadsheet();
        }

        [HttpPost]
        public ActionResult Insert(AplicacaoInserirDto dto)
        {
            new AplicacaoData().InsertRowOnSpreadsheet(new Aplicacao()
            {
                Codmaterial = dto.Codmaterial,
                NomeMarca = dto.NomeMarca,
                NomeModelo = dto.NomeModelo,
                Cilindradas = dto.Cilindradas,
                Id = 0,
            });
            return NoContent();
        }

        [HttpGet("{id}")]
        public ActionResult<Aplicacao> GetById(int id)
        {
            ActionResult<Aplicacao> response;
            try
            {
                Aplicacao aplicacao = new AplicacaoData().GetFromSpreadsheetById(id);
                if (aplicacao == null)
                {
                    response = NotFound();
                }
                else
                {
                    response = Ok(aplicacao);
                }
            }
            catch (ArgumentOutOfRangeException e)
            {
                response = BadRequest(new { Message = e.Message });
            }

            return response;
        }

        [HttpDelete("{id}")]
        public ActionResult DeleteById(int id)
        {
            try
            {
                bool deleted = new AplicacaoData().DeleteFromSpreadsheetById(id);
                if (deleted)
                {
                    return NoContent();
                }
                return NotFound();
            }
            catch (ArgumentOutOfRangeException e)
            {
                return BadRequest(new { Message = e.Message });
            }
        }
    }
}
