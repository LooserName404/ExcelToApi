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
    public class LinhasController : ControllerBase
    {
        [HttpGet]
        public ActionResult<List<Linha>> Index()
        {
            return new LinhaData().GetAllFromSpreadsheet();
        }

        [HttpPost]
        public ActionResult Insert(LinhaInserirDto dto)
        {
            new LinhaData().InsertRowOnSpreadsheet(new Linha
            {
                HierarqProdutos = dto.HierarqProdutos,
                Denominacao = dto.Denominacao,
                Id = 0,
            });
            return NoContent();
        }

        [HttpGet("{id}")]
        public ActionResult<Linha> GetById(int id)
        {
            ActionResult<Linha> response;
            try
            {
                Linha linha = new LinhaData().GetFromSpreadsheetById(id);
                if (linha == null)
                {
                    response = NotFound();
                }
                else
                {
                    response = Ok(linha);
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
                bool deleted = new LinhaData().DeleteFromSpreadsheetById(id);
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
