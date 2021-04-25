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
    public class ProdutosController : ControllerBase
    {
        [HttpGet]
        public List<Produto> Index()
        {
            return new ProdutoData().GetAllFromSpreadsheet();
        }

        [HttpPost]
        public ActionResult Insert(ProdutoInserirDto dto)
        {
            new ProdutoData().InsertRowOnSpreadsheet(new Produto
            {
                Codmaterial = dto.Codmaterial,
                Codigofamilia = dto.Codigofamilia,
                CodigoNCM = dto.CodigoNCM,
                Id = 0,
            });
            return NoContent();
        }

        [HttpGet("{id}")]
        public ActionResult<Produto> GetById(int id)
        {
            ActionResult<Produto> response;
            try
            {
                Produto produto = new ProdutoData().GetFromSpreadsheetById(id);
                if (produto == null)
                {
                    response = NotFound();
                }
                else
                {
                    response = Ok(produto);
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
                bool deleted = new ProdutoData().DeleteFromSpreadsheetById(id);
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
