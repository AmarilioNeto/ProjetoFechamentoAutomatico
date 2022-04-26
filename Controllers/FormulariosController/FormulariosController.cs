using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.IO;
using OfficeOpenXml;
using ProjetoFechamentoAutomatico.Models.Formularios;
using ProjetoFechamentoAutomatico.Models.Pessoa;
using System.Data.OleDb;
using System.Data;
using System.Collections.Generic;
using System.Linq;


namespace ProjetoFechamentoAutomatico.Controllers.controlador
{
    public class FormulariosController : Controller
    {
        // GET: FormulariosController


        public ActionResult DDCVI(object caminho)
        {
            var idusuario = TempData["Usuario"];
            if (idusuario != null)
            {
                ViewData["Exibir"] = "sim";
                //retorna o caminho temporario do arquivo e salva em uma string.
                caminho = TempData["caminho"];
                if (caminho != null)
                {
                    string caminho1 = (string)caminho;
                    ConexaoOraFormularios excel = new ConexaoOraFormularios();
                    var form = excel.Excel(caminho1);
                    TempData["caminho"] = caminho;
                    TempData["Usuario"] = idusuario;
                    return View(form);
                }
                else
                {
                    return RedirectToAction("Home", "Pessoa");
                }

            }
            else
            {
                return RedirectToAction("Login", "Pessoa");
            }
        }


        public ActionResult Planilha(object caminho)
        {
            var idusuario = TempData["Usuario"];
            if (idusuario != null)
            {
                ViewData["Exibir"] = "sim";
                caminho = TempData["caminho"];
                if (caminho != null)
                {
                    string caminho1 = (string)caminho;
                    ConexaoOraFormularios excel = new ConexaoOraFormularios();
                    var form = excel.ExcelAlterado(caminho1);
                    TempData["Usuario"] = idusuario;
                    TempData["caminho"] = caminho;
                    return View(form);
                }
                else
                {
                    return RedirectToAction("Home", "Pessoa");
                }

            }
            else
            {
                return RedirectToAction("Login", "Pessoa");
            }

        }
        public ActionResult Confirmado(object caminho)
        {
           var  idusuario = TempData["Usuario"];
            if (idusuario != null)
            {
                int perfil = Convert.ToInt32(TempData["Usuario"]);
                ConexaoOraFormularios perfil2 = new ConexaoOraFormularios();
                var usuario = perfil2.PerfilPesq(perfil);
               
                if (usuario == 1)
                {
                    ViewData["Exibir"] = "sim";
                    caminho = TempData["caminho"];

                    if (caminho != null)
                    {
                        string caminho1 = (string)caminho;
                        ConexaoOraFormularios excel = new ConexaoOraFormularios();
                        var form = excel.AtualConfirmado(caminho1);
                        TempData["Usuario"] = idusuario;
                        return View(form);
                    }
                    else
                    {
                        return RedirectToAction("Home", "Pessoa");
                    }
                }
                else
                {
                    return RedirectToAction(nameof(EnviarEmail));
                }
                

            }
            else
            {
                return RedirectToAction("Login", "Pessoa");
            }
           
        }
        public ActionResult NaoConfirmado()
        {
            var idusuario = TempData["Usuario"];
            if(idusuario != null)
            {
                return RedirectToAction("Home", "Pessoa");
            }
            else
            {
                return RedirectToAction("Login", "Pessoa");
            }
          
        }
        public ActionResult  EnviarEmail()
        {
            return RedirectToAction("EnviarEmail", "Pessoa");
        }

    }
}
