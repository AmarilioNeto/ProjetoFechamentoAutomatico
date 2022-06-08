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
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Builder;
using System.Net;

namespace ProjetoFechamentoAutomatico.Controllers.controlador
{
    public class FormulariosController : Controller
    {
        // GET: FormulariosController
        static object documento;
        static object documento1;
        public ActionResult DDCVI(object caminho)
        {

            var idusuario = TempData.Peek("Pessoa");
            var USUARIO = TempData["USUARIO"];
            ViewData["exibir login"] = "sim";
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
                    if (form != null)
                    {
                        TempData["caminho"] = caminho;
                        TempData["Pessoa"] = idusuario;
                        TempData["USUARIO"] = USUARIO;
                        return View(form);
                    }
                    else
                    {
                        TempData["erro form"] = "Erro ao carregar a planilha Excel. Escolha um arquivo excel valido.";
                        return View(form);
                    }

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

        public ActionResult TipoMovimento(object caminho)
        {

            var idusuario = TempData.Peek("Pessoa");
            // var USUARIO = TempData["USUARIO"];
            ViewData["exibir login"] = "sim";
            if (idusuario != null)
            {
                ViewData["Exibir"] = "sim";
                //retorna o caminho temporario do arquivo e salva em uma string.
                caminho = TempData["caminho"];
                if (caminho != null)
                {
                    string caminho1 = (string)caminho;
                    ConexaoOraFormularios excel = new ConexaoOraFormularios();
                    var form = excel.ExcelTipoMovimento(caminho1);
                    if (form != null)
                    {
                        TempData["caminho"] = caminho;
                        TempData["Pessoa"] = idusuario;
                        //TempData["USUARIO"] = USUARIO;
                        return View(form);
                    }
                    else
                    {
                        TempData["erro form"] = "Erro ao carregar a planilha Excel. Escolha um arquivo excel valido.";
                        return View(form);
                    }

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
        public ActionResult TipoMovimentoContabil(object caminho)
        {

            var id = Request.Path;
            var idusuario = id.Value.Split("/Formularios/TipoMovimentoContabil/")[1];
            //var idusuario = TempData.Peek("Pessoa");
            //var USUARIO = TempData["USUARIO"];
            //TempData.Remove("movimentos");
            ViewData["exibir login"] = "sim";
            if (idusuario != null)
            {

                ViewData["Exibir"] = "sim";
                //retorna o caminho tratado temporario do arquivo e salva em uma string.
                var cami = Request.QueryString;

                caminho = cami.Value.Replace("?caminho=C%3A%5C", " C:\\").Replace("%5C", "\\").Trim();
                if (caminho != null)
                {
                    string caminho1 = (string)caminho;
                    int Linhas = 0;
                    ConexaoOraFormularios excel = new ConexaoOraFormularios();
                    var form = excel.ExcelTipoMovimentoContabil(caminho1, Linhas);
                    if (form != null)
                    {
                        List<string> reduzido = new List<string>();
                        for (var i = 0; i < form[0].Linhas; i++)
                        {
                            string red = form[i].Reduzido;
                            reduzido.Add(red);
                        }
                        List<string> depositos = new List<string>();
                        for (var i = 0; i < form[0].Linhas; i++)
                        {
                            string dep = form[i].Deposito;
                            depositos.Add(dep);
                        }
                        List<string> Tipomovimento = new List<string>();
                        for (var i = 0; i < form[0].Linhas; i++)
                        {
                            string tipmov = form[i].TipoMov;
                            Tipomovimento.Add(tipmov);
                        }
                        List<string> Datas = new List<string>();
                        for (var i = 0; i < form[0].Linhas; i++)
                        {
                            string data = form[i].Data;
                            Datas.Add(data);
                        }
                        List<string> NumDocto = new List<string>();
                        for (var i = 0; i < form[0].Linhas; i++)
                        {
                            string numdocto = form[i].NumDocto;
                            NumDocto.Add(numdocto);
                        }
                        List<string> ContaEstoque = new List<string>();
                        for (var i = 0; i < form[0].Linhas; i++)
                        {
                            string contaestoque = form[i].ContaEstoque;
                            ContaEstoque.Add(contaestoque);
                        }
                        List<string> ContaContrapartida = new List<string>();
                        for (var i = 0; i < form[0].Linhas; i++)
                        {
                            string contacontrapartida = form[i].ContaContrapartida;
                            ContaContrapartida.Add(contacontrapartida);

                        }
                        //string serial = JsonSerializer.Serialize(form);

                        HttpWebRequest req = WebRequest.CreateHttp("https://localhost:44319/Formularios/TipoMovimentoContabil/");
                        req.Method = "GET";
                        //req.ContentType = "application/json";
                        req.Headers.Add("caminho1", caminho1);
                        req.Headers.Add("idusuario", idusuario);
                        documento = req;

                        return View(form);
                    }
                    else
                    {
                        TempData["erro form"] = "Erro ao carregar a planilha Excel. Escolha um arquivo excel valido.";
                        return View(form);
                    }

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
            var idusuario = TempData["Pessoa"];
            ViewData["exibir login"] = "sim";
            if (idusuario != null)
            {
                ViewData["Exibir"] = "sim";
                caminho = TempData["caminho"];
                if (caminho != null)
                {
                    string caminho1 = (string)caminho;
                    ConexaoOraFormularios excel = new ConexaoOraFormularios();
                    var form = excel.ExcelAlterado(caminho1);
                    TempData["Pessoa"] = idusuario;
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
            var idusuario = TempData["Pessoa"];
            ViewData["exibir login"] = "sim";
            if (idusuario != null)
            {

                int perfil = Convert.ToInt32(TempData["Pessoa"]);
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
                        TempData["Pessoa"] = idusuario;
                        return View(form);
                    }
                    else
                    {
                        return RedirectToAction("Home", "Pessoa");
                    }
                }
                else
                {
                    //desativado por enquanto
                    return RedirectToAction(nameof(EnviarEmail));
                }


            }
            else
            {
                return RedirectToAction("Login", "Pessoa");
            }

        }
        public ActionResult PlanilhaTipoMovimento()
        {
            var redirect = TempData["redirect"];
            var idusuario = TempData.Peek("Pessoa");
            var USUARIO = TempData["USUARIO"];
            ViewData["Exibir"] = "sim";
            ViewData["exibir login"] = "sim";
            if (redirect != null)
            {
                TempData["errotipomovimento"] = "Tipo de Movimento invalido. Escolha um tipo de movimento valido";
            }

            return View();
        }
        public ActionResult PlanilhaTipoMovimentoContabil()
        {
            var idusuario = TempData.Peek("Pessoa");
            var USUARIO = TempData["USUARIO"];
            ViewData["Exibir"] = "sim";
            ViewData["exibir login"] = "sim";
            return View();
        }
        public ActionResult PlanilhaTipoMovimentoConfirmar(TipoMovimento tipo)
        {
            var idusuario = TempData.Peek("Pessoa");
            var USUARIO = TempData["USUARIO"];
            ViewData["DE"] = tipo.DE;
            ViewData["PARA"] = tipo.PARA;
            ViewData["Exibir"] = "sim";
            ViewData["exibir login"] = "sim";
            //TempData["de"] = tipo.DE;
            //TempData["para"] = tipo.PARA;
            return View();
        }
        public ActionResult PlanilhaContabilPasso1(string caminho1, int Linhas)
        {


            ViewData["Exibir"] = "sim";
            ViewData["exibir login"] = "sim";
            // TempData["Linhas"] = linhas;
            //TempData["deposito"] = deposito;

            if (TempData["data"] != null)
            {
                ViewData["Errodata"] = "A data final deve ser maior o igual a inicial";
            }
            if (TempData["campo"] != null)
            {
                ViewData["ErrocampoObrigatorio"] = "Os campos obrigatorios devem ser preenchidos";
            }
            if (TempData["planilha"] != null)
            {
                ViewData["Erroplanilha"] = "Quantidade de registro diferente da planilha.";
            }

            return View();
        }
        [HttpPost]
        public ActionResult PlanilhaContabilPasso1(ContabilPasso1 contabil, string caminho1, int Linhas)
        {

            caminho1 = ((System.Net.HttpWebRequest)documento).Headers.ToString().Split("idusuario: ")[0].Split("caminho1: ")[1].Trim();
            var idusuario = ((System.Net.HttpWebRequest)documento).Headers.ToString().Split("idusuario: ")[1].Trim();
            ConexaoOraFormularios excel = new ConexaoOraFormularios();
            var form = excel.ExcelTipoMovimentoContabil(caminho1, Linhas);

            if (form != null)
            {
                List<string> reduzido = new List<string>();
                for (var i = 0; i < form[0].Linhas; i++)
                {
                    string red = form[i].Reduzido;
                    reduzido.Add(red);
                }
                List<string> depositos = new List<string>();
                for (var i = 0; i < form[0].Linhas; i++)
                {
                    string dep = form[i].Deposito;
                    depositos.Add(dep);
                }
                List<string> Tipomovimento = new List<string>();
                for (var i = 0; i < form[0].Linhas; i++)
                {
                    string tipmov = form[i].TipoMov;
                    Tipomovimento.Add(tipmov);
                }
                List<string> Datas = new List<string>();
                for (var i = 0; i < form[0].Linhas; i++)
                {
                    string data = form[i].Data;
                    Datas.Add(data);
                }
                List<string> NumDocto = new List<string>();
                for (var i = 0; i < form[0].Linhas; i++)
                {
                    string numdocto = form[i].NumDocto;
                    NumDocto.Add(numdocto);
                }
                List<string> ContaEstoque = new List<string>();
                for (var i = 0; i < form[0].Linhas; i++)
                {
                    string contaestoque = form[i].ContaEstoque;
                    ContaEstoque.Add(contaestoque);
                }
                List<string> ContaContrapartida = new List<string>();
                for (var i = 0; i < form[0].Linhas; i++)
                {
                    string contacontrapartida = form[i].ContaContrapartida;
                    ContaContrapartida.Add(contacontrapartida);

                }
            }

            var linhas = form[0].Linhas;
            var deposito = form[0].Deposito;


            ViewData["Exibir"] = "sim";
            ViewData["exibir login"] = "sim";


            if (contabil.NumTipoMovimento != null && contabil.Deposito != null)
            {
                if (contabil.DataInicial >= contabil.DataFinal)
                {
                    TempData["data"] = "sim";
                    return RedirectToAction(nameof(PlanilhaContabilPasso1));
                }
                else
                {
                    ConexaoOraFormularios conexao = new ConexaoOraFormularios();
                    var passo1 = conexao.Contabils(contabil);
                    var linhas1 = (int)linhas;
                    string Linha = Convert.ToString(form[0].Linhas);

                    if (passo1.Count == linhas1 && deposito.ToString() == contabil.Deposito)
                    {
                        List<int> movimentos = new List<int>();
                        for (int elements = 0; elements < passo1.Count; elements++)
                        {
                            var numeroMovimento = Convert.ToInt32(form[elements].NumDocto);
                            movimentos.Add(numeroMovimento);
                        }
                        string mov = JsonSerializer.Serialize(movimentos);
                        HttpWebRequest req = WebRequest.CreateHttp("https://localhost:44319/Formularios/PlanilhaContabilPasso2/");
                        req.Method = "GET";
                        req.Headers.Add("mov", mov);
                        req.Headers.Add("idusuario", idusuario);
                        req.Headers.Add("caminho1", caminho1);
                        req.Headers.Add("Linha", Linha);
                        documento = req;

                        return RedirectToAction(nameof(PlanilhaContabilPasso2));
                    }
                    else
                    {
                        TempData["planilha"] = "sim";
                        return RedirectToAction(nameof(PlanilhaContabilPasso1));
                    }

                }

            }
            else
            {
                TempData["campo"] = "sim";
                return RedirectToAction(nameof(PlanilhaContabilPasso1));
            }

        }
        public ActionResult PlanilhaContabilPasso2()
        {

            var idusuario = ((System.Net.HttpWebRequest)documento).Headers.ToString().Split("caminho1: ")[0].Split("idusuario: ")[1].Trim();
            //var USUARIO = TempData["USUARIO"];
            var caminhi1 = ((System.Net.HttpWebRequest)documento).Headers.ToString().Split("caminho1: ")[1].Trim();
            var movimentos = ((System.Net.HttpWebRequest)documento).Headers.ToString().Split("caminho1: ")[0].Split("idusuario: ")[0].Split("mov: ")[1].Trim();


            ViewData["Exibir"] = "sim";


            if (TempData["PesquisaAvancada"] != null)
            {
                ViewData["pesquisaavancada"] = "sim";
            }
            TempData.Remove("PesquisaAvancada");

            if (TempData["dataerrada"] != null)
            {
                ViewData["Errodata"] = "A data final deve ser maior o igual a inicial";
            }
            if (TempData["obrigatorio"] != null)
            {
                ViewData["Erroobrigatorio"] = "Os campos obrigatorios devem ser preenchidos";
            }
            return View();
        }
        [HttpPost]
        public ActionResult TabelaContabilPasso2(Tblmovimentacao tblmovimentacao, string caminho1, int Linhas)
        {
            var idusuario = ((System.Net.HttpWebRequest)documento).Headers.ToString().Split("caminho1: ")[0].Split("idusuario: ")[1].Trim();
            caminho1 = ((System.Net.HttpWebRequest)documento).Headers.ToString().Split("caminho1: ")[1].Split("Linha: ")[0].Trim();
            var mov = ((System.Net.HttpWebRequest)documento).Headers.ToString().Split("caminho1: ")[0].Split("idusuario: ")[0].Split("mov: ")[1].Trim();
            var Linha = ((System.Net.HttpWebRequest)documento).Headers.ToString().Split("caminho1: ")[1].Split("Linha: ")[1].Trim();
            Linhas = Convert.ToInt32(Linha);
            ConexaoOraFormularios excel = new ConexaoOraFormularios();
            var form = excel.ExcelTipoMovimentoContabil(caminho1, Linhas);
            var movimentos = mov.ToString().Replace("[", " ").Replace("]", " ").Trim();
            if (form != null)
            {
                List<string> reduzido = new List<string>();
                for (var i = 0; i < form[0].Linhas; i++)
                {
                    string red = form[i].Reduzido;
                    reduzido.Add(red);
                }
                List<string> depositos = new List<string>();
                for (var i = 0; i < form[0].Linhas; i++)
                {
                    string dep = form[i].Deposito;
                    depositos.Add(dep);
                }
                List<string> Tipomovimento = new List<string>();
                for (var i = 0; i < form[0].Linhas; i++)
                {
                    string tipmov = form[i].TipoMov;
                    Tipomovimento.Add(tipmov);
                }
                List<string> Datas = new List<string>();
                for (var i = 0; i < form[0].Linhas; i++)
                {
                    string data = form[i].Data;
                    Datas.Add(data);
                }
                List<string> NumDocto = new List<string>();
                for (var i = 0; i < form[0].Linhas; i++)
                {
                    string numdocto = form[i].NumDocto;
                    NumDocto.Add(numdocto);
                }
                List<string> ContaEstoque = new List<string>();
                for (var i = 0; i < form[0].Linhas; i++)
                {
                    string contaestoque = form[i].ContaEstoque;
                    ContaEstoque.Add(contaestoque);
                }
                List<string> ContaContrapartida = new List<string>();
                for (var i = 0; i < form[0].Linhas; i++)
                {
                    string contacontrapartida = form[i].ContaContrapartida;
                    ContaContrapartida.Add(contacontrapartida);

                }
            }

            ViewData["Exibir"] = "sim";
            ViewData["exibir login"] = "sim";

            if (tblmovimentacao.DATA_INICIAL.Year == 0001 || tblmovimentacao.DATA_FINAL.Year == 0001 || tblmovimentacao.DEPOSITO == 0)
            {
                TempData["obrigatorio"] = "sim";
                return RedirectToAction(nameof(PlanilhaContabilPasso2));
            }
            if (tblmovimentacao.DATA_FINAL >= tblmovimentacao.DATA_INICIAL)
            {

                ConexaoOraFormularios conexao = new ConexaoOraFormularios();
                var passo = conexao.Contabils2(tblmovimentacao, movimentos);

                List<string> reduzido1 = new List<string>();
                for (var i = 0; i < passo.Count; i++)
                {
                    string red = Convert.ToString(passo[i].REDUZIDO_ITEM);
                    reduzido1.Add(red);
                }
                List<string> deposito1 = new List<string>();
                for (var i = 0; i < passo.Count; i++)
                {
                    string dep = Convert.ToString(passo[i].DEPOSITO);
                    deposito1.Add(dep);
                }
                List<string> Tipomovimento1 = new List<string>();
                for (var i = 0; i < passo.Count; i++)
                {
                    string tipmov = Convert.ToString(passo[i].NUM_TIPO_MOVIMENTO);
                    Tipomovimento1.Add(tipmov);
                }
                List<string> Datas1 = new List<string>();
                for (var i = 0; i < passo.Count; i++)
                {
                    string datas = Convert.ToString(passo[i].DATA_DOCUMENTO);
                    Datas1.Add(datas);
                }
                List<string> NumDoct1 = new List<string>();
                for (var i = 0; i < passo.Count; i++)
                {
                    string numdoct = Convert.ToString(passo[i].NUM_DOCUMENTO);
                    NumDoct1.Add(numdoct);
                }
                List<string> ContaEstoque1 = new List<string>();
                for (var i = 0; i < passo.Count; i++)
                {
                    string contestoque = Convert.ToString(passo[i].CONTA_CONTABIL);
                    ContaEstoque1.Add(contestoque);
                }
                List<string> ContaContrapartida1 = new List<string>();
                for (var i = 0; i < passo.Count; i++)
                {
                    string contcontra = Convert.ToString(passo[i].CONTA_CONTABIL_SAIDA);
                    ContaContrapartida1.Add(contcontra);
                }
                string deposito = Convert.ToString(tblmovimentacao.DEPOSITO);
                string dataInicial = Convert.ToString(tblmovimentacao.DATA_INICIAL);
                string dataFinal = Convert.ToString(tblmovimentacao.DATA_FINAL);
                string Cdfilial = Convert.ToString(tblmovimentacao.CDFILIAL);
                string numDocmento = Convert.ToString(tblmovimentacao.NUM_DOCUMENTO);
                string Reduzido = Convert.ToString(tblmovimentacao.REDUZIDO_ITEM);
                string TipoMovimento = Convert.ToString(tblmovimentacao.NUM_TIPO_MOVIMENTO);
                string Situacao = Convert.ToString(tblmovimentacao.SITUACAO);
                string Opecacao = Convert.ToString(tblmovimentacao.OPERACAO);
                HttpWebRequest req = WebRequest.CreateHttp("https://localhost:44319/Formularios/PlanilhaContabilPasso2/");
                req.Method = "GET";
                req.Headers.Add("deposito", deposito);
                req.Headers.Add("dataInicial", dataInicial);
                req.Headers.Add("dataFinal", dataFinal);
                req.Headers.Add("Cdfilial", Cdfilial);
                req.Headers.Add("numDocmento", numDocmento);
                req.Headers.Add("Reduzido", Reduzido);
                req.Headers.Add("TipoMovimento", TipoMovimento);
                req.Headers.Add("Situacao", Situacao);
                req.Headers.Add("Opecacao", Opecacao);
                documento1 = req;

                return View(passo);
            }
            else
            {
                TempData["dataerrada"] = "sim";
                return RedirectToAction(nameof(PlanilhaContabilPasso2));
            }

        }
        public ActionResult Update(Tblmovimentacao tblmovimentacao, string caminho1, int Linhas)
        {
            //var idusuario = TempData.Peek("Pessoa");
            var USUARIO = TempData["USUARIO"];


            var idusuario = ((System.Net.HttpWebRequest)documento).Headers.ToString().Split("caminho1: ")[0].Split("idusuario: ")[1].Trim();
            caminho1 = ((System.Net.HttpWebRequest)documento).Headers.ToString().Split("caminho1: ")[1].Split("Linha: ")[0].Trim();
            var mov = ((System.Net.HttpWebRequest)documento).Headers.ToString().Split("caminho1: ")[0].Split("idusuario: ")[0].Split("mov: ")[1].Trim();
            var movimentos = mov.ToString().Replace("[", " ").Replace("]", " ").Trim();
            var Linha = ((System.Net.HttpWebRequest)documento).Headers.ToString().Split("caminho1: ")[1].Split("Linha: ")[1].Trim();
            Linhas = Convert.ToInt32(Linha);

            ConexaoOraFormularios excel = new ConexaoOraFormularios();
            var form = excel.ExcelTipoMovimentoContabil(caminho1, Linhas);


            var deposito = ((System.Net.HttpWebRequest)documento1).Headers.ToString().Split("Opecacao: ")[0].Split("Situacao: ")[0].Split("TipoMovimento: ")[0].Split("Reduzido: ")[0].Split("numDocmento: ")[0].Split("Cdfilial: ")[0].Split("dataFinal: ")[0].Split("dataInicial: ")[0].Split("deposito: ")[1].Trim();
            var dataInicial = ((System.Net.HttpWebRequest)documento1).Headers.ToString().Split("Opecacao: ")[0].Split("Situacao: ")[0].Split("TipoMovimento: ")[0].Split("Reduzido: ")[0].Split("numDocmento: ")[0].Split("Cdfilial: ")[0].Split("dataFinal: ")[0].Split("dataInicial: ")[1].Trim();
            var dataFinal = ((System.Net.HttpWebRequest)documento1).Headers.ToString().Split("Opecacao: ")[0].Split("Situacao: ")[0].Split("TipoMovimento: ")[0].Split("Reduzido: ")[0].Split("numDocmento: ")[0].Split("Cdfilial: ")[0].Split("dataFinal: ")[1].Trim();
            var Cdfilial = ((System.Net.HttpWebRequest)documento1).Headers.ToString().Split("Opecacao: ")[0].Split("Situacao: ")[0].Split("TipoMovimento: ")[0].Split("Reduzido: ")[0].Split("numDocmento: ")[0].Split("Cdfilial: ")[1].Trim();
            var numDocmento = ((System.Net.HttpWebRequest)documento1).Headers.ToString().Split("Opecacao: ")[0].Split("Situacao: ")[0].Split("TipoMovimento: ")[0].Split("Reduzido: ")[0].Split("numDocmento: ")[1].Trim();
            var Reduzido = ((System.Net.HttpWebRequest)documento1).Headers.ToString().Split("Opecacao: ")[0].Split("Situacao: ")[0].Split("TipoMovimento: ")[0].Split("Reduzido: ")[1].Trim();
            var TipoMovimento = ((System.Net.HttpWebRequest)documento1).Headers.ToString().Split("Opecacao: ")[0].Split("Situacao: ")[0].Split("TipoMovimento: ")[1].Trim();
            var Situacao = ((System.Net.HttpWebRequest)documento1).Headers.ToString().Split("Opecacao: ")[0].Split("Situacao: ")[1].Trim();
            var Opecacao = ((System.Net.HttpWebRequest)documento1).Headers.ToString().Split("Opecacao: ")[1].Trim();

            tblmovimentacao.DEPOSITO = Convert.ToInt32(deposito);
            tblmovimentacao.DATA_INICIAL = Convert.ToDateTime(dataInicial);
            tblmovimentacao.DATA_FINAL = Convert.ToDateTime(dataFinal);
            tblmovimentacao.CDFILIAL = Convert.ToInt32(Cdfilial);
            tblmovimentacao.NUM_DOCUMENTO = Convert.ToInt32(numDocmento);
            tblmovimentacao.REDUZIDO_ITEM = Convert.ToInt32(Reduzido);
            tblmovimentacao.NUM_TIPO_MOVIMENTO = Convert.ToInt32(TipoMovimento);
            tblmovimentacao.SITUACAO = Convert.ToInt32(Situacao);
            tblmovimentacao.OPERACAO = Convert.ToInt32(Opecacao);


            //bolar um jeito de apagar o static documento sem comprometer o F5 de atualização da pagina
            ConexaoOraFormularios conexao = new ConexaoOraFormularios();
            var passo = conexao.Contabils2(tblmovimentacao, movimentos);

            ViewData["Exibir"] = "sim";
            ViewData["exibir login"] = "sim";

            List<int> verdade = new List<int>();
            List<int> falso = new List<int>();
            List<string> pass = new List<string>();
            // tem na pesquisa no banco porem não tem na planilha do excel
            foreach (var passo1 in passo)
            {

                //verdade = null;
                for (int reduz = 0; reduz < Linhas; reduz++)
                {
                    //olhar final da linha está comentado
                    if (passo1.REDUZIDO_ITEM.ToString().Contains("" + form[reduz].Reduzido + "") == true && passo1.DEPOSITO.ToString().Contains("" + form[reduz].Deposito + "") == true && passo1.NUM_TIPO_MOVIMENTO.ToString().Contains("" + form[reduz].TipoMov.Split("-FAT")[0] + "") == true && passo1.DATA_DOCUMENTO.ToString().Contains("" + form[reduz].Data + "") == true && passo1.NUMERO_MOVIMENTO.ToString().Contains("" + form[reduz].NumDocto + "") == true /*&& passo1.CONTA_CONTABIL.ToString().Contains("" + form[reduz].ContaEstoque + "") == true && passo1.CONTA_CONTABIL_SAIDA.ToString().Contains("" + form[reduz].ContaContrapartida + "") == true*/)
                    {

                        verdade.Add(reduz);
                    }
                    else
                    {

                        falso.Add(reduz);

                    }
                }
                pass.Add(Convert.ToString(passo1));
            }
            var fals = falso.Distinct().ToList();
            var veda = verdade.Distinct().ToList();
            foreach (var ved in veda)
            {
                if (fals.Contains(ved))
                {
                    fals.Remove(ved);
                }
            }
            var fal = fals;
            List<int> fa = new List<int>();
            for (int i = 0; i < fal.Count; i++)
            {
                var f = fal[i] + 2;
                fa.Add(f);
            }
            if (fal.Count > 0)
            {
                ViewData["linhaPosicao"] = "Linhas " + String.Join(", ", fa) + " da planilha Excel não estão na consulta";
               
            }


            //testar ainda 
            List<int> verdade1 = new List<int>();
            List<int> falso1 = new List<int>();
            List<string> pass1 = new List<string>();
            // tem na planilha do excel porem não  tem na pesquisa no banco 
            foreach (var form1 in form)
            {

                //verdade = null;
                for (int reduz = 0; reduz < Linhas; reduz++)
                {
                    //olhar final da linha está comentado
                    if (form1.Reduzido.ToString().Contains("" + passo[reduz].REDUZIDO_ITEM + "") == true && form1.Deposito.ToString().Contains("" + passo[reduz].DEPOSITO + "") == true && form1.TipoMov.Split("-FAT")[0].ToString().Contains("" + passo[reduz].NUM_TIPO_MOVIMENTO + "") == true && form1.Data.ToString().Contains("" + passo[reduz].DATA_DOCUMENTO + "") == true && form1.NumDocto.ToString().Contains("" + passo[reduz].NUMERO_MOVIMENTO + "") == true /*&& form1.ContaEstoque.ToString().Contains("" + passo[reduz].CONTA_CONTABIL + "") == true && form1.ContaContrapartida.ToString().Contains("" + passo[reduz].CONTA_CONTABIL_SAIDA + "") == true*/)
                    {

                        verdade1.Add(reduz);
                    }
                    else
                    {

                        falso1.Add(reduz);

                    }
                }
                pass1.Add(Convert.ToString(form1));
            }
            var fals1 = falso1.Distinct().ToList();
            var veda1 = verdade1.Distinct().ToList();
            foreach (var ved in veda)
            {
                if (fals1.Contains(ved))
                {
                    fals1.Remove(ved);
                }
            }
            var fal1 = fals;
            List<int> fa1 = new List<int>();
            for (int i = 0; i < fal.Count; i++)
            {
                var f = fal[i] + 2;
                fa.Add(f);
            }


            return View();
        }

        public ActionResult PesquisaAvancada()
        {
            TempData["PesquisaAvancada"] = "sim";
            return RedirectToAction(nameof(PlanilhaContabilPasso2));
        }


        public ActionResult PlanilhaTipoMovimentoAleterado(TipoMovimento tipo, object caminho)
        {
            TempData.Remove("redirect");
            var idusuario = TempData["Pessoa"];
            ViewData["exibir login"] = "sim";
            tipo.DE = (string)TempData["de"];
            tipo.PARA = (string)TempData["para"];
            if (idusuario != null)
            {

                int perfil = Convert.ToInt32(TempData["Pessoa"]);
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
                        var form = excel.MovimentoAlterado(tipo, caminho1);
                        if (form == null)
                        {
                            TempData["redirect"] = "sim";
                            return RedirectToAction(nameof(PlanilhaTipoMovimento));
                        }
                        TempData["Pessoa"] = idusuario;
                        ViewData["DE"] = tipo.DE;
                        ViewData["PARA"] = tipo.PARA;
                        return View();
                    }
                    else
                    {
                        return RedirectToAction("Home", "Pessoa");
                    }
                }
                else
                {
                    //desativado por enquanto
                    return RedirectToAction(nameof(EnviarEmail));
                }


            }
            else
            {
                return RedirectToAction("Login", "Pessoa");
            }

        }


        //desativado por enquanto
        public ActionResult NaoConfirmado()
        {
            var idusuario = TempData["Pessoa"];
            if (idusuario != null)
            {
                return RedirectToAction("Home", "Pessoa");
            }
            else
            {
                return RedirectToAction("Login", "Pessoa");
            }

        }
        //desativado por enquanto
        public ActionResult EnviarEmail()
        {

            ViewData["exibir login"] = "sim";
            ViewData["Exibir"] = "sim";
            return RedirectToAction("EnviarEmail", "Pessoa");
        }

    }
}
