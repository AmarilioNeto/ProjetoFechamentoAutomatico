using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc.Rendering;
using Oracle.ManagedDataAccess.Client;
using System.Data.OracleClient;
using ProjetoFechamentoAutomatico.Models.Formularios;
using ProjetoFechamentoAutomatico.Models.Pessoa;
using System.Data.OleDb;
using System.Data;
using System.IO;
using OfficeOpenXml;
using System.Data.OracleClient;
using System.Runtime;
using Microsoft.VisualBasic;


namespace ProjetoFechamentoAutomatico.Models.Formularios
{
    public class ConexaoOraFormularios
    {
        public string conexaoOracle = @"DATA SOURCE=(DESCRIPTION=(ADDRESS_LIST=
(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.1.29)(PORT=1521)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=tubdhmlg)));User Id=*****; Password=****;";

        // conexão com base de produção
        //        public string conexaoOracle = @"DATA SOURCE=(DESCRIPTION =(ADDRESS_LIST =
        //(ADDRESS = (PROTOCOL = TCP)(HOST = 192.168.1.10)(PORT = 1521)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME = tubdprd)));User Id = *****;Password = ******;";


        public OracleConnection conexaoOra;
        public OracleCommand Command;
        public OracleParameter parameter;
        public OracleParameterCollection parametros;

        public List<DDCVI> Excel(string caminho1)
        {

            List<DDCVI> formulario = new List<DDCVI>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // captura o caminho temporario do arquivo
            var packager = new ExcelPackage(new FileInfo(caminho1));

            var workbook = packager.Workbook;
            var sheet = workbook.Worksheets[0];
            object cells = sheet.Cells["A3:O184"].Value;

            if (sheet.Cells["A1"].Value == null || sheet.Cells["B1"].Value == null || sheet.Cells["C1"].Value == null || sheet.Cells["D1"].Value == null ||
                sheet.Cells["E1"].Value == null || sheet.Cells["F1"].Value == null || sheet.Cells["G1"].Value == null || sheet.Cells["H1"].Value == null ||
                sheet.Cells["I1"].Value == null || sheet.Cells["J1"].Value == null || sheet.Cells["K1"].Value == null || sheet.Cells["L1"].Value == null ||
                sheet.Cells["M1"].Value == null || sheet.Cells["N1"].Value == null || sheet.Cells["O1"].Value == null || sheet.Cells["P1"].Value != null)
            {
                formulario = null;
                return formulario;
            }
            else
            {


                for (int cell = 0; cell < 181; cell++)
                {
                    var red = ((object[,])cells)[cell, 0];
                    var codigo = ((object[,])cells)[cell, 1];
                    var unid = ((object[,])cells)[cell, 2];
                    var descricao = ((object[,])cells)[cell, 3];
                    var dep = ((object[,])cells)[cell, 4];
                    var mov = ((object[,])cells)[cell, 5];
                    var data = ((object[,])cells)[cell, 6];
                    var docto = ((object[,])cells)[cell, 7];
                    var forne = ((object[,])cells)[cell, 8];
                    var quantE = ((object[,])cells)[cell, 9];
                    var unit = ((object[,])cells)[cell, 10];
                    var entradas = ((object[,])cells)[cell, 11];
                    var quantS = ((object[,])cells)[cell, 12];
                    var unitS = ((object[,])cells)[cell, 13];
                    var saidas = ((object[,])cells)[cell, 14];



                    DDCVI dDCVI = new DDCVI();
                    if (red != null)
                    {
                        dDCVI.Red = red.ToString();
                    }
                    else
                    {
                        dDCVI.Red = null;
                    }
                    if (codigo != null)
                    {
                        dDCVI.Codigo = codigo.ToString();
                    }
                    else
                    {
                        dDCVI.Codigo = null;
                    }
                    if (unid != null)
                    {
                        dDCVI.Unid = unid.ToString();
                    }
                    else
                    {
                        dDCVI.Unid = null;
                    }
                    if (descricao != null)
                    {
                        dDCVI.Descricao = descricao.ToString();
                    }
                    else
                    {
                        dDCVI.Descricao = null;
                    }
                    if (dep != null)
                    {
                        dDCVI.Dep = dep.ToString();
                    }
                    else
                    {
                        dDCVI.Dep = null;
                    }
                    if (mov != null)
                    {
                        dDCVI.Mov = mov.ToString();
                    }
                    else
                    {
                        dDCVI.Mov = null;
                    }
                    if (data != null)
                    {
                        dDCVI.Data = data.ToString();
                    }
                    else
                    {
                        dDCVI.Data = null;
                    }
                    if (docto != null)
                    {
                        dDCVI.Docto = docto.ToString();
                    }
                    else
                    {
                        dDCVI.Docto = null;
                    }
                    if (forne != null)
                    {
                        dDCVI.Forne = forne.ToString();
                    }
                    else
                    {
                        dDCVI.Forne = null;
                    }
                    if (quantE != null)
                    {
                        dDCVI.QuantE = quantE.ToString();
                    }
                    else
                    {
                        dDCVI.QuantE = null;
                    }
                    if (unit != null)
                    {
                        dDCVI.Unit = unit.ToString();
                    }
                    else
                    {
                        dDCVI.Unit = null;
                    }
                    if (entradas != null)
                    {
                        dDCVI.Entradas = entradas.ToString();
                    }
                    else
                    {
                        dDCVI.Entradas = null;
                    }
                    if (quantS != null)
                    {
                        dDCVI.QuantS = quantS.ToString();
                    }
                    else
                    {
                        dDCVI.QuantS = null;
                    }
                    if (unitS != null)
                    {
                        dDCVI.UnitS = unitS.ToString();
                    }
                    else
                    {
                        dDCVI.UnitS = null;
                    }
                    if (saidas != null)
                    {
                        dDCVI.Saidas = saidas.ToString();
                    }
                    else
                    {
                        dDCVI.Saidas = null;
                    }


                    formulario.Add(dDCVI);
                }

                return formulario;
            }

        }
        public List<TipoMovimento> ExcelTipoMovimento(string caminho1)
        {

            List<TipoMovimento> formulario = new List<TipoMovimento>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // captura o caminho temporario do arquivo
            var packager = new ExcelPackage(new FileInfo(caminho1));

            var workbook = packager.Workbook;
            var sheet = workbook.Worksheets[0];
            object cells = sheet.Cells["A2:AA1500"].Value;

            if (sheet.Cells["A1"].Value == null || sheet.Cells["B1"].Value == null || sheet.Cells["C1"].Value == null || sheet.Cells["D1"].Value == null ||
               sheet.Cells["E1"].Value == null || sheet.Cells["F1"].Value == null || sheet.Cells["G1"].Value == null || sheet.Cells["H1"].Value == null ||
               sheet.Cells["I1"].Value == null || sheet.Cells["J1"].Value == null || sheet.Cells["K1"].Value == null || sheet.Cells["L1"].Value == null ||
               sheet.Cells["M1"].Value == null || sheet.Cells["N1"].Value == null || sheet.Cells["O1"].Value == null || sheet.Cells["P1"].Value == null ||
               sheet.Cells["Q1"].Value == null || sheet.Cells["R1"].Value == null || sheet.Cells["S1"].Value == null || sheet.Cells["T1"].Value == null ||
               sheet.Cells["U1"].Value == null || sheet.Cells["V1"].Value == null || sheet.Cells["W1"].Value == null || sheet.Cells["X1"].Value == null ||
               sheet.Cells["Y1"].Value == null || sheet.Cells["Z1"].Value == null || sheet.Cells["AA1"].Value == null || sheet.Cells["AB1"].Value != null)
            {
                formulario = null;
                return formulario;
            }
            else
            {
                for (int cell = 1; cell < 1499; cell++)
                {
                    if (sheet.Cells["F" + cell + ""].Value != null)
                    {
                        var linha = cell;


                    }
                    else
                    {
                        break;
                    }
                }

                for (int cell = 0; cell < 1499; cell++)
                {
                    var reduzido = ((object[,])cells)[cell, 0];
                    var codigoIndustrial = ((object[,])cells)[cell, 1];
                    var unidMed = ((object[,])cells)[cell, 2];
                    var descricao = ((object[,])cells)[cell, 3];
                    var deposito = ((object[,])cells)[cell, 4];
                    var tipoMov = ((object[,])cells)[cell, 5];
                    var data = ((object[,])cells)[cell, 6];
                    var numDocto = ((object[,])cells)[cell, 7];
                    var fornecedor = ((object[,])cells)[cell, 8];
                    var quantEntradas = ((object[,])cells)[cell, 9];
                    var entraUnit = ((object[,])cells)[cell, 10];
                    var entradas = ((object[,])cells)[cell, 11];
                    var quantSaidas = ((object[,])cells)[cell, 12];
                    var saidaUnit = ((object[,])cells)[cell, 13];
                    var saidas = ((object[,])cells)[cell, 14];
                    var serieNFE = ((object[,])cells)[cell, 15];
                    var naturezaOperacao = ((object[,])cells)[cell, 16];
                    var destacaICMS = ((object[,])cells)[cell, 17];
                    var destacaPIS = ((object[,])cells)[cell, 18];
                    var destacaConfis = ((object[,])cells)[cell, 19];
                    var destacaIPI = ((object[,])cells)[cell, 20];
                    var destacaFrete = ((object[,])cells)[cell, 21];
                    var icmsFrete = ((object[,])cells)[cell, 22];
                    var pisFrete = ((object[,])cells)[cell, 23];
                    var confinsFrete = ((object[,])cells)[cell, 24];
                    var contaEstoque = ((object[,])cells)[cell, 25];
                    var contaContrapartida = ((object[,])cells)[cell, 26];

                    TipoMovimento tipoMovimento = new TipoMovimento();
                    if (reduzido != null)
                    {
                        tipoMovimento.Reduzido = reduzido.ToString();
                    }
                    else
                    {
                        tipoMovimento.Reduzido = null;
                    }
                    if (codigoIndustrial != null)
                    {
                        tipoMovimento.CodigoIndustrial = codigoIndustrial.ToString();
                    }
                    else
                    {
                        tipoMovimento.CodigoIndustrial = null;
                    }
                    if (unidMed != null)
                    {
                        tipoMovimento.UnidMed = unidMed.ToString();
                    }
                    else
                    {
                        tipoMovimento.UnidMed = null;
                    }
                    if (descricao != null)
                    {
                        tipoMovimento.Descricao = descricao.ToString();
                    }
                    else
                    {
                        tipoMovimento.Descricao = null;
                    }
                    if (deposito != null)
                    {
                        tipoMovimento.Deposito = deposito.ToString();
                    }
                    else
                    {
                        tipoMovimento.Deposito = null;
                    }
                    if (tipoMov != null)
                    {
                        tipoMovimento.TipoMov = tipoMov.ToString();
                    }
                    else
                    {
                        tipoMovimento.TipoMov = null;
                    }
                    if (data != null)
                    {
                        tipoMovimento.Data = data.ToString();
                    }
                    else
                    {
                        tipoMovimento.Data = null;
                    }
                    if (numDocto != null)
                    {
                        tipoMovimento.NumDocto = numDocto.ToString();
                    }
                    else
                    {
                        tipoMovimento.NumDocto = null;
                    }
                    if (fornecedor != null)
                    {
                        tipoMovimento.Fornecedor = fornecedor.ToString();
                    }
                    else
                    {
                        tipoMovimento.Fornecedor = null;
                    }
                    if (quantEntradas != null)
                    {
                        tipoMovimento.QuantEntradas = quantEntradas.ToString();
                    }
                    else
                    {
                        tipoMovimento.QuantEntradas = null;
                    }
                    if (entraUnit != null)
                    {
                        tipoMovimento.EntraUnit = entraUnit.ToString();
                    }
                    else
                    {
                        tipoMovimento.EntraUnit = null;
                    }
                    if (entradas != null)
                    {
                        tipoMovimento.Entradas = entradas.ToString();
                    }
                    else
                    {
                        tipoMovimento.Entradas = null;
                    }
                    if (quantSaidas != null)
                    {
                        tipoMovimento.QuantSaidas = quantSaidas.ToString();
                    }
                    else
                    {
                        tipoMovimento.QuantSaidas = null;
                    }
                    if (saidaUnit != null)
                    {
                        tipoMovimento.SaidaUnit = saidaUnit.ToString();
                    }
                    else
                    {
                        tipoMovimento.SaidaUnit = null;
                    }
                    if (saidas != null)
                    {
                        tipoMovimento.Saidas = saidas.ToString();
                    }
                    else
                    {
                        tipoMovimento.Saidas = null;
                    }

                    if (serieNFE != null)
                    {
                        tipoMovimento.SerieNFE = serieNFE.ToString();
                    }
                    else
                    {
                        tipoMovimento.SerieNFE = null;
                    }

                    if (naturezaOperacao != null)
                    {
                        tipoMovimento.NaturezaOperacao = naturezaOperacao.ToString();
                    }
                    else
                    {
                        tipoMovimento.NaturezaOperacao = null;
                    }

                    if (destacaICMS != null)
                    {
                        tipoMovimento.DestacaICMS = destacaICMS.ToString();
                    }
                    else
                    {
                        tipoMovimento.DestacaICMS = null;
                    }

                    if (destacaPIS != null)
                    {
                        tipoMovimento.DestacaPIS = destacaPIS.ToString();
                    }
                    else
                    {
                        tipoMovimento.DestacaPIS = null;
                    }

                    if (destacaConfis != null)
                    {
                        tipoMovimento.DestacaCofins = destacaConfis.ToString();
                    }
                    else
                    {
                        tipoMovimento.DestacaCofins = null;
                    }

                    if (destacaIPI != null)
                    {
                        tipoMovimento.DestacaIPI = destacaIPI.ToString();
                    }
                    else
                    {
                        tipoMovimento.DestacaIPI = null;
                    }

                    if (destacaFrete != null)
                    {
                        tipoMovimento.DestacaFrete = destacaFrete.ToString();
                    }
                    else
                    {
                        tipoMovimento.DestacaFrete = null;
                    }

                    if (icmsFrete != null)
                    {
                        tipoMovimento.ICMSFrete = icmsFrete.ToString();
                    }
                    else
                    {
                        tipoMovimento.ICMSFrete = null;
                    }

                    if (pisFrete != null)
                    {
                        tipoMovimento.PISFrete = pisFrete.ToString();
                    }
                    else
                    {
                        tipoMovimento.PISFrete = null;
                    }

                    if (confinsFrete != null)
                    {
                        tipoMovimento.CofinsFrete = confinsFrete.ToString();
                    }
                    else
                    {
                        tipoMovimento.CofinsFrete = null;
                    }

                    if (contaEstoque != null)
                    {
                        tipoMovimento.ContaEstoque = contaEstoque.ToString();
                    }
                    else
                    {
                        tipoMovimento.ContaEstoque = null;
                    }

                    if (contaContrapartida != null)
                    {
                        tipoMovimento.ContaContrapartida = contaContrapartida.ToString();
                    }
                    else
                    {
                        tipoMovimento.ContaContrapartida = null;
                    }
                    formulario.Add(tipoMovimento);
                }

                return formulario;
            }

        }

        public List<Formularios> ExcelAlterado(string caminho1)
        {
            List<Formularios> formularios = new List<Formularios>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // mudar para deixar o caminho da pasta dinamico.
            var packager = new ExcelPackage(new FileInfo(caminho1));

            var workbook = packager.Workbook;
            var sheet = workbook.Worksheets[0];
            object cells = sheet.Cells["A3:L184"].Value;
            for (int cell = 0; cell <= 181; cell++)
            {
                //capturando valores das celulas
                var red = ((object[,])cells)[cell, 0];
                var codigo = ((object[,])cells)[cell, 1];
                var unid = ((object[,])cells)[cell, 2];
                var descricao = ((object[,])cells)[cell, 3];
                var dep = ((object[,])cells)[cell, 4];
                var mov = ((object[,])cells)[cell, 5];
                var data = ((object[,])cells)[cell, 6];
                var docto = ((object[,])cells)[cell, 7];
                var forne = ((object[,])cells)[cell, 8];
                var quantE = ((object[,])cells)[cell, 9];
                var unit = ((object[,])cells)[cell, 10];
                var entradas = ((object[,])cells)[cell, 11];
                string Unit1 = "" + unit;

                if (red != null && unit != null && Unit1 != " ")
                {
                    Formularios formularios1 = new Formularios();
                    if (red != null)
                    {
                        formularios1.Red = red.ToString();
                    }
                    else
                    {
                        formularios1.Red = null;
                    }
                    if (codigo != null)
                    {
                        formularios1.Codigo = codigo.ToString();
                    }
                    else
                    {
                        formularios1.Codigo = null;
                    }
                    if (unid != null)
                    {
                        formularios1.Unid = unid.ToString();
                    }
                    else
                    {
                        formularios1.Unid = null;
                    }
                    if (descricao != null)
                    {
                        formularios1.Descricao = descricao.ToString();
                    }
                    else
                    {
                        formularios1.Descricao = null;
                    }
                    if (dep != null)
                    {
                        formularios1.Dep = dep.ToString();
                    }
                    else
                    {
                        formularios1.Dep = null;
                    }
                    if (mov != null)
                    {
                        formularios1.Mov = mov.ToString();
                    }
                    else
                    {
                        formularios1.Mov = null;
                    }
                    if (data != null)
                    {
                        formularios1.Data = data.ToString();
                    }
                    else
                    {
                        formularios1.Data = null;
                    }
                    if (docto != null)
                    {
                        formularios1.Docto = docto.ToString();
                    }
                    else
                    {
                        formularios1.Docto = null;
                    }
                    if (forne != null)
                    {
                        formularios1.Forne = forne.ToString();
                    }
                    else
                    {
                        formularios1.Forne = null;
                    }
                    if (quantE != null)
                    {
                        formularios1.QuantE = quantE.ToString();
                    }
                    else
                    {
                        formularios1.QuantE = null;
                    }
                    if (unit != null)
                    {
                        formularios1.Unit = unit.ToString();
                    }
                    else
                    {
                        formularios1.Unit = null;
                    }
                    if (entradas != null)
                    {
                        formularios1.Entradas = entradas.ToString();
                    }
                    else
                    {
                        formularios1.Entradas = null;
                    }
                    formularios.Add(formularios1);

                }

            }

            return formularios;
        }

        public List<Formularios> AtualConfirmado(string caminho1)
        {
            List<Formularios> formularios = new List<Formularios>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var packager = new ExcelPackage(new FileInfo(caminho1));

            var workbook = packager.Workbook;
            var sheet = workbook.Worksheets[0];
            object cells = sheet.Cells["A3:L184"].Value;
            for (int cell = 0; cell <= 181; cell++)
            {
                //capturando valores das celulas
                var red = ((object[,])cells)[cell, 0];
                var codigo = ((object[,])cells)[cell, 1];
                var unid = ((object[,])cells)[cell, 2];
                var descricao = ((object[,])cells)[cell, 3];
                var dep = ((object[,])cells)[cell, 4];
                var mov = ((object[,])cells)[cell, 5];
                var data = ((object[,])cells)[cell, 6];
                var docto = ((object[,])cells)[cell, 7];
                var forne = ((object[,])cells)[cell, 8];
                var quantE = ((object[,])cells)[cell, 9];
                var unit = ((object[,])cells)[cell, 10];
                var entradas = ((object[,])cells)[cell, 11];
                string Unit1 = "" + unit;

                if (red != null && unit != null && Unit1 != " ")
                {

                    conexaoOra = new OracleConnection(conexaoOracle);
                    Command = conexaoOra.CreateCommand();
                    Command.CommandText = @"select * from movimentacao m 
                    where m.deposito= " + dep + " and m.documento_sintetico = " + docto + " and m.reduzido_item=" + red + "";
                    conexaoOra.Open();
                    OracleDataReader dr = Command.ExecuteReader();
                    dr.Read();

                    if (dr.HasRows)
                    {
                        Command.CommandText = @" update movimentacao m set m.preco_medio = " + Unit1.Replace(',', '.') + ", m.preco = " + Unit1.Replace(',', '.') + "" +
                   " where m.deposito= " + dep + " and m.documento_sintetico= " + docto + "  and m.reduzido_item=" + red + "";
                        Command.ExecuteNonQuery();
                    }

                    Formularios formularios1 = new Formularios();
                    if (red != null)
                    {
                        formularios1.Red = red.ToString();
                    }
                    else
                    {
                        formularios1.Red = null;
                    }
                    if (codigo != null)
                    {
                        formularios1.Codigo = codigo.ToString();
                    }
                    else
                    {
                        formularios1.Codigo = null;
                    }
                    if (unid != null)
                    {
                        formularios1.Unid = unid.ToString();
                    }
                    else
                    {
                        formularios1.Unid = null;
                    }
                    if (descricao != null)
                    {
                        formularios1.Descricao = descricao.ToString();
                    }
                    else
                    {
                        formularios1.Descricao = null;
                    }
                    if (dep != null)
                    {
                        formularios1.Dep = dep.ToString();
                    }
                    else
                    {
                        formularios1.Dep = null;
                    }
                    if (mov != null)
                    {
                        formularios1.Mov = mov.ToString();
                    }
                    else
                    {
                        formularios1.Mov = null;
                    }
                    if (data != null)
                    {
                        formularios1.Data = data.ToString();
                    }
                    else
                    {
                        formularios1.Data = null;
                    }
                    if (docto != null)
                    {
                        formularios1.Docto = docto.ToString();
                    }
                    else
                    {
                        formularios1.Docto = null;
                    }
                    if (forne != null)
                    {
                        formularios1.Forne = forne.ToString();
                    }
                    else
                    {
                        formularios1.Forne = null;
                    }
                    if (quantE != null)
                    {
                        formularios1.QuantE = quantE.ToString();
                    }
                    else
                    {
                        formularios1.QuantE = null;
                    }
                    if (unit != null)
                    {
                        formularios1.Unit = unit.ToString();
                    }
                    else
                    {
                        formularios1.Unit = null;
                    }
                    if (entradas != null)
                    {
                        formularios1.Entradas = entradas.ToString();
                    }
                    else
                    {
                        formularios1.Entradas = null;
                    }
                    formularios.Add(formularios1);

                    dr.Close();
                    conexaoOra.Close();
                }

            }

            return formularios;
        }
        public List<TipoMovimento> MovimentoAlterado(TipoMovimento tipo, object caminho1)
        {
            List<TipoMovimento> movimentos = new List<TipoMovimento>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var packager = new ExcelPackage(new FileInfo((string)caminho1));

            var workbook = packager.Workbook;
            var sheet = workbook.Worksheets[0];
            object cells = sheet.Cells["A2:AA1500"].Value;

            if (sheet.Cells["A1"].Value == null || sheet.Cells["B1"].Value == null || sheet.Cells["C1"].Value == null || sheet.Cells["D1"].Value == null ||
               sheet.Cells["E1"].Value == null || sheet.Cells["F1"].Value == null || sheet.Cells["G1"].Value == null || sheet.Cells["H1"].Value == null ||
               sheet.Cells["I1"].Value == null || sheet.Cells["J1"].Value == null || sheet.Cells["K1"].Value == null || sheet.Cells["L1"].Value == null ||
               sheet.Cells["M1"].Value == null || sheet.Cells["N1"].Value == null || sheet.Cells["O1"].Value == null || sheet.Cells["P1"].Value == null ||
               sheet.Cells["Q1"].Value == null || sheet.Cells["R1"].Value == null || sheet.Cells["S1"].Value == null || sheet.Cells["T1"].Value == null ||
               sheet.Cells["U1"].Value == null || sheet.Cells["V1"].Value == null || sheet.Cells["W1"].Value == null || sheet.Cells["X1"].Value == null ||
               sheet.Cells["Y1"].Value == null || sheet.Cells["Z1"].Value == null || sheet.Cells["AA1"].Value == null || sheet.Cells["AB1"].Value != null)
            {
                movimentos = null;
                return movimentos;
            }
            else
            {
                for (int cell = 0; cell < 1499; cell++)
                {
                    var reduzido = ((object[,])cells)[cell, 0];
                    var deposito = ((object[,])cells)[cell, 4];
                    var tipoMov = ((object[,])cells)[cell, 5];
                    var numDocto = ((object[,])cells)[cell, 7];
                    if (reduzido != null && deposito != null && numDocto != null && tipoMov != null)
                    {

                        conexaoOra = new OracleConnection(conexaoOracle);
                        Command = conexaoOra.CreateCommand();
                        Command.CommandText = @"select * from movimentacao m 
                    where m.deposito= " + deposito + " and m.documento_sintetico = " + numDocto + " and m.reduzido_item=" + reduzido + " and m.num_tipo_movimento = " + tipo.DE + "";
                        conexaoOra.Open();
                        OracleDataReader dr = Command.ExecuteReader();
                        dr.Read();

                        if (dr.HasRows)
                        {
                            try
                            {
                                Command.CommandText = @" update movimentacao m set m.num_tipo_movimento = " + tipo.PARA + "" +
                      " where m.deposito= " + deposito + " and m.documento_sintetico= " + numDocto + "  and m.reduzido_item=" + reduzido + " and m.num_tipo_movimento = " + tipo.DE + "";

                                Command.ExecuteNonQuery();
                                TipoMovimento tipoMovimento = new TipoMovimento();
                                movimentos.Add(tipoMovimento);
                                dr.Close();
                                conexaoOra.Close();
                            }
                            catch
                            {
                                movimentos = null;
                                dr.Close();
                                conexaoOra.Close();
                                return movimentos;
                            }


                        }

                    }


                }
            }
            return movimentos;
        }
        public List<TipoMovimentoContabil> ExcelTipoMovimentoContabil(string caminho1, int Linhas)
        {
            List<TipoMovimentoContabil> formulario = new List<TipoMovimentoContabil>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // captura o caminho temporario do arquivo
            var packager = new ExcelPackage(new FileInfo(caminho1));

            var workbook = packager.Workbook;
            var sheet = workbook.Worksheets[0];
            object cells = sheet.Cells["A2:S1500"].Value;
            if (sheet.Cells["A1"].Value == null || sheet.Cells["B1"].Value == null || sheet.Cells["C1"].Value == null || sheet.Cells["D1"].Value == null ||
              sheet.Cells["E1"].Value == null || sheet.Cells["F1"].Value == null || sheet.Cells["G1"].Value == null || sheet.Cells["H1"].Value == null ||
              sheet.Cells["I1"].Value == null || sheet.Cells["J1"].Value == null || sheet.Cells["K1"].Value == null || sheet.Cells["L1"].Value == null ||
              sheet.Cells["M1"].Value == null || sheet.Cells["N1"].Value == null || sheet.Cells["O1"].Value == null || sheet.Cells["P1"].Value == null ||
              sheet.Cells["Q1"].Value == null || sheet.Cells["R1"].Value == null || sheet.Cells["S1"].Value == null || sheet.Cells["T1"].Value != null)
            {
                formulario = null;
                return formulario;
            }
            else
            {
                for (int cell = 0; cell < 1499; cell++)
                {
                    var cell1 = cell + 1;
                    if (sheet.Cells["F" + cell1 + ""].Value != null)
                    {
                        Linhas = cell1 - 1;

                    }
                    else
                    {
                        break;
                    }
                }

                for (int cell = 0; cell < 1499; cell++)
                {


                    var reduzido = ((object[,])cells)[cell, 0];
                    var codigoIndustrial = ((object[,])cells)[cell, 1];
                    var unidMed = ((object[,])cells)[cell, 2];
                    var descricao = ((object[,])cells)[cell, 3];
                    var deposito = ((object[,])cells)[cell, 4];
                    var tipoMov = ((object[,])cells)[cell, 5];
                    var data = ((object[,])cells)[cell, 6];
                    var numDocto = ((object[,])cells)[cell, 7];
                    var fornecedor = ((object[,])cells)[cell, 8];
                    var quantEntradas = ((object[,])cells)[cell, 9];
                    var entraUnit = ((object[,])cells)[cell, 10];
                    var entradas = ((object[,])cells)[cell, 11];
                    var quantSaidas = ((object[,])cells)[cell, 12];
                    var saidaUnit = ((object[,])cells)[cell, 13];
                    var saidas = ((object[,])cells)[cell, 14];
                    var serieNFE = ((object[,])cells)[cell, 15];
                    var naturezaOperacao = ((object[,])cells)[cell, 16];
                    var contaEstoque = ((object[,])cells)[cell, 17];
                    var contaContrapartida = ((object[,])cells)[cell, 18];

                    TipoMovimentoContabil tipoMovimento = new TipoMovimentoContabil();

                    tipoMovimento.Linhas = Linhas;

                    if (reduzido != null)
                    {
                        tipoMovimento.Reduzido = reduzido.ToString();
                    }
                    else
                    {
                        tipoMovimento.Reduzido = null;
                    }
                    if (codigoIndustrial != null)
                    {
                        tipoMovimento.CodigoIndustrial = codigoIndustrial.ToString();
                    }
                    else
                    {
                        tipoMovimento.CodigoIndustrial = null;
                    }
                    if (unidMed != null)
                    {
                        tipoMovimento.UnidMed = unidMed.ToString();
                    }
                    else
                    {
                        tipoMovimento.UnidMed = null;
                    }
                    if (descricao != null)
                    {
                        tipoMovimento.Descricao = descricao.ToString();
                    }
                    else
                    {
                        tipoMovimento.Descricao = null;
                    }
                    if (deposito != null)
                    {
                        tipoMovimento.Deposito = deposito.ToString();
                    }
                    else
                    {
                        tipoMovimento.Deposito = null;
                    }
                    if (tipoMov != null)
                    {
                        tipoMovimento.TipoMov = tipoMov.ToString();
                    }
                    else
                    {
                        tipoMovimento.TipoMov = null;
                    }
                    if (data != null)
                    {
                        tipoMovimento.Data = data.ToString();
                    }
                    else
                    {
                        tipoMovimento.Data = null;
                    }
                    if (numDocto != null)
                    {
                        tipoMovimento.NumDocto = numDocto.ToString();
                    }
                    else
                    {
                        tipoMovimento.NumDocto = null;
                    }
                    if (fornecedor != null)
                    {
                        tipoMovimento.Fornecedor = fornecedor.ToString();
                    }
                    else
                    {
                        tipoMovimento.Fornecedor = null;
                    }
                    if (quantEntradas != null)
                    {
                        tipoMovimento.QuantEntradas = quantEntradas.ToString();
                    }
                    else
                    {
                        tipoMovimento.QuantEntradas = null;
                    }
                    if (entraUnit != null)
                    {
                        tipoMovimento.EntraUnit = entraUnit.ToString();
                    }
                    else
                    {
                        tipoMovimento.EntraUnit = null;
                    }
                    if (entradas != null)
                    {
                        tipoMovimento.Entradas = entradas.ToString();
                    }
                    else
                    {
                        tipoMovimento.Entradas = null;
                    }
                    if (quantSaidas != null)
                    {
                        tipoMovimento.QuantSaidas = quantSaidas.ToString();
                    }
                    else
                    {
                        tipoMovimento.QuantSaidas = null;
                    }
                    if (saidaUnit != null)
                    {
                        tipoMovimento.SaidaUnit = saidaUnit.ToString();
                    }
                    else
                    {
                        tipoMovimento.SaidaUnit = null;
                    }
                    if (saidas != null)
                    {
                        tipoMovimento.Saidas = saidas.ToString();
                    }
                    else
                    {
                        tipoMovimento.Saidas = null;
                    }

                    if (serieNFE != null)
                    {
                        tipoMovimento.SerieNFE = serieNFE.ToString();
                    }
                    else
                    {
                        tipoMovimento.SerieNFE = null;
                    }

                    if (naturezaOperacao != null)
                    {
                        tipoMovimento.NaturezaOperacao = naturezaOperacao.ToString();
                    }
                    else
                    {
                        tipoMovimento.NaturezaOperacao = null;
                    }

                    if (contaEstoque != null)
                    {
                        tipoMovimento.ContaEstoque = contaEstoque.ToString();
                    }
                    else
                    {
                        tipoMovimento.ContaEstoque = null;
                    }

                    if (contaContrapartida != null)
                    {
                        tipoMovimento.ContaContrapartida = contaContrapartida.ToString();
                    }
                    else
                    {
                        tipoMovimento.ContaContrapartida = null;
                    }
                    formulario.Add(tipoMovimento);
                }

                return formulario;
            }
        }
        public List<TipoMovimentoContabil> MovimentoContabilAlterado(TipoMovimentoContabil tipo, object caminho1)
        {
            List<TipoMovimentoContabil> movimentos = new List<TipoMovimentoContabil>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var packager = new ExcelPackage(new FileInfo((string)caminho1));

            var workbook = packager.Workbook;
            var sheet = workbook.Worksheets[0];
            object cells = sheet.Cells["A2:AA1500"].Value;

            if (sheet.Cells["A1"].Value == null || sheet.Cells["B1"].Value == null || sheet.Cells["C1"].Value == null || sheet.Cells["D1"].Value == null ||
               sheet.Cells["E1"].Value == null || sheet.Cells["F1"].Value == null || sheet.Cells["G1"].Value == null || sheet.Cells["H1"].Value == null ||
               sheet.Cells["I1"].Value == null || sheet.Cells["J1"].Value == null || sheet.Cells["K1"].Value == null || sheet.Cells["L1"].Value == null ||
               sheet.Cells["M1"].Value == null || sheet.Cells["N1"].Value == null || sheet.Cells["O1"].Value == null || sheet.Cells["P1"].Value == null ||
               sheet.Cells["Q1"].Value == null || sheet.Cells["R1"].Value == null || sheet.Cells["S1"].Value == null || sheet.Cells["T1"].Value != null)

            {
                movimentos = null;
                return movimentos;
            }
            else
            {
                for (int cell = 0; cell < 1499; cell++)
                {
                    var reduzido = ((object[,])cells)[cell, 0];
                    var deposito = ((object[,])cells)[cell, 4];
                    var tipoMov = ((object[,])cells)[cell, 5];
                    var numDocto = ((object[,])cells)[cell, 7];
                    if (reduzido != null && deposito != null && numDocto != null && tipoMov != null)
                    {

                        conexaoOra = new OracleConnection(conexaoOracle);
                        Command = conexaoOra.CreateCommand();
                        Command.CommandText = @"select * from movimentacao m 
                    where m.deposito= " + deposito + " and m.documento_sintetico = " + numDocto + " and m.reduzido_item=" + reduzido + " and m.num_tipo_movimento = " + tipo.DETipoMov + " and ";
                        conexaoOra.Open();
                        OracleDataReader dr = Command.ExecuteReader();
                        dr.Read();

                        if (dr.HasRows)
                        {
                            Command.CommandText = @" update movimentacao m set m.num_tipo_movimento = " + tipo.PARATipoMov + "" +
                       " where m.deposito= " + deposito + " and m.documento_sintetico= " + numDocto + "  and m.reduzido_item=" + reduzido + " and m.num_tipo_movimento = " + tipo.DETipoMov + "";
                            Command.ExecuteNonQuery();
                        }
                        TipoMovimentoContabil tipoMovimento = new TipoMovimentoContabil();
                        movimentos.Add(tipoMovimento);
                        dr.Close();
                        conexaoOra.Close();


                    }


                }
            }
            return movimentos;
        }
        public List<ContabilPasso1> Contabils(ContabilPasso1 contabil)
        {
            List<ContabilPasso1> contabils = new List<ContabilPasso1>();
            conexaoOra = new OracleConnection(conexaoOracle);
            Command = conexaoOra.CreateCommand();


            string dataInicial = Convert.ToString(contabil.DataInicial);
            var inicial = dataInicial.Replace("00:00:00", "").Trim().Split("/");
            var anoInicial = inicial[2];
            var mesInicial = inicial[1];
            var diaInicial = inicial[0];
            var Inicialformatada = anoInicial + mesInicial + diaInicial;

            string dataFinal = Convert.ToString(contabil.DataFinal);
            var final = dataFinal.Replace("00:00:00", "").Trim().Split("/");
            var anoFinal = final[2];
            var mesfinal = final[1];
            var diafinal = final[0];
            var Finalformatada = anoFinal + mesfinal + diafinal;

            Command.Parameters.Clear();
            Command.CommandText += @"select m.NUMERO_MOVIMENTO," + Constants.vbCrLf;
            Command.CommandText += "m.SITUACAO," + Constants.vbCrLf;
            Command.CommandText += "m.REDUZIDO_ITEM," + Constants.vbCrLf;
            Command.CommandText += "m.OPERACAO," + Constants.vbCrLf;
            Command.CommandText += "m.NUM_TIPO_MOVIMENTO," + Constants.vbCrLf;
            Command.CommandText += "m.NUM_DOCUMENTO," + Constants.vbCrLf;
            Command.CommandText += "m.DATA_DOCUMENTO," + Constants.vbCrLf;
            Command.CommandText += "m.PRECO," + Constants.vbCrLf;
            Command.CommandText += "m.QUANTIDADE_REAL," + Constants.vbCrLf;
            Command.CommandText += "m.SECAO," + Constants.vbCrLf;
            Command.CommandText += "m.DEPOSITO," + Constants.vbCrLf;
            Command.CommandText += "m.CONTA_CONTABIL," + Constants.vbCrLf;
            Command.CommandText += "m.CONTA_CONTABIL_SAIDA," + Constants.vbCrLf;
            Command.CommandText += "m.IDPESSOASFJ" + Constants.vbCrLf;
            Command.CommandText += "from vw_movi_entr_said_valor m" + Constants.vbCrLf;
            Command.CommandText += " where m.DATA_DOCUMENTO between " + Inicialformatada + " and " + Finalformatada + "" + Constants.vbCrLf;
            Command.CommandText += "and m.DEPOSITO in (" + contabil.Deposito + ")" + Constants.vbCrLf;
            Command.CommandText += "and m.NUM_TIPO_MOVIMENTO = " + contabil.NumTipoMovimento + "" + Constants.vbCrLf;

            if (contabil.Situacao != null)
            {
                Command.CommandText += "and m.SITUACAO = " + contabil.Situacao + " " + Constants.vbCrLf;

            }
            if (contabil.Operacao != null)
            {
                Command.CommandText += "and m.OPERACAO <> " + contabil.Operacao + "" + Constants.vbCrLf;
            }

            conexaoOra.Open();
            OracleDataReader dr = Command.ExecuteReader();
            while (dr.Read())
            {
                ContabilPasso1 passo1 = new ContabilPasso1();
                passo1.NumeroMovimento = dr["NUMERO_MOVIMENTO"].ToString();
                passo1.Situacao = dr["SITUACAO"].ToString();
                passo1.ReduzidoItem = dr["REDUZIDO_ITEM"].ToString();
                passo1.Operacao = dr["OPERACAO"].ToString();
                passo1.NumTipoMovimento = dr["NUM_TIPO_MOVIMENTO"].ToString();
                passo1.NumDocumento = dr["NUM_DOCUMENTO"].ToString();
                passo1.DataDocumento = dr["DATA_DOCUMENTO"].ToString();
                passo1.Preco = dr["PRECO"].ToString();
                passo1.QuantidadeReal = dr["QUANTIDADE_REAL"].ToString();
                passo1.Secao = dr["SECAO"].ToString();
                passo1.Deposito = dr["DEPOSITO"].ToString();
                passo1.ContaContabil = dr["CONTA_CONTABIL"].ToString();
                passo1.ContaContabilSaida = dr["CONTA_CONTABIL_SAIDA"].ToString();
                passo1.IdPessoaSFJ = dr["IDPESSOASFJ"].ToString();
                contabils.Add(passo1);

            }
            dr.Close();
            conexaoOra.Close();
            return contabils;

        }
        public List<Tblmovimentacao> Contabils2(Tblmovimentacao tblmovimentacao, object movimentos)
        {
            List<Tblmovimentacao> contabils = new List<Tblmovimentacao>();
            conexaoOra = new OracleConnection(conexaoOracle);
            Command = conexaoOra.CreateCommand();
            object mov = movimentos;

            string dataInicial = Convert.ToString(tblmovimentacao.DATA_INICIAL);
            var inicial = dataInicial.Replace("00:00:00", "").Trim().Split("/");
            var anoInicial = inicial[2];
            var mesInicial = inicial[1];
            var diaInicial = inicial[0];
            var Inicialformatada = anoInicial + mesInicial + diaInicial;

            string dataFinal = Convert.ToString(tblmovimentacao.DATA_FINAL);
            var final = dataFinal.Replace("00:00:00", "").Trim().Split("/");
            var anoFinal = final[2];
            var mesfinal = final[1];
            var diafinal = final[0];
            var Finalformatada = anoFinal + mesfinal + diafinal;
            
           
            Command.Parameters.Clear();

            Command.CommandText += @"SELECT m.* " + Constants.vbCrLf;
            Command.CommandText += " from movimentacao m" + Constants.vbCrLf;
            Command.CommandText += "where 1 = 1" + Constants.vbCrLf;
            Command.CommandText += "and m.cdfilial = " + tblmovimentacao.CDFILIAL + "" + Constants.vbCrLf;
            Command.CommandText += "and data_documento between " + Inicialformatada + " and " + Finalformatada + "" + Constants.vbCrLf;
            if (movimentos != null)

            {                            
                //tentar FAZER DO JEITO CERTO COM DAPER
                //Command.CommandText += "and m.numero_movimento in ("+ String.Join(",", elementos) + ") " + Constants.vbCrLf;
                Command.CommandText += "and m.numero_movimento in (" + movimentos + ") " + Constants.vbCrLf;
            }
            Command.CommandText += "and m.DEPOSITO in (" + tblmovimentacao.DEPOSITO + ") " + Constants.vbCrLf;

            if (tblmovimentacao.REDUZIDO_ITEM != 0)
            {
                Command.CommandText += "and m.REDUZIDO_ITEM in (" + tblmovimentacao.REDUZIDO_ITEM + ") " + Constants.vbCrLf;
            }
            if (tblmovimentacao.NUM_DOCUMENTO != 0)
            {
                Command.CommandText += "and m.NUM_DOCUMENTO in (" + tblmovimentacao.NUM_DOCUMENTO + ") " + Constants.vbCrLf;
            }
            if (tblmovimentacao.NUM_TIPO_MOVIMENTO != 0)
            {
                Command.CommandText += "and m.NUM_TIPO_MOVIMENTO in (" + tblmovimentacao.NUM_TIPO_MOVIMENTO + ") " + Constants.vbCrLf;
            }
            if (tblmovimentacao.SITUACAO != 0)
            {
                Command.CommandText += " and m.situacao = " + tblmovimentacao.SITUACAO + " " + Constants.vbCrLf;
            }
            if (tblmovimentacao.OPERACAO != 0)
            {
                Command.CommandText += "and m.operacao = " + tblmovimentacao.OPERACAO + " " + Constants.vbCrLf;
            }

            conexaoOra.Open();
            //DataTable table = new DataTable();
            // Command.ExecuteNonQuery();
           
            OracleDataReader dr = Command.ExecuteReader();

            //table.Load(dr);
            while (dr.Read())
            {
                Tblmovimentacao tblmovimentacao1 = new Tblmovimentacao();
                tblmovimentacao1.NUMERO_MOVIMENTO = Convert.ToInt32(dr["NUMERO_MOVIMENTO"]);
                tblmovimentacao1.SITUACAO = Convert.ToInt32(dr["SITUACAO"]);
                tblmovimentacao1.REDUZIDO_ITEM = Convert.ToInt32(dr["REDUZIDO_ITEM"]);
                tblmovimentacao1.OPERACAO = Convert.ToInt32(dr["OPERACAO"]);
                tblmovimentacao1.NUM_TIPO_MOVIMENTO = Convert.ToInt32(dr["NUM_TIPO_MOVIMENTO"]);
                tblmovimentacao1.MOV_CONFIRMADO = Convert.ToInt32(dr["MOV_CONFIRMADO"]);
                tblmovimentacao1.NUM_DOCUMENTO = Convert.ToInt32(dr["NUM_DOCUMENTO"]);
                tblmovimentacao1.DATA_DOCUMENTO = Convert.ToInt32(dr["DATA_DOCUMENTO"]);
                tblmovimentacao1.DATA_CONF_DOCUMENTO = Convert.ToInt32(dr["DATA_CONF_DOCUMENTO"]);
                tblmovimentacao1.CONCENTRACAOREALINSU = Convert.ToInt32(dr["CONCENTRACAOREALINSU"]);
                tblmovimentacao1.LOTE = dr["LOTE"].ToString();
                tblmovimentacao1.QUALIDADE = Convert.ToInt32(dr["QUALIDADE"]);
                tblmovimentacao1.PRECO_SEM_TAXA = (float)Convert.ToDouble(dr["PRECO_SEM_TAXA"]);
                tblmovimentacao1.PRECO = (float)Convert.ToDouble(dr["PRECO"]);
                tblmovimentacao1.PRECO_MEDIO = (float)Convert.ToDouble(dr["PRECO_MEDIO"]);
                tblmovimentacao1.QUANTIDADE_REAL = (float)Convert.ToDouble(dr["QUANTIDADE_REAL"]);
                tblmovimentacao1.QUANTIDADE_PREVISTA = (float)Convert.ToDouble(dr["QUANTIDADE_PREVISTA"]);
                tblmovimentacao1.VOLUMES = Convert.ToInt32(dr["VOLUMES"]);
                tblmovimentacao1.DESTINO = Convert.ToInt32(dr["DESTINO"]);
                tblmovimentacao1.SECAO = dr["SECAO"].ToString();
                tblmovimentacao1.DEPOSITO = Convert.ToInt32(dr["DEPOSITO"]);
                tblmovimentacao1.CONTA_CONTABIL = dr["CONTA_CONTABIL"].ToString();
                tblmovimentacao1.CONTA_CONTABIL_SAIDA = dr["CONTA_CONTABIL_SAIDA"].ToString();
                tblmovimentacao1.GERA_ESTOQUE_BLOQUEA = dr["GERA_ESTOQUE_BLOQUEA"].ToString();
                tblmovimentacao1.CONTA_NUMERADA = Convert.ToInt32(dr["CONTA_NUMERADA"]);
                tblmovimentacao1.TIPO_REPROCESSO = Convert.ToInt32(dr["TIPO_REPROCESSO"]);
                tblmovimentacao1.QUANTIDADE_BALANCO = Convert.ToInt32(dr["QUANTIDADE_BALANCO"]);
                tblmovimentacao1.DOCUMENTO_SINTETICO = Convert.ToInt32(dr["DOCUMENTO_SINTETICO"]);
                tblmovimentacao1.DEPOSITO_CUSTO = Convert.ToInt32(dr["DEPOSITO_CUSTO"]);
                tblmovimentacao1.VALOR_MOVIMENTO_CUST = Convert.ToInt32(dr["VALOR_MOVIMENTO_CUST"]);
                tblmovimentacao1.IDPESSOASFJ = Convert.ToInt32(dr["IDPESSOASFJ"]);
                tblmovimentacao1.IDPROCEDENCIA = Convert.ToInt32(dr["IDPROCEDENCIA"]);
                tblmovimentacao1.CDFILIAL = Convert.ToInt32(dr["CDFILIAL"]);
                tblmovimentacao1.TPMOVIMENTO = Convert.ToInt32(dr["TPMOVIMENTO"]);
                tblmovimentacao1.HRMOVIMENTO = Convert.ToInt32(dr["HRMOVIMENTO"]);
                tblmovimentacao1.TISTATUSINTEGRACAO = Convert.ToInt32(dr["TISTATUSINTEGRACAO"]);
                tblmovimentacao1.CODREDUSUARIO = Convert.ToInt32(dr["CODREDUSUARIO"]);
                contabils.Add(tblmovimentacao1);
            }

            Command.ExecuteNonQuery();
            dr.Close();
            conexaoOra.Close();
            return contabils;
        }
        // desativado
        public int PerfilPesq(int perfil)
        {

            conexaoOra = new OracleConnection(conexaoOracle);
            Command = conexaoOra.CreateCommand();
            Command.CommandText = @"select PERFIL from tbl_pessoa where idpessoa = " + perfil + "";
            conexaoOra.Open();
            OracleDataReader dr = Command.ExecuteReader();
            dr.Read();

            if (dr.HasRows)
            {
                perfil = Convert.ToInt32(dr["PERFIL"]);
            }
            conexaoOra.Close();
            return perfil;
        }
    }

}
