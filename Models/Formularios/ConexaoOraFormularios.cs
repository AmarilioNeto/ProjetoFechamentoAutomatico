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
using System.Runtime;

namespace ProjetoFechamentoAutomatico.Models.Formularios
{
    public class ConexaoOraFormularios
    {
        public string conexaoOracle = @"DATA SOURCE=(DESCRIPTION=(ADDRESS_LIST=
(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.1.29)(PORT=1521)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=tubdhmlg)));User Id=teste;Password=teste;";

        public OracleConnection conexaoOra;
        public OracleCommand Command;

        public List<DDCVI> Excel(string caminho1)
        {
            
            List<DDCVI> formulario = new List<DDCVI>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // mudar para deixar o caminho da pasta dinamico.
            var packager = new ExcelPackage(new FileInfo(caminho1));

            var workbook = packager.Workbook;
            var sheet = workbook.Worksheets["DD"];
            object cells = sheet.Cells["A3:O184"].Value;
            for (int cell = 0; cell < 181; cell++)
            {
                var red = ((object[,])cells)[cell, 0];
                var codigo = ((object[,])cells)[cell, 1];
                var unid = ((object[,])cells)[cell, 2];
                var descricao = ((object[,])cells)[cell, 3];
                var dep = ((object[,])cells)[cell, 4];
                var mov = ((object[,]) cells)[cell, 5];
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
                if(red != null)
                {
                    dDCVI.Red = red.ToString();
                }
                else
                {
                    dDCVI.Red = null;
                }
                if(codigo != null)
                {
                    dDCVI.Codigo = codigo.ToString();
                }
                else
                {
                    dDCVI.Codigo = null;
                }
                if(unid != null)
                {
                    dDCVI.Unid = unid.ToString();
                }
                else
                {
                    dDCVI.Unid = null;
                }
               if(descricao != null)
                {
                    dDCVI.Descricao = descricao.ToString();
                }
               else
                {
                    dDCVI.Descricao = null;
                }
                if(dep != null)
                {
                    dDCVI.Dep = dep.ToString();
                }
                else
                {
                    dDCVI.Dep = null;
                }
               if(mov != null)
                {
                    dDCVI.Mov = mov.ToString();
                }
               else
                {
                    dDCVI.Mov = null;
                }
                if(data != null)
                {
                    dDCVI.Data = data.ToString();
                }
                else
                {
                    dDCVI.Data = null;
                }
                if(docto != null)
                {
                    dDCVI.Docto = docto.ToString();
                }
                else
                {
                    dDCVI.Docto = null;
                }
                if(forne != null)
                {
                    dDCVI.Forne = forne.ToString();
                }
                else
                {
                    dDCVI.Forne = null;
                }
                if(quantE != null)
                {
                    dDCVI.QuantE = quantE.ToString();
                }
                else
                {
                    dDCVI.QuantE = null;
                }
                if(unit != null)
                {
                    dDCVI.Unit = unit.ToString();
                }
                else
                {
                    dDCVI.Unit = null;
                }
                if(entradas != null)
                {
                    dDCVI.Entradas = entradas.ToString();
                }
               else
                {
                    dDCVI.Entradas = null;
                }
              if(quantS != null)
                {
                    dDCVI.QuantS = quantS.ToString();
                }
               else
                {
                    dDCVI.QuantS = null;
                }
                if(unitS != null)
                {
                    dDCVI.UnitS = unitS.ToString();
                }
                else
                {
                    dDCVI.UnitS = null;
                }
                if(saidas != null)
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

        public List<Formularios> ExcelAlterado(string caminho1)
        {
            List<Formularios> formularios = new List<Formularios>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // mudar para deixar o caminho da pasta dinamico.
            var packager = new ExcelPackage(new FileInfo(caminho1));

            var workbook = packager.Workbook;
            var sheet = workbook.Worksheets["DD"];
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
                    if(red != null )
                    {
                        formularios1.Red = red.ToString();
                    }
                    else
                    {
                        formularios1.Red = null;
                    }
                    if(codigo != null)
                    {
                        formularios1.Codigo = codigo.ToString();
                    }
                    else
                    {
                        formularios1.Codigo = null;
                    }
                    if(unid != null)
                    {
                        formularios1.Unid = unid.ToString();
                    }
                    else
                    {
                        formularios1.Unid = null;
                    }
                    if(descricao != null)
                    {
                        formularios1.Descricao = descricao.ToString();
                    }
                    else
                    {
                        formularios1.Descricao = null;
                    }
                    if(dep != null)
                    {
                        formularios1.Dep = dep.ToString();
                    }
                    else
                    {
                        formularios1.Dep = null;
                    }
                    if(mov != null)
                    {
                        formularios1.Mov = mov.ToString();
                    }
                    else
                    {
                        formularios1.Mov = null;
                    }
                    if(data != null)
                    {
                        formularios1.Data = data.ToString();
                    }
                    else
                    {
                        formularios1.Data = null;
                    }
                    if(docto != null)
                    {
                        formularios1.Docto = docto.ToString();
                    }
                    else
                    {
                        formularios1.Docto = null;
                    }
                    if(forne != null)
                    {
                        formularios1.Forne = forne.ToString();
                    }
                    else
                    {
                        formularios1.Forne = null;
                    }
                    if(quantE != null)
                    {
                        formularios1.QuantE = quantE.ToString();
                    }
                    else
                    {
                        formularios1.QuantE = null;
                    }
                    if(unit != null)
                    {
                        formularios1.Unit = unit.ToString();
                    }
                    else
                    {
                        formularios1.Unit = null;
                    }
                    if(entradas != null)
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

        public List<Formularios>AtualConfirmado(string caminho1)
        {
            List<Formularios> formularios = new List<Formularios>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // mudar para deixar o caminho da pasta dinamico.
            var packager = new ExcelPackage(new FileInfo(caminho1));

            var workbook = packager.Workbook;
            var sheet = workbook.Worksheets["DD"];
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
