using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc.Rendering;
using Oracle.ManagedDataAccess.Client;
using ProjetoFechamentoAutomatico.Models;
using System.Data.OracleClient;
using System.IO;


namespace ProjetoFechamentoAutomatico.Models.Pessoa
{

    class ConexaoOracle
    {
        public string conexaoOracle = @"DATA SOURCE=(DESCRIPTION=(ADDRESS_LIST=
(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.1.29)(PORT=1521)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=tubdhmlg)));User Id=teste;Password=teste;";

        public OracleConnection conexaoOra;
        public OracleCommand Command;
        public OracleTransaction Transaction;
        
        

        public Pessoa Conectar(Pessoa pessoa)
        {
            //Pessoa pessoa1 = new Pessoa();
            conexaoOra = new OracleConnection(conexaoOracle);            
            Command = conexaoOra.CreateCommand();          
           Command.CommandText = @"select IdPessoa, acesso from tbl_pessoa where usuario = '" + pessoa.USUARIO + "'and senha = '"+pessoa.SENHA + "'";           
           conexaoOra.Open();         
            OracleDataReader dr = Command.ExecuteReader();
            dr.Read();
            
           if(dr.HasRows)
          {
                
                pessoa.IDPESSOA = Convert.ToInt32(dr["IDPESSOA"]);
                pessoa.ACESSO = dr["ACESSO"].ToString();
             
           }       
           else
            {
                pessoa.IDPESSOA = 0;
            }
            dr.Close();
            conexaoOra.Close();            
            return pessoa;                          
        }
      
        public Pessoa inserirUsuario(Pessoa pessoa)
        {
            
            conexaoOra = new OracleConnection(conexaoOracle);
            Command = conexaoOra.CreateCommand();          
            Command.CommandText = @"select usuario from tbl_pessoa where usuario='" + pessoa.USUARIO + "'";
            conexaoOra.Open();
            Transaction = conexaoOra.BeginTransaction();
            OracleDataReader dr = Command.ExecuteReader();
            dr.Read();
            if(dr.HasRows != true)
            {
                Command.CommandText = @"insert into tbl_pessoa
(idpessoa, nome, usuario,setor,cargo,senha,email,acesso,perfil) 
values(sq_usuarios_cad.nextval,'" + pessoa.NOME + "','" + pessoa.USUARIO + "','" + pessoa.SETOR + "','" + pessoa.CARGO + "','" + pessoa.SENHA + "','" + pessoa.EMAIL + "','NAO',"+pessoa.PERFIL+")";
                Command.ExecuteNonQuery();
                Transaction.Commit();
                dr.Close();
                conexaoOra.Close();

                Command.CommandText = @"select MAX(IDPESSOA) from tbl_pessoa";
                conexaoOra.Open();
               
                OracleDataReader dr1 = Command.ExecuteReader();
                dr1.Read();
                if (dr1.HasRows)
                {                  
                     var pessoa1 = Convert.ToInt32(dr1["MAX(IDPESSOA)"]);
                    pessoa.IDPESSOA = pessoa1;
                }
            }
            dr.Close();      
            conexaoOra.Close();
            return pessoa;
        }
        public Pessoa updateSenha(Pessoa pessoa)
        {
            conexaoOra = new OracleConnection(conexaoOracle);
            Command = conexaoOra.CreateCommand();            
            Command.CommandText = @"update tbl_pessoa set senha = '" + pessoa.SENHA + "',acesso = 'SIM' WHERE IDPESSOA = " + pessoa.IDPESSOA + "";
            conexaoOra.Open();
            Command.ExecuteNonQuery();          
            conexaoOra.Close();
            return pessoa;
        }
       public List<EsqueciSenha> EsqueciSenha(EsqueciSenha senha)
        {
            List<EsqueciSenha> senhas = new List<EsqueciSenha>();
            conexaoOra = new OracleConnection(conexaoOracle);
            Command = conexaoOra.CreateCommand();
            Command.CommandText = @"select senha from tbl_pessoa where usuario ='" + senha.USUARIO + "' and email = '" + senha.EMAIL + "'";
            conexaoOra.Open();
            OracleDataReader dr = Command.ExecuteReader();
            dr.Read();
            if(dr.HasRows)
            {
                EsqueciSenha esqueci = new EsqueciSenha();
                esqueci.SENHA = dr["SENHA"].ToString();               
                senhas.Add(esqueci);                
            }
            
            dr.Close();
            conexaoOra.Close();
            return senhas;
        }
        public List<Pessoa> Emails()
        {
            List<Pessoa> emails = new List<Pessoa>();
            conexaoOra = new OracleConnection(conexaoOracle);
            Command = conexaoOra.CreateCommand();
            Command.CommandText = @"select EMAIL, IDPESSOA from tbl_pessoa where PERFIL= 1";
            conexaoOra.Open();
            OracleDataReader dr = Command.ExecuteReader();
           while(dr.Read())
            {
                Pessoa email = new Pessoa();
                email.EMAIL = dr["EMAIL"].ToString();
                email.IDPESSOA = Convert.ToInt32(dr["IDPESSOA"]);
                emails.Add(email);
            }
            dr.Close();
            conexaoOra.Close();
            return emails;
        }
        public List<Pessoa> Email(string emails)
        {
            List<Pessoa> pessoa = new List<Pessoa>();
            conexaoOra = new OracleConnection(conexaoOracle);
            Command = conexaoOra.CreateCommand();
            Command.CommandText = @"select EMAIL, NOME, SETOR, CARGO from tbl_pessoa where IDPESSOA = " + emails + "";
            conexaoOra.Open();
            OracleDataReader dr = Command.ExecuteReader();
            dr.Read();
               if(dr.HasRows)
            {
                Pessoa email = new Pessoa();
                 email.EMAIL = dr["EMAIL"].ToString();
                email.NOME = dr["NOME"].ToString();
                email.SETOR = dr["SETOR"].ToString();
                email.CARGO = dr["CARGO"].ToString();
                pessoa.Add(email);
            }
            dr.Close();
            conexaoOra.Close();
            return pessoa;
        }
        
  
    }
}
