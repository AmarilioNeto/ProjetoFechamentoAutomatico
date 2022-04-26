using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.VisualBasic.CompilerServices;


namespace ProjetoFechamentoAutomatico.Models.Pessoa
{
    public class Pessoa
    {
        public int IDPESSOA { get; set; }
        public string NOME { get; set; }
        public string USUARIO { get; set; }
        public string SETOR { get; set; }
        public string CARGO { get; set; }
        public string SENHA { get; set; }
        public string EMAIL { get; set; }
        public string ACESSO { get; set; }
        public string CONFSENHA { get; set; }
        public string PERFIL { get; set; }
    }
}
