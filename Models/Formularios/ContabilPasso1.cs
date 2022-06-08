using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ProjetoFechamentoAutomatico.Models.Formularios
{
    public class ContabilPasso1
    {
        public string NumeroMovimento { get; set; }
        public string Situacao { get; set; }
        public string ReduzidoItem { get; set; }
        public string Operacao { get; set; }
        public string NumTipoMovimento { get; set; }
        public string NumDocumento { get; set; }
        public string DataDocumento { get; set; }
        public string Preco { get; set; }
        public string QuantidadeReal { get; set; }
        public string Secao { get; set; }
        public string Deposito { get; set; }
        public string ContaContabil { get; set; }
        public string ContaContabilSaida { get; set; }
        public string IdPessoaSFJ { get; set; }       
        public DateTime DataInicial { get; set; }
        public DateTime DataFinal { get; set; }
       
    }
}
