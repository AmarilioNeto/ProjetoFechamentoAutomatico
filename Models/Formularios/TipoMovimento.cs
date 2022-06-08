using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ProjetoFechamentoAutomatico.Models.Formularios
{
    public class TipoMovimento
    {
        public string Reduzido { get; set; }
        public string CodigoIndustrial { get; set; }
        public string UnidMed { get; set; }
        public string Descricao { get; set; }
        public string Deposito { get; set; }
        public string TipoMov { get; set; }
        public string Data { get; set; }
        public string NumDocto { get; set; }
        public string Fornecedor { get; set; }
        public string QuantEntradas { get; set; }
        public string EntraUnit { get; set; }
        public string Entradas { get; set; }
        public string QuantSaidas { get; set; }
        public string SaidaUnit { get; set; }
        public string Saidas { get; set; }
        public string SerieNFE { get; set; }
        public string NaturezaOperacao { get; set; }
        public string DestacaICMS { get; set; }
        public string DestacaPIS { get; set; }
        public string DestacaCofins { get; set; }
        public string DestacaIPI { get; set; }
        public string DestacaFrete { get; set; }
        public string ICMSFrete { get; set; }
        public string PISFrete { get; set; }
        public string CofinsFrete { get; set; }
        public string ContaEstoque { get; set; }
        public string ContaContrapartida { get; set; }   
        public string DE { get; set; }
        public string PARA { get; set; }
    }
}
