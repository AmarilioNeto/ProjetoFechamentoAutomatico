using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ProjetoFechamentoAutomatico.Models.Formularios
{
    public class ContabilPasso2
    {
        public string cdFilial { get; set; }
        public DateTime dataInicial { get; set; }
        public DateTime dataFinal { get; set; }
        public string deposito { get; set; }
        public string situacao { get; set; }
        public string operacao { get; set; }
        public string contaContabil { get; set; }
        public string numeroMovimento { get; set; }
        public string tipoMovimento { get; set; }
        public string contaContabilSaida { get; set; }
        public string reduzidoItem { get; set; }
        public string numeroDocumento { get; set; }
    }
}
