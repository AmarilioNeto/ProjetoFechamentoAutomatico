using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ProjetoFechamentoAutomatico.Models.Formularios
{
    public class Tblmovimentacao
    {
        public int NUMERO_MOVIMENTO { get; set; }
        public int SITUACAO { get; set; }
        public int REDUZIDO_ITEM { get; set; }
        public int OPERACAO { get; set; }
        public int NUM_TIPO_MOVIMENTO { get; set; }
        public int MOV_CONFIRMADO { get; set; }
        public int NUM_DOCUMENTO { get; set; }
        public int DATA_DOCUMENTO { get; set; }
        public int DATA_CONF_DOCUMENTO { get; set; }
        public int CONCENTRACAOREALINSU { get; set; }
        public string LOTE { get; set; }
        public int QUALIDADE { get; set; }
        public float  PRECO_SEM_TAXA { get; set; }
        public float PRECO { get; set; }
        public float  PRECO_MEDIO { get; set; }
        public float QUANTIDADE_REAL { get; set; }
        public float QUANTIDADE_PREVISTA { get; set; }
        public int VOLUMES { get; set; }
        public int DESTINO { get; set; }
        public string SECAO { get; set; }
        public int DEPOSITO { get; set; }
        public string CONTA_CONTABIL { get; set; }
        public string CONTA_CONTABIL_SAIDA { get; set; }
        public string GERA_ESTOQUE_BLOQUEA { get; set; }
        public int CONTA_NUMERADA { get; set; }
        public int TIPO_REPROCESSO { get; set; }
        public float QUANTIDADE_BALANCO { get; set; }
        public int DOCUMENTO_SINTETICO { get; set; }
        public int DEPOSITO_CUSTO { get; set; }
        public float VALOR_MOVIMENTO_CUST { get; set; }
        public int IDPESSOASFJ { get; set; }
        public int IDPROCEDENCIA { get; set; }
        public int CDFILIAL { get; set; }
        public int TPMOVIMENTO { get; set; }
        public int HRMOVIMENTO { get; set; }
        public int TISTATUSINTEGRACAO { get; set; }
        public int CODREDUSUARIO { get; set; }
        public DateTime DATA_INICIAL { get; set; }
        public DateTime DATA_FINAL { get; set; }
    }
}
