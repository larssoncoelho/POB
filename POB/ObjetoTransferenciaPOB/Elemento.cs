using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace POB.ObjetoTransferenciaPOB
{
    public class ElementoPOB
    {
        public string Item { get; internal set; }
        public string Descricao { get; internal set; }
        public string Unid { get; internal set; }
        public double Qtde { get; internal set; }
        public string TipoDeSistema { get; internal set; }
        public string ClassificacaoSistema { get; internal set; }
    }
}
