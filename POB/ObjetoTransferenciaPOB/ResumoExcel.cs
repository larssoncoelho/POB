using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace POB.ObjetoTransferenciaPOB
{
    public class ResumoExcel
    {
        internal int IdMaterial;

        public double tocRepeticoesDeTorres { get; set; }
        public double tocRepeticoesDoPavimento { get; set; }
        public  double tocRepeticoesNoPavimento { get; set; }

        public string Pavimento { get; internal set; }
        public string Ambiente { get; internal set; }
        public string Folha { get; internal set; }
        public string NumFolhaReferencia { get; internal set; }
        public double AreaVertical { get; internal set; }
        public double AreaHorizontal { get; internal set; }
        public string CategoriaDoMaterial { get; internal set; }
        public string CodigoImper { get; internal set; }
        public string NomeSistema { get; internal set; }
        public double AlturaDaParade { get; internal set; }
        public int IdMontagem { get; internal set; }
        public int IdElemento { get; internal set; }
        public string TipoDeElemento { get; internal set; }
        public double ComprimentoJunta { get; internal set; }
        public string NomeDoMaterial { get; internal set; }
        public double AreaTotal { get; internal set; }
    }
}
