using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.ObjectModel;
using ObjetoDeTranferencia;

namespace POB.DockableDialogs
{
    public  class ColecaoServico
    {
        internal bool abrindo;
       /* public ObservableCollection<ObjetoDeTranferencia.ServicoEtapa> DataSouce = new ObservableCollection<ObjetoDeTranferencia.ServicoEtapa>();
        //private DadosIntegracao _dadosIntegracao;
        private ICollection<ObjetoDeTranferencia.Servico> servicos;
        private List<ObjetoDeTranferencia.ETAPA> etapas;
        private IEnumerable<ServicoEtapa> servicoEtapa;
        private List<ObjetoDeTranferencia.ServicoEtapa> servicoEtapa_sem_filtro;
       */
        public ColecaoServico(DadosIntegracao dadosIntegracao)
        {
           // _dadosIntegracao = dadosIntegracao;
        }

        public bool CachedUpdates { get; internal set; }

        public void Open()
        {
           /* servicos = new Negocios.ACESSO_SERVICO().SelecionaPorDicionario(_dadosIntegracao).Values;
            etapas = new Negocios.ACESSO_ETAPA().SelecionaEtapaUtiizada(_dadosIntegracao);
            var etapaParaServico = from ETAPA e in etapas
                                   select new ServicoEtapa
                                   {
                                       SERVICO_ID = e.ETAPA_ID ?? -1,
                                       ITEM = e.ITEM,
                                       SERVICO = e.ETAPA_DESCRICAO,
                                       TIPO = 0,
                                       COMPLEMENTO = "-1"
                                   };
            servicoEtapa = from Servico servico in servicos
                           join ETAPA etapa in etapas on servico.ETAPA_ID equals etapa.ETAPA_ID into servicoEtapa1
                           from se in servicoEtapa1.DefaultIfEmpty()
                           select new ServicoEtapa
                           {
                               SERVICO_ID = servico.ServicoId,
                               SERVICO = servico.DescServico,
                               ELEMENTO = servico.Elemento,
                               UNID = servico.Unid,
                               COMPLEMENTO = servico.Complemento,
                               ITEM = se?.ITEM ?? "1000.1000",
                               POSICAO = servico.POSICAO,
                               ETAPA_ID = servico.ETAPA_ID,
                               ETAPA = se?.ETAPA_DESCRICAO ?? String.Empty,
                               TIPO = 1
                           };
            servicoEtapa = servicoEtapa.Union(etapaParaServico);
            DataSouce = new ObservableCollection<ServicoEtapa>(servicoEtapa.OrderBy(x => x.ITEM).ThenBy(x => x.TIPO).ThenBy(x => x.POSICAO));
            servicoEtapa_sem_filtro = servicoEtapa.ToList();*/
        }

    }
}
