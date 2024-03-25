using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI.Selection;
using revitDB = Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using appRevit = Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;

using Funcoes;
using db = Autodesk.Revit.DB;
using Autodesk.Revit.DB;
using System.Collections.ObjectModel;
using ObjetoDeTranferencia;
using System.Data.SqlClient;

namespace POB.DockableDialogs
{
    /// <summary>
    /// Interaction logic for MainPage.xaml
    /// </summary>
    public partial class MainPage : Page
    {
        private DadosIntegracao _dadosIntegracao;
       // public ColecaoServico bsServico;
        //public FbsqlConnection sqlConnection = new FbsqlConnection();

        public Boolean continuar = false;
       // public Plan_servico_amoNegocio manipulacao;
        public string dir;
        public string UAU_COMP_SELECIONADA = "";
        public List<db.Element> ListaDeElemento = new List<db.Element>();
        private revitDB.Document _uidoc;
        private bool filtrando;

        public MainPage()
        {
            InitializeComponent();
        }
        public void Abrir(string idir, List<db.Element> lista, DadosIntegracao dadosIntegracao, revitDB.Document iuDoc)
        {

          /*  dir = idir;
            _uidoc = iuDoc;
            manipulacao = new Plan_servico_amoNegocio(idir);
            ListaDeElemento = lista;
            if (dadosIntegracao != null)
                _dadosIntegracao = dadosIntegracao;
            var obra = new ACESSO_OBRA().Listar(_dadosIntegracao.Diretorio, (int)_dadosIntegracao.ObraId)[0];
            _dadosIntegracao.LoginUAU = obra.USUARIOAPI;
            _dadosIntegracao.Senha = obra.SENHAAPI;
            _dadosIntegracao.UrlApi = obra.LINKAPI;
            _dadosIntegracao.Token = obra.TOKENAUTORIZACAO;
            bsServico = new ColecaoServico(_dadosIntegracao);
            bsServico.Open();
            /*DefaultCellStyle1 = grdServico.DefaultCellStyle;
            DefaultCellStyle1.BackColor = System.Drawing.Color.SkyBlue;

            DefaultCellStylePadrao = grdServico.DefaultCellStyle;
            DefaultCellStylePadrao.BackColor = System.Drawing.Color.White;
            */
            using (var d = Dispatcher.DisableProcessing())
            {
              /*  bsServico.CachedUpdates = false;
                bsServico.abrindo = true;
                sqlConnection.Diretorio = dir;
                sqlConnection.Ativo(true);
                grdServico.ItemsSource = bsServico.DataSouce;
                /*FuncaoApresentacao.TrataColunas("Complemento;Descrição;Parâmetro;Posicao;Item;Etapa;Unid;Id;Id;", ";", grdServico);
                FuncaoApresentacao.TrataColunasLargura("80;250;80;60;80;150;100;150;80;", ";", grdServico);
                grdServico.Columns["ITEM"].ReadOnly = true;
                grdServico.Columns["ETAPA_ID"].ReadOnly = true;
                grdServico.Columns["ETAPA"].ReadOnly = true;
                bsServico.abrindo = false;*/
            }

        }
        private void DockableDialogs_Loaded(object sender, RoutedEventArgs e)
        {

        }

        private void btnVerificar_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btnSubstituir_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btnAtribuirCodigo_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btnAtribuirItem_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
