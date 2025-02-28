using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Reflection;
using wf = System.Windows.Forms;
using cl = System.Windows.Clipboard;
using System.Windows.Media.Imaging;
using Funcoes;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Plumbing;

using Autodesk.Revit.Attributes;
using System.Data;
using RevitCreation = Autodesk.Revit.Creation;
using System.Runtime.InteropServices;
using Autodesk.Revit;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using System.Windows.Media.Animation;

using dck =  Revit.SDK.Samples.DockableDialogs.CS;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.UI.Events;
using POB.Updater;
namespace POB
{

    public class CriarMenu : IExternalApplication
    {

        private static string diretorio = "";
        public static List<string> ListaDeCategoriasDeItensContaveis = new List<string>();
        public static List<string> ListaCategoriaItensSistemaPorMetro = new List<string>();
        public static OverrideGraphicSettings OrgProdutoInexistente;
        // public static List<string> ListaQueTemTipoDeSistema = new List<string>();
        public static List<POB.ObjetoTransferenciaPOB.DadosChapa> ListaDeChapas = new List<POB.ObjetoTransferenciaPOB.DadosChapa>();
        public static string Diretorio
        {
            get
            {
                return Assembly.GetExecutingAssembly().Location.Substring(
                                             0, Assembly.GetExecutingAssembly().Location.Length - 7);
            }
            set { diretorio = value; }
        }

        internal static OverrideGraphicSettings getOrgProdutoInexistente(Document uidoc)
        {
            OverrideGraphicSettings org = new OverrideGraphicSettings();
            Color c = new Color(System.Drawing.Color.Red.R, System.Drawing.Color.Red.G,
               System.Drawing.Color.Red.R);
            org.SetProjectionLineColor(c);
           
            org.SetCutLineColor(c);
#if DEBUG20201
#else

         /*   org.SetCutFillColor(c);
            org.SetProjectionFillPatternId(Util.GetSolidFill(uidoc));
            org.SetProjectionFillPatternVisible(true);
            org.SetProjectionFillColor(c);*/
#endif
            org.SetHalftone(false);
            org.SetSurfaceTransparency(0);
            return org;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public static void GetQueTemTipoDeSistema(Document uiDoc)
        {
            //CriarMenu.ListaQueTemTipoDeSistema.Add(Category.GetCategory(uiDoc, BuiltInCategory.OST_PipeFitting).Name);
            //CriarMenu.ListaQueTemTipoDeSistema.Add(Category.GetCategory(uiDoc, BuiltInCategory.OST_PlumbingFixtures).Name);
        }
        public static void GetCategoriaItensSistemaPorMetro(Document uiDoc)
        {
            CriarMenu.ListaCategoriaItensSistemaPorMetro.Add(Category.GetCategory(uiDoc, BuiltInCategory.OST_PipeCurves).Name);

            CriarMenu.ListaCategoriaItensSistemaPorMetro.Add(Category.GetCategory(uiDoc, BuiltInCategory.OST_Conduit).Name);
            CriarMenu.ListaCategoriaItensSistemaPorMetro.Add(Category.GetCategory(uiDoc, BuiltInCategory.OST_DuctCurves).Name);
            CriarMenu.ListaCategoriaItensSistemaPorMetro.Add(Category.GetCategory(uiDoc, BuiltInCategory.OST_FlexPipeCurves).Name);
            CriarMenu.ListaCategoriaItensSistemaPorMetro.Add(Category.GetCategory(uiDoc, BuiltInCategory.OST_FlexDuctCurves).Name);
            CriarMenu.ListaCategoriaItensSistemaPorMetro.Add(Category.GetCategory(uiDoc, BuiltInCategory.OST_CableTray).Name);

        }
            
        public static void GetCategoriasItensContaveisPorQtde(Document iuDoc)
        {
         ListaDeCategoriasDeItensContaveis.Add(Category.GetCategory(iuDoc, BuiltInCategory.OST_ElectricalFixtures).Name);
            ListaDeCategoriasDeItensContaveis.Add(Category.GetCategory(iuDoc, BuiltInCategory.OST_LightingFixtures).Name);
            ListaDeCategoriasDeItensContaveis.Add(Category.GetCategory(iuDoc, BuiltInCategory.OST_PlumbingFixtures).Name);
            ListaDeCategoriasDeItensContaveis.Add(Category.GetCategory(iuDoc, BuiltInCategory.OST_ElectricalEquipment).Name);
            ListaDeCategoriasDeItensContaveis.Add(Category.GetCategory(iuDoc, BuiltInCategory.OST_MechanicalEquipment).Name);
            ListaDeCategoriasDeItensContaveis.Add(Category.GetCategory(iuDoc, BuiltInCategory.OST_SpecialityEquipment).Name);
            ListaDeCategoriasDeItensContaveis.Add(Category.GetCategory(iuDoc, BuiltInCategory.OST_DataDevices).Name);
            ListaDeCategoriasDeItensContaveis.Add(Category.GetCategory(iuDoc, BuiltInCategory.OST_LightingDevices).Name);
            ListaDeCategoriasDeItensContaveis.Add(Category.GetCategory(iuDoc, BuiltInCategory.OST_FireAlarmDevices).Name);
            ListaDeCategoriasDeItensContaveis.Add(Category.GetCategory(iuDoc, BuiltInCategory.OST_NurseCallDevices).Name);
            ListaDeCategoriasDeItensContaveis.Add(Category.GetCategory(iuDoc, BuiltInCategory.OST_SecurityDevices).Name);
            ListaDeCategoriasDeItensContaveis.Add(Category.GetCategory(iuDoc, BuiltInCategory.OST_TelephoneDevices).Name);
            ListaDeCategoriasDeItensContaveis.Add(Category.GetCategory(iuDoc, BuiltInCategory.OST_CableTrayFitting).Name);
            ListaDeCategoriasDeItensContaveis.Add(Category.GetCategory(iuDoc, BuiltInCategory.OST_ConduitFitting).Name);
            ListaDeCategoriasDeItensContaveis.Add(Category.GetCategory(iuDoc, BuiltInCategory.OST_GenericModel).Name);
            ListaDeCategoriasDeItensContaveis.Add(Category.GetCategory(iuDoc, BuiltInCategory.OST_PipeAccessory).Name);
            ListaDeCategoriasDeItensContaveis.Add(Category.GetCategory(iuDoc, BuiltInCategory.OST_DuctFitting).Name);
     
            ListaDeCategoriasDeItensContaveis.Add(Category.GetCategory(iuDoc, BuiltInCategory.OST_PipeFitting).Name);
           
        }

        internal static CategorySet ObterCategoriasDeInstalacoesRevit(Document iuDoc)
        {
            CategorySet categorySet = new CategorySet();
            foreach (var item in CriarMenu.ListaCategorias())
            {
                categorySet.Insert(Category.GetCategory(iuDoc, item));
            }

            categorySet.Insert(Category.GetCategory(iuDoc, BuiltInCategory.OST_CableTray));
            categorySet.Insert(Category.GetCategory(iuDoc, BuiltInCategory.OST_FlexDuctCurves));
           categorySet.Insert(Category.GetCategory(iuDoc, BuiltInCategory.OST_Conduit));
            categorySet.Insert(Category.GetCategory(iuDoc, BuiltInCategory.OST_PipeCurves));
            categorySet.Insert(Category.GetCategory(iuDoc, BuiltInCategory.OST_FlexPipeCurves));
            categorySet.Insert(Category.GetCategory(iuDoc, BuiltInCategory.OST_DuctCurves));
            categorySet.Insert(Category.GetCategory(iuDoc, BuiltInCategory.OST_Walls));
            categorySet.Insert(Category.GetCategory(iuDoc, BuiltInCategory.OST_Floors));
            categorySet.Insert(Category.GetCategory(iuDoc, BuiltInCategory.OST_Ceilings));
            categorySet.Insert(Category.GetCategory(iuDoc, BuiltInCategory.OST_Doors));
            categorySet.Insert(Category.GetCategory(iuDoc, BuiltInCategory.OST_Windows));
            categorySet.Insert(Category.GetCategory(iuDoc, BuiltInCategory.OST_GenericModel));


            categorySet.Insert(Category.GetCategory(iuDoc, BuiltInCategory.OST_Conduit));
            return categorySet;
        }

        public static LogicalOrFilter DefinirFiltroRegisterUpdateNivelAndSpace()
        {

            IList<ElementFilter> b = new List<ElementFilter>();
            b.Add(new ElementClassFilter(typeof(Level)));
            b.Add(new ElementClassFilter(typeof(SpatialElement)));
       

            return new LogicalOrFilter(b);

        }

        public static List<BuiltInCategory> ListaCategorias()
        {
        return    new List<BuiltInCategory>() { 
            //    BuiltInCategory[] bics = new BuiltInCategory[] {
                BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_CableTrayFitting,
                BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_ConduitFitting,
                BuiltInCategory.OST_DuctCurves,
                BuiltInCategory.OST_DuctFitting,
                BuiltInCategory.OST_DuctTerminal,
                BuiltInCategory.OST_ElectricalEquipment,
                BuiltInCategory.OST_ElectricalFixtures,
                BuiltInCategory.OST_LightingDevices,
                BuiltInCategory.OST_LightingFixtures,
                BuiltInCategory.OST_CommunicationDevices,
                BuiltInCategory.OST_FireAlarmDevices,
                BuiltInCategory.OST_DataDevices,
                BuiltInCategory.OST_SecurityDevices,
                BuiltInCategory.OST_MechanicalEquipment,
                BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_PipeAccessory,
                BuiltInCategory.OST_PlaceHolderPipes,
                BuiltInCategory.OST_PlumbingFixtures,
                BuiltInCategory.OST_SpecialityEquipment,
                BuiltInCategory.OST_Sprinklers,
                //BuiltInCategory.OST_Wire,
              };

        }
        public static LogicalOrFilter DefinirFiltroRegisterUpdateDescricaoEelemento()
        {

            List<BuiltInCategory> bics = ListaCategorias();
            IList<ElementFilter> a = new List<ElementFilter>(bics.Count());

            foreach (BuiltInCategory bic in bics)
            {
                a.Add(new ElementCategoryFilter(bic));
            }
            LogicalOrFilter categoryFilter = new LogicalOrFilter(a);


            LogicalAndFilter familyInstanceFilter = 
                new LogicalAndFilter(categoryFilter, new ElementClassFilter(typeof(FamilyInstance)));



            IList<ElementFilter> b = new List<ElementFilter>();
            b.Add(new ElementClassFilter(typeof(Autodesk.Revit.DB.Electrical.CableTray)));
            b.Add(new ElementClassFilter(typeof(Autodesk.Revit.DB.Electrical.Conduit)));
            b.Add(new ElementClassFilter(typeof(Duct)));
            b.Add(new ElementClassFilter(typeof(Pipe)));
            b.Add(new ElementClassFilter(typeof(FlexDuct)));
            b.Add(new ElementClassFilter(typeof(FlexPipe)));

            b.Add(familyInstanceFilter);
            b.Add(new ElementClassFilter(typeof(SpatialElement)));
            b.Add(new ElementClassFilter(typeof(Level)));


            return new LogicalOrFilter(b);

        }
        void OnViewActivated( object sender, ViewActivatedEventArgs e)
        {
            View vPrevious = e.PreviousActiveView;
            View vCurrent = e.CurrentActiveView;
            if (vCurrent is ViewSheet)
                TaskDialog.Show("_", "é uma folha");
         
        }
        public Result OnStartup(UIControlledApplication application)
        {

            string dir = Diretorio;

             RevitUpdaterAmbienteElemento revitUpdater =   new RevitUpdaterAmbienteElemento(application.ActiveAddInId);
             UpdaterRegistry.RegisterUpdater(revitUpdater);
             ElementCategoryFilter f = new ElementCategoryFilter(Autodesk.Revit.DB.BuiltInCategory.OST_PipeFitting);


            UpdaterRegistry.AddTrigger(revitUpdater.GetUpdaterId(), DefinirFiltroRegisterUpdateDescricaoEelemento(), Element.GetChangeTypeAny());
            UpdaterRegistry.AddTrigger(revitUpdater.GetUpdaterId(), DefinirFiltroRegisterUpdateDescricaoEelemento(), Element.GetChangeTypeElementAddition());

            

            BitmapImage icone_nivel24x24px = new BitmapImage(new Uri(@dir + "icone_nivel24x24px.png"));

            BitmapImage icone_nivel16x16px = new BitmapImage(new Uri(@dir + "icone_nivel16x16px.png"));
      
            BitmapImage icone_imperminacao16x16px = new BitmapImage(new Uri(@dir + "icone_imperminação16x16px.png"));

            BitmapImage icone_imperminacaao24x24px = new BitmapImage(new Uri(@dir + "icone_imperminação24x24px.png"));


          


            /*RevitUpdaterEspaco revitUpdater1 = new RevitUpdaterEspaco(application.ActiveAddInId);
            UpdaterRegistry.RegisterUpdater(revitUpdater1);
            */
            //ElementCategoryFilter f1 = new ElementCategoryFilter(Autodesk.Revit.DB.BuiltInCategory.OST_PipeFitting);


            //  UpdaterRegistry.AddTrigger(revitUpdater1.GetUpdaterId(), DefinirFiltroRegisterUpdateNivelAndSpace(), Element.GetChangeTypeAny());
            //   UpdaterRegistry.AddTrigger(revitUpdater1.GetUpdaterId(), DefinirFiltroRegisterUpdateNivelAndSpace(), Element.GetChangeTypeAny());


            //UpdaterRegistry.AddTrigger(revitUpdater.GetUpdaterId(), DefinirFiltroRegisterUpdateDescricaoEelemento(), Element.GetChangeTypeElementAddition());

            //definir 
            // application.ViewActivated += new EventHandler<ViewActivatedEventArgs>(OnViewActivated);
            //voltar

            /* amb = new ArvoreMedicaoBloco(dir);
             amb.Dir = dir;
             amb.inputar(ref aa, ref f);*/
            /*Uri uripilar = new Uri(@dir + "Pilar.jpg");
            BitmapImage pilar = new BitmapImage(uripilar);

            Uri uriViga= new Uri(@dir + "Viga.png");
            BitmapImage viga = new BitmapImage(uriViga);

            Uri uriLaje = new Uri(@dir + "Viga.png");
            BitmapImage laje = new BitmapImage(uriLaje);
            */
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
            try
            {
                application.CreateRibbonTab(Properties.Settings.Default.nomePainelFerramentaModelagem);
            }
            catch
            {
            }

            /*application.CreateRibbonTab( dck.Globals.DiagnosticsTabName);
            RibbonPanel panel = application.CreateRibbonPanel(dck.Globals.DiagnosticsTabName, dck.Globals.DiagnosticsPanelName);

            panel.AddSeparator();

            PushButtonData pushButtonRegisterPageData = new PushButtonData(dck.Globals.RegisterPage, dck.Globals.RegisterPage,
            dck.FileUtility.GetAssemblyFullName(), typeof(dck.ExternalCommandRegisterPage).FullName);
            pushButtonRegisterPageData.LargeImage = new BitmapImage(new Uri(dck.FileUtility.GetApplicationResourcesPath() + "Register.png"));
            PushButton pushButtonRegisterPage = panel.AddItem(pushButtonRegisterPageData) as PushButton;
            pushButtonRegisterPage.AvailabilityClassName = typeof(dck.ExternalCommandRegisterPage).FullName;


            PushButtonData pushButtonShowPageData = new PushButtonData(dck.Globals.ShowPage, dck.Globals.ShowPage, dck.FileUtility.GetAssemblyFullName(), typeof(dck.ExternalCommandShowPage).FullName);
            pushButtonShowPageData.LargeImage = new BitmapImage(new Uri(dck.FileUtility.GetApplicationResourcesPath() + "Show.png"));
            PushButton pushButtonShowPage = panel.AddItem(pushButtonShowPageData) as PushButton;
            pushButtonShowPage.AvailabilityClassName = typeof(dck.ExternalCommandShowPage).FullName;


            PushButtonData pushButtonHidePageData = new PushButtonData(dck.Globals.HidePage, dck.Globals.HidePage, dck.FileUtility.GetAssemblyFullName(), typeof(dck.ExternalCommandHidePage).FullName);
            pushButtonHidePageData.LargeImage = new BitmapImage(new Uri(dck.FileUtility.GetApplicationResourcesPath() + "Hide.png"));
            PushButton pushButtonHidePage = panel.AddItem(pushButtonHidePageData) as PushButton;
            pushButtonHidePage.AvailabilityClassName = typeof(dck.ExternalCommandHidePage).FullName;
            */
            RibbonPanel painelEstrutura;
          
         

            painelEstrutura = application.CreateRibbonPanel(Properties.Settings.Default.nomePainelFerramentaModelagem,"Impermeabilização");


            string path = Assembly.GetExecutingAssembly().Location;
            


            /*PushButtonData criarTabicaDaLajeData = new PushButtonData("cmdCriarTabica", "Criar tabica", thisAssemblyPath, "POB.CriarTabica1");
            PushButton criarTabicaDaLaje = painelEstrutura.AddItem(criarTabicaDaLajeData) as PushButton;
            //criarTabicaDaLaje.LargeImage = largeImage;*/
            painelEstrutura.AddSeparator();

            /*            PushButtonData lookaheadElementoData = new PushButtonData("cmdLookahead", "Lookahead", thisAssemblyPath, "criarMenu.Lookahead");
                        PushButton lookaheadElemento = painelAvanco.AddItem(lookaheadElementoData) as PushButton;
                        lookaheadElemento.LargeImage = largeMeta;
                        painelAvanco.AddSeparator();
                        */


            /*PushButtonData criarPilarDoIFCData = new PushButtonData("cmdAvancar", "Pilar \n IFC", thisAssemblyPath, "POB.CriarPilarDoIFC" );
            PushButton criarPilarDoIFC = painelEstrutura.AddItem(criarPilarDoIFCData) as PushButton;
            //criarPilarDoIFC.LargeImage = largeImage;
            painelEstrutura.AddSeparator();
           
            PushButtonData criarVigaDoIFCData = new PushButtonData("cmdVigaIFC", "Viga \n IFC", thisAssemblyPath, "POB.CriarVigaDoIFC");
            PushButton criarVigaDoIFC = painelEstrutura.AddItem(criarVigaDoIFCData) as PushButton;
            //criarVigaDoIFC.LargeImage = largeImage;
            painelEstrutura.AddSeparator();

            PushButtonData criarLajeDIFCData = new PushButtonData("cmdPisoIFC", "Laje \n IFC", thisAssemblyPath, "POB.CriarPisoDoIfc");
            PushButton criarLajeDIFC = painelEstrutura.AddItem(criarLajeDIFCData) as PushButton;
            //criarVigaDoIFC.LargeImage = largeImage;
         
            PushButtonData criarTabicaDaLajeData = new PushButtonData("cmdCriarTabica", "Criar tabica", thisAssemblyPath, "POB.CriarTabica1");
            PushButton criarTabicaDaLaje = painelEstrutura.AddItem(criarTabicaDaLajeData) as PushButton;
            //criarTabicaDaLaje.LargeImage = largeImage;
            painelEstrutura.AddSeparator();           

          

            PushButtonData herdarData = new PushButtonData("cmdHerdar", "Herdar", thisAssemblyPath, "POB.HerdarValoresParametros");
            PushButton herdar = painelEstrutura.AddItem(herdarData) as PushButton;
            //criarTabicaDaLaje.LargeImage = largeImage;
            painelEstrutura.AddSeparator();

            /*
            PushButtonData pcpInformacoesData = new PushButtonData("cmdInfo", "INFO", thisAssemblyPath, "POB.InformacaoElemento");
            PushButton pcpInformacoes = painelEstrutura.AddItem(pcpInformacoesData) as PushButton;
            //pcpElemento.LargeImage = largePCP;
            painelEstrutura.AddSeparator();

            PushButtonData pcpGeometriaPecaData = new PushButtonData("cmdGeometriaPeca", "INFO", thisAssemblyPath, "POB.GeometriaPeca");
            PushButton GeometriaPeca = painelEstrutura.AddItem(pcpGeometriaPecaData) as PushButton;
            //pcpElemento.LargeImage = largePCP; 
            
            painelEstrutura.AddSeparator();

            */
            /* PushButtonData revestiPilarData = new PushButtonData("cmdRevestirPilar", "Revestir \n Pilar", thisAssemblyPath, "POB.RevestiPilar");
             PushButton revestiPilar = painelEstrutura.AddItem(revestiPilarData) as PushButton;
             painelEstrutura.AddSeparator();





             PushButtonData criarFormaData = new PushButtonData("cmdcriarForma", "Forma", thisAssemblyPath, "POB.CriarForma");
             PushButton criarForma = painelEstrutura.AddItem(criarFormaData) as PushButton;


             painelEstrutura.AddSeparator();


             PushButtonData medidaJanelaData = new PushButtonData("cmdmedidaJanela", "Medida \n janela", thisAssemblyPath, "POB.MedidaJanela");
             PushButton medidaJanela = painelEstrutura.AddItem(medidaJanelaData) as PushButton;*/


             PushButtonData ObterAreaSupeficieBaseData = new PushButtonData("cmdObterAreaSupeficieBase", "Área \n base", thisAssemblyPath, "POB.ObterAreaSupeficieBase");
             PushButton ObterAreaSupeficieBaseJanela = painelEstrutura.AddItem(ObterAreaSupeficieBaseData) as PushButton;
           
             PushButtonData ObterAreaSupeficieLateralData = new PushButtonData("cmdObterAreaSupeficieLateral", 
                                           "Área \n lateral", 
                                           thisAssemblyPath, 
                                           "POB.ObterAreaLateral");
             PushButton ObterAreaSupeficieLateral = painelEstrutura.AddItem(ObterAreaSupeficieLateralData) as PushButton;

             PushButtonData ObterDiametroData = new PushButtonData("cmdDiâmetro", "Obter \n diâmetro", thisAssemblyPath, "POB.ObterDiametro");
             PushButton ObterDiametro = painelEstrutura.AddItem(ObterDiametroData) as PushButton;


           /* PushButtonData GetNivelExtraidoData = new PushButtonData("cmdGetNivelExtraidoData", "Nível \n extraído", thisAssemblyPath, "POB.GetNivelExtraido");
               PushButton GetNivelExtraido = painelEstrutura.AddItem(GetNivelExtraidoData) as PushButton;
              */ 
            PushButtonData extrairData = new PushButtonData("cmdextrair", "Extrair", thisAssemblyPath, "POB.ExtrairDados");
            PushButton extrair = painelEstrutura.AddItem(extrairData) as PushButton;
            //criarTabicaDaLaje.LargeImage = largeImage;
            painelEstrutura.AddSeparator();

            PushButtonData obterData = new PushButtonData("cmdObter", "Obter volume", thisAssemblyPath, "POB.ObterVolume");
            PushButton obter = painelEstrutura.AddItem(obterData) as PushButton;

            PushButtonData obterDataMaiorFace = new PushButtonData("cmdObterMaiorFace", "Obter maior face", thisAssemblyPath, "POB.ObterMaiorFace");
            PushButton obterMaiorFace = painelEstrutura.AddItem(obterDataMaiorFace) as PushButton;

            PushButtonData pisoParaForroData = new PushButtonData("cmdpisoParaForroData", "Piso para forro", thisAssemblyPath, "POB.CriarForroAPartirdoPiso");
            PushButton pisoParaForro = painelEstrutura.AddItem(pisoParaForroData) as PushButton;

            PushButtonData obterDataMaiorLado = new PushButtonData("cmdObterMaiorLado", "Obter maior lado", thisAssemblyPath, "POB.ObterMaiorComprimentoDoSolido");
            PushButton obterMaiorLado = painelEstrutura.AddItem(obterDataMaiorLado) as PushButton;

            PushButtonData dadosVista3DData = new PushButtonData("cmdDadosVista3D", "Dados 3D", thisAssemblyPath, "POB.DadosVista3D");
            PushButton dadosVista3D = painelEstrutura.AddItem(dadosVista3DData) as PushButton;

            PushButtonData dadosVista3DImagemData = new PushButtonData("cmdDadosVista3DImagem", "Dados 3D Imagem", thisAssemblyPath, "POB.DadosVista3DImagem");
            PushButton dadosVista3DImagem = painelEstrutura.AddItem(dadosVista3DImagemData) as PushButton;

            PushButtonData dadosDadosEixo = new PushButtonData("cmdDadosEixo", "DadosEixo", thisAssemblyPath, "POB.DadosEixo");
            PushButton DadosEixo = painelEstrutura.AddItem(dadosDadosEixo) as PushButton;

            PushButtonData dadosDadosVista = new PushButtonData("cmdDadosVista", "DadosVista", thisAssemblyPath, "POB.DadosVista");
            PushButton DadosVista = painelEstrutura.AddItem(dadosDadosVista) as PushButton;

            PushButtonData CriarTemplateData = new PushButtonData("cmdCriarTemplate", "CriarTemplate", thisAssemblyPath, "POB.CriarTemplate");
            PushButton CriarTemplate = painelEstrutura.AddItem(CriarTemplateData) as PushButton;

            painelEstrutura.AddSeparator();
            PushButtonData DadosImperData = new PushButtonData("cmdDadosImper", "Dados \n imper", thisAssemblyPath, "POB.DadosImper");
            PushButton DadosImper1 = painelEstrutura.AddItem(DadosImperData) as PushButton;
            DadosImper1.LargeImage = icone_imperminacaao24x24px;
            DadosImper1.Image = icone_imperminacao16x16px;

            PushButtonData GerarInclinacaoPisoData = new PushButtonData("cmdinclinacao", "Inclinação", thisAssemblyPath, "POB.GerarInclinacaoNoPiso");
            PushButton GerarInclinacaoPiso = painelEstrutura.AddItem(GerarInclinacaoPisoData) as PushButton;
            GerarInclinacaoPiso.LargeImage = icone_nivel24x24px;
            GerarInclinacaoPiso.Image = icone_nivel16x16px;
            /*PushButtonData renumeraMontagemData = new PushButtonData("cmRenumerar", "Renumerar", thisAssemblyPath, "POB.RenumeraItens");
            PushButton renumeraMontagem  = painelEstrutura.AddItem(renumeraMontagemData) as PushButton;
            /*PushButtonData ObterAmbienteData = new PushButtonData("cmdObterAmbiente", "Obter ambiente", thisAssemblyPath, "POB.GetAmbiente");
            PushButton ObterAmbiente = painelEstrutura.AddItem(ObterAmbienteData) as PushButton;
            */
            /*      PushButtonData BuscarProplemasData = new PushButtonData("cmdBuscarProplemas", " Buscar", thisAssemblyPath, "POB.BuscarProblemas");
                  PushButton BuscarProplemas = painelEstrutura.AddItem(BuscarProplemasData) as PushButton;
            */
            PushButtonData ExpProplemasData = new PushButtonData("ObterId", "ObterId", 
                thisAssemblyPath, "POB.ObterId");
            PushButton ExpProplemas = painelEstrutura.AddItem(ExpProplemasData) as PushButton;
            
            CriarMenu.ListaDeChapas.Add(new ObjetoTransferenciaPOB.DadosChapa
            {
                Bitola= "#26",
                DensidadeArea =  3.662,
                MenorLado = 0,
                MaiorLado = 300 / 1000 / 0.3048

            });
            CriarMenu.ListaDeChapas.Add(new ObjetoTransferenciaPOB.DadosChapa
            {
                Bitola = "#24",
                DensidadeArea = 4.882,
                MenorLado =310.01 / 1000 / 0.3048,
                MaiorLado = 750 / 1000 / 0.3048

            });
            CriarMenu.ListaDeChapas.Add(new ObjetoTransferenciaPOB.DadosChapa
            {
                Bitola = "#22",
                DensidadeArea = 6.103,
                MenorLado = 750.01/1000/0.3048,
                MaiorLado = 1200 / 1000 / 0.3048

            });
            CriarMenu.ListaDeChapas.Add(new ObjetoTransferenciaPOB.DadosChapa
            {
                Bitola = "#20",
                DensidadeArea = 7.324,
                MenorLado = 1200.01 / 1000 /0.3048,
                MaiorLado = 1500 / 1000 / 0.3048

            });
            return Result.Succeeded;

        }
    }
}