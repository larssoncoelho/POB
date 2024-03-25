using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Reflection;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Data;
using RevitCreation = Autodesk.Revit.Creation;
using System.Runtime.InteropServices;
using Autodesk.Revit;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;


namespace POB
{
    

    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]

    public class SubstituirMaterial : IExternalCommand
    {

        FamilySymbol fs1;
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit,
            ref string message, ElementSet elements)
        {
            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;
            Transaction t = new Transaction(uiDoc);

            FilteredElementCollector colecao = new FilteredElementCollector(uiDoc);
            var lista1 = colecao.OfClass(typeof(WallType)).Cast<WallType>().ToList();
            var lista2 = lista1.Where(x => x.FamilyName == "Parede básica").ToList();

            // Funcoes.ProgressoFuncao progressoFuncao = new ProgressoFuncao();
            //    progressoFuncao.pgc.Maximum = lista2.Count();
            //    progressoFuncao.pgc.Step = 1;
            //    progressoFuncao.Show();
            //   progressoFuncao.TopMost = true;
            Apresentacao.Perguntar perguntar = new Apresentacao.Perguntar("Digite o nome do novo material");
            perguntar.ShowDialog();
            if(perguntar.Continuar)
            {
                return Result.Cancelled;
            }
            Material material = colecao.OfClass(typeof(Material)).
                Where(x => x.Name == perguntar.Texto).
                Cast<Material>().FirstOrDefault();
            perguntar.Texto = "Digite o nome do material antigo";
            perguntar.ShowDialog();
            if (perguntar.Continuar)
            {
                return Result.Cancelled;
            }
            string materialAntigo = perguntar.Texto;
            foreach (WallType w in lista2)
            {
                t.Start("teste");
                try
                {
                    CompoundStructure composicao = w.GetCompoundStructure();
                    foreach (var camada in composicao.GetLayers())
                    {
                        if((Math.Round(camada.Width*0.3048*100,0)==5))
                            if((uiDoc.GetElement( camada.MaterialId) as Material).Name.Contains(materialAntigo))
                            {
                                camada.MaterialId = material.Id;
                            }
                    }                 
                    w.SetCompoundStructure(composicao);

                    t.Commit();

                }
                catch
                {
                    t.RollBack();
                }

            }



            return Result.Succeeded;
        }
    }




}
