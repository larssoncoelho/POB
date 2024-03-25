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
using Funcoes;



namespace POB
{

    
    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]
    public class Renomear : IExternalCommand
    {
  /*    public string nomeFamily = "";

     
        public static void BindSharedParam(

          Document doc,
          Category cat,
          string paramName,
          string grpName,
          ParameterType paramType,
          bool visible,
          bool instanceBinding)
        {
            try // generic
            {
                Autodesk.Revit.ApplicationServices.Application app = doc.Application;
                // This is needed already here to 
                // store old ones for re-inserting

                CategorySet catSet = app.Create.NewCategorySet();

                // Loop all Binding Definitions
                // IMPORTANT NOTE: Categories.Size is ALWAYS 1 !?
                // For multiple categories, there is really one 
                // pair per each category, even though the 
                // Definitions are the same...

                DefinitionBindingMapIterator iter = doc.ParameterBindings.ForwardIterator();

                while (iter.MoveNext())
                {
                    Definition def = iter.Key;
                    ElementBinding elemBind = (ElementBinding)iter.Current;

                    // Got param name match

                    if (paramName.Equals(def.Name, StringComparison.CurrentCultureIgnoreCase))
                    {
                        // Check for category match - Size is always 1!

                        if (elemBind.Categories.Contains(cat))
                        {
    
                            if (instanceBinding)
                            {
                            }
                            else
                            {
                            }

                        }

    
                        else
                        {
                            foreach (Category catOld
                              in elemBind.Categories)
                                catSet.Insert(catOld); // 1 only, but no index...
                        }
                    }
                }

    
                DefinitionFile defFile = GetOrCreateSharedParamsFile(app);

                DefinitionGroup defGrp = GetOrCreateSharedParamsGroup(defFile, grpName);

                Definition definition = GetOrCreateSharedParamDefinition(defGrp, paramType, paramName, visible);

                catSet.Insert(cat);

                InstanceBinding bind = null;

                if (instanceBinding)
                {
                    bind = app.Create.NewInstanceBinding(
                      catSet);
                }
                else
                {
                }
    
                if (doc.ParameterBindings.Insert(definition, bind))
                {
                }
                else
                {
                    if (doc.ParameterBindings.ReInsert(definition, bind))
                    {
                    }
                    else
                    {
                    }
                }
            }
            catch (Exception ex)
            { }
        }

        private static Definition GetOrCreateSharedParamDefinition(DefinitionGroup defGrp, ParameterType paramType, string paramName, bool visible)
        {
            throw new NotImplementedException();
        }

        private static DefinitionGroup GetOrCreateSharedParamsGroup(DefinitionFile defFile, string grpName)
        {
            throw new NotImplementedException();
        }

        private static DefinitionFile GetOrCreateSharedParamsFile(Application app)
        {
            throw new NotImplementedException();
        }
        */
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
          /*  UIApplication uiApp = commandData.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;

            Selection sel = uiApp.ActiveUIDocument.Selection;
            Util.uiDoc = uiDoc;
            StringBuilder sb = new StringBuilder();
            ElementClassFilter FamilyFilter = new ElementClassFilter(typeof(Family));
            FilteredElementCollector FamilyInstanceCollector = new FilteredElementCollector(uiDoc);
            ICollection<Element> AllFamily = FamilyInstanceCollector.WherePasses(FamilyFilter).ToElements();
            List<string> ParamLst = new List<string>();
            FamilySymbol FmlySmbl;
            Family Fmly;
            Category cat;
            string paramName;
            string grpName;
            ParameterType paramType;
            bool visible;
            bool instanceBinding;
            //Autodesk.Revit.ApplicationServices.Application app = uiDoc.Application;

            CategorySet catSet = uiDoc.Application.Create.NewCategorySet();
            Categories categories = uiDoc.Settings.Categories;
            foreach (Category categoria in categories)
            {
                if (CsAddPanel.lista.Contains(categoria.Name))
                    catSet.Insert(categoria);
            }
            DefinitionFile dfile = uiDoc.Application.OpenSharedParameterFile();
            InstanceBinding ib = null;
            ib = uiDoc.Application.Create.NewInstanceBinding(catSet);
            if (dfile != null)
            {
                DefinitionGroup grupoAvanco = dfile.Groups.get_Item("Insira aqui o novo nome") as DefinitionGroup;
                foreach (Definition definicao in grupoAvanco.Definitions)
                {
                    uiDoc.ParameterBindings.Insert(definicao, ib, BuiltInParameterGroup.PG_IDENTITY_DATA);
                }
            }

   
            foreach (Family FmlyInst in AllFamily)
            {
                try
                {
                    if (FmlyInst is FamilyInstance)
                    {
                        sb.AppendLine(FmlyInst.Name + "\t" + FmlyInst.Category + "\t" + " família de instância");
                    }
                    else
                    {

                        if (FmlyInst.Document.IsFamilyDocument)
                        {
                            sb.AppendLine(FmlyInst.Name + "\t" + FmlyInst.Category + "\t" + " Família de doc");
                        }
                        else
                        {
                            sb.AppendLine(FmlyInst.Name + "\t" + FmlyInst.Category + "\t" + "outro");
                        }
                    }
                }
                catch (Exception e)
                {
                    sb.AppendLine(FmlyInst.Name + "\t" + FmlyInst.Category + "\t" + e.Message);

                }

            }

         


                    TaskDialog.Show("_", sb.ToString());*/
            return Result.Succeeded;

        }
    }


    public class ItemInformacaoElemento
    {
        public int parameterId { get; set; }

        public bool Origem { get; set; }
        public string Nome { get; set; }
        public string Valor { get; set; }
        public int Id { get; set; }
        public double ValorNatural { get; set; }
    }

    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]
    public class InformacaoElemento : IExternalCommand
    {
        public double d;

        public double diferenca;
        public double diferencAnterior;
        public Level menorLevel;



        internal SketchPlane CreateSketchPlane(Autodesk.Revit.DB.XYZ normal, Autodesk.Revit.DB.XYZ origin,
                                               Document m_familyDocument, UIApplication uiDoc)
        {
            /* // First create a Geometry.Plane which need in NewSketchPlane() method
            Plane geometryPlane = uiDoc.Application.Create.NewPlane(normal, origin);
            if (null == geometryPlane)  // assert the creation is successful
            {
                throw new Exception("Create the geometry plane failed.");
            }
            // Then create a sketch plane using the Geometry.Plane
            SketchPlane plane = SketchPlane.Create(m_familyDocument, geometryPlane);
            // throw exception if creation failed
            if (null == plane)
            {
                throw new Exception("Create the sketch plane failed.");
            }
            return plane;*/
            return null;
        }



        private void CreateExtrusion(UIApplication uiDoc, Document m_familyDocument, RevitCreation.FamilyItemFactory m_creationFamily,
                                     CurveArrArray curveArrArray1, double altura)
        {
            #region Create rectangle profile
#endregion
            // here create rectangular extrusion
            Autodesk.Revit.DB.XYZ normal = Autodesk.Revit.DB.XYZ.BasisZ;
            SketchPlane sketchPlane = CreateSketchPlane(normal, Autodesk.Revit.DB.XYZ.Zero, m_familyDocument, uiDoc);



            Extrusion rectExtrusion = m_creationFamily.NewExtrusion(true, curveArrArray1, sketchPlane, 10);
#if D23 || D24
            var p1 = SpecTypeId.Int.Integer;
            m_familyDocument.FamilyManager.AddParameter("Material", BuiltInParameterGroup.PG_MATERIALS, Autodesk.Revit.DB.Category.GetCategory(m_familyDocument, BuiltInCategory.OST_Materials), true);

#else
                 var p1 =ParameterType.Material;
                    m_familyDocument.FamilyManager.AddParameter("Material", BuiltInParameterGroup.PG_MATERIALS, p1,   true);
         
#endif




            m_familyDocument.FamilyManager.AssociateElementParameterToFamilyParameter(rectExtrusion.LookupParameter("Material"),
                                                                                      m_familyDocument.FamilyManager.get_Parameter(Properties.Settings.Default.Material));


            // move to proper place
            Autodesk.Revit.DB.XYZ transPoint1 = new Autodesk.Revit.DB.XYZ(-16, 0, 0);
            ElementTransformUtils.MoveElement(m_familyDocument, rectExtrusion.Id, transPoint1);

        }

        private string CriarGenericModel(ExternalCommandData revit, CurveArrArray curveArrArray, double altura)
        {
            Document m_familyDocument = revit.Application.ActiveUIDocument.Document;
            // create new family document if active document is not a family document
            if (!m_familyDocument.IsFamilyDocument)
            {
                m_familyDocument = revit.Application.Application.NewFamilyDocument(@"C:\ProgramData\Autodesk\RVT 2015\Family Templates\English\Metric Generic Model.rft");


                if (null == m_familyDocument)
                {


                }
            }

            //m_familyDocument.FamilyManager.RenameCurrentType("teste");
            RevitCreation.FamilyItemFactory m_creationFamily = m_familyDocument.FamilyCreate;
            // create generic model family in the document



            Transaction transaction = new Transaction(m_familyDocument, "CreateGenericModel");
            transaction.Start();


            CreateExtrusion(revit.Application, m_familyDocument, m_creationFamily, curveArrArray, altura);
            //    m_familyDocument.FamilyManager.RenameCurrentType("issso ae");
            //m_creationFamily.
            transaction.Commit();
            //m_familyDocument.Save();
            m_familyDocument.LoadFamily(revit.Application.ActiveUIDocument.Document);
            return m_familyDocument.Title;
        }

        // The main Execute method (inherited from IExternalCommand) must be public
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit,
            ref string message, ElementSet elements)
        {

            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;
            StringBuilder sb = new StringBuilder();

      
            List<ItemInformacaoElemento> lista = new List<ItemInformacaoElemento>();


            Element ele;
            foreach (ElementId eleId in sel.GetElementIds())
            {
                ele = uiDoc.GetElement(eleId);
                
              
                foreach (Parameter param in ele.Parameters)
                {


                    string nome = param.Definition.Name;
                    string tipo = param.StorageType.ToString();
                    switch (param.StorageType)
                    {
                        case StorageType.String:
                            if (param.HasValue)
                                if (!string.IsNullOrEmpty(param.AsString()))
                                    lista.Add(new ItemInformacaoElemento { Nome = nome, Origem = true, Valor = param.AsString(), ValorNatural = 0, Id = ele.Id.IntegerValue });
                            break;


                        case StorageType.Double:
                        case StorageType.Integer:
                            if (param.HasValue)
                                if (!string.IsNullOrEmpty(param.AsValueString()))
                                    lista.Add(new ItemInformacaoElemento { Nome = nome, Origem = true, ValorNatural = param.AsDouble(), Valor = param.AsValueString(), Id = ele.Id.IntegerValue, parameterId = param.Id.IntegerValue });
                            break;
                    }
                    
                }


                if (ele is FamilyInstance)
                {
                    var fs = (ele as FamilyInstance).Symbol;

                    
                    foreach (Parameter param in fs.Parameters)
                    {


                        string nome = param.Definition.Name;
                        string tipo = param.StorageType.ToString();
                        switch (param.StorageType)
                        {
                            case StorageType.String:
                                if (param.HasValue)
                                    if (!string.IsNullOrEmpty(param.AsString()))
                                        lista.Add(new ItemInformacaoElemento { Nome = nome, Origem = false, Valor = param.AsString(), ValorNatural = 0, Id = ele.Id.IntegerValue });
                                break;


                            case StorageType.Double:
                            case StorageType.Integer:
                                if (param.HasValue)
                                    if (!string.IsNullOrEmpty(param.AsValueString()))
                                        lista.Add(new ItemInformacaoElemento { Nome = nome, Origem = false, ValorNatural = param.AsDouble(), Valor = param.AsValueString(), Id = ele.Id.IntegerValue, parameterId = param.Id.IntegerValue });
                                break;
                        }

                    }
                }

            }

            Apresentacao.VerificarParametros verificarParametros = new Apresentacao.VerificarParametros(uiDoc, 
                new List<Element>(), 
                lista);
            verificarParametros.ShowDialog();

            
            


            return Autodesk.Revit.UI.Result.Succeeded;


        }
    }
}
