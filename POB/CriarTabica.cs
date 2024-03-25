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

    public class CriarTabica1 : IExternalCommand
    {

        FamilySymbol fs1;


        private void CreateSweep(ExternalCommandData revit, Element ele, UIApplication uiDoc, ref Document m_familyDocument,
            ref RevitCreation.FamilyItemFactory m_creationFamily, ref XYZ P)
        {

            Autodesk.Revit.ApplicationServices.Application m_revit = revit.Application.Application;

            //Document profile_doc = m_revit.OpenDocumentFile("");

            Face f = Funcoes.Util.GetTopFace(Funcoes.Util.GetSolids(ele)[0]);
            ElementosPontos novoPonto = new ElementosPontos();
            CurveArray percuso = new CurveArray();
           CurveLoop cl = f.GetEdgesAsCurveLoops()[0];
        
                CurveLoopIterator cli = cl.GetCurveLoopIterator();

                while (cli.MoveNext())
                {
                    Line l = (cli.Current as Line);
                    XYZ p1 = new XYZ();
                    p1 = l.GetEndPoint(0);
                    XYZ p10 = new XYZ(p1.X, p1.Y, 0);
                    novoPonto.pontos.Add(p10);
                    XYZ p2 = new XYZ();

                    p2 = l.GetEndPoint(1);
                    XYZ p20 = new XYZ(p2.X, p2.Y, 0);
                    novoPonto.pontos.Add(p20);
                    Curve curve1 = Line.CreateBound(p10, p20);
                    percuso.Append(curve1);
                    P = p20;
                }
       

            #region Create rectangular profile and path curve
            CurveArrArray arrarr = new CurveArrArray();
            CurveArray arr = new CurveArray();

            Autodesk.Revit.DB.XYZ normal = Autodesk.Revit.DB.XYZ.BasisZ;
            SketchPlane sketchPlane = Funcoes.Util.CreateSketchPlane(normal, Autodesk.Revit.DB.XYZ.Zero,
                                                                                                          m_familyDocument, uiDoc);

            Autodesk.Revit.DB.XYZ pnt1 = new Autodesk.Revit.DB.XYZ(0.02 / 0.3048, 0, 0);
            Autodesk.Revit.DB.XYZ pnt2 = new Autodesk.Revit.DB.XYZ(0.02 / 0.3048, 0.02 / 0.3048, 0);
            Autodesk.Revit.DB.XYZ pnt3 = new Autodesk.Revit.DB.XYZ(-0.03 / 0.3048, 0.02 / 0.3048, 0);
            Autodesk.Revit.DB.XYZ pnt4 = new Autodesk.Revit.DB.XYZ(-0.03 / 0.3048, 0, 0);
            arr.Append(Line.CreateBound(pnt1, pnt2));
            arr.Append(Line.CreateBound(pnt2, pnt3));
            arr.Append(Line.CreateBound(pnt3, pnt4));
            arr.Append(Line.CreateBound(pnt4, pnt1));


            /* arr.Append(Arc.Create(pnt2, 1.0d, 0.0d, 180.0d, Autodesk.Revit.DB.XYZ.BasisX, Autodesk.Revit.DB.XYZ.BasisY));
             arr.Append(Arc.Create(pnt1, pnt3, pnt2));*/
            arrarr.Append(arr);
            SweepProfile profile = m_revit.Create.NewCurveLoopsProfile(arrarr);

            /* Autodesk.Revit.DB.XYZ pnt4 = new Autodesk.Revit.DB.XYZ(10, 0, 0);
             Autodesk.Revit.DB.XYZ pnt5 = new Autodesk.Revit.DB.XYZ(0, 10, 0);
             Curve curve = Line.CreateBound(pnt4, pnt5);

             CurveArray curves = new CurveArray();
             curves.Append(curve);*/
#endregion
// here create one sweep with two arcs formed the profile

#if D23 || D24


#else
            Sweep rectSweep = m_creationFamily.NewSweep(true,
                                                                                        percuso, sketchPlane, profile, 0,
                                                                                        ProfilePlaneLocation.Start);


            m_familyDocument.FamilyManager.AddParameter("Material", BuiltInParameterGroup.PG_MATERIALS, ParameterType.Material,
                                                   true);
          
            var par = ManipulaParametroCompartilhado.GetOrInsertCompartilhadoExternal(m_revit, "Comprimento tabica", "Eletrica", ParameterType.Length, true);
              m_familyDocument.FamilyManager.AddParameter(par, BuiltInParameterGroup.PG_CONSTRUCTION, true);

            try
            {
                m_familyDocument.FamilyManager.AssociateElementParameterToFamilyParameter(rectSweep.LookupParameter("Material"),
                                                                                          m_familyDocument.FamilyManager.get_Parameter("Material"));
            }
            catch
            {

            }



            // move to proper place
            Autodesk.Revit.DB.XYZ transPoint1 = new Autodesk.Revit.DB.XYZ(0, 0, 0);

            ElementTransformUtils.MoveElement(m_familyDocument, rectSweep.Id, transPoint1);
#endif
        }

        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit,
            ref string message, ElementSet elements)
        {
            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;
            XYZ P = new XYZ(0, 0, 0);



            foreach (ElementId id in sel.GetElementIds())
            {
                Element ele = uiDoc.GetElement(id);
                Document m_familyDocument = revit.Application.Application.NewFamilyDocument(@"D:\Onedrive\Engenharia\INTERNO\INTERNO - TEMPLATES REVIT\Metric Generic Model.rft");
                RevitCreation.FamilyItemFactory m_creationFamily = m_familyDocument.FamilyCreate;
                if (ele is Autodesk.Revit.DB.Ceiling)
                {

                    Transaction transaction = new Transaction(m_familyDocument, "CreateGenericModel");
                    transaction.Start();

                    CreateSweep(revit, ele, uiApp, ref m_familyDocument, ref m_creationFamily, ref P);



                    transaction.Commit();
                    Family f = m_familyDocument.LoadFamily(revit.Application.ActiveUIDocument.Document);
                    
                  


                    Transaction transaction1 = new Transaction(f.Document, "CreateGenericModel1");
                    transaction1.Start();
                    f.Name = "Tabica - " + ele.Id.ToString() + " - " + DateTime.Today.ToString("yyddMM");
                    foreach (ElementId item in f.GetFamilySymbolIds())
                    {
                        FamilySymbol fs = uiDoc.GetElement(item) as FamilySymbol;
                        fs.Name = "Tab 1";
                        if (!fs.IsActive)
                            fs.Activate();
                        fs1 = fs;

                        // Curve point = ((ele as Ceiling).Location as LocationCurve).Curve;
                    }
                    uiDoc.Regenerate();
                    // XYZ ponto =  uiApp.ActiveUIDocument.Selection.PickPoint("Selecione o ponto de inserção");
                    XYZ P1 = new XYZ(0, 0, 0);
                    FamilyInstance tabica = uiApp.ActiveUIDocument.Document.Create.NewFamilyInstance(P1,
                                 fs1, (uiDoc.GetElement((ele as Ceiling).LevelId) as Level), 0);
                    tabica.LookupParameter("Deslocamento do hospedeiro").Set((ele as Ceiling).LookupParameter("Altura do deslocamento do nível").AsDouble() + 0.03 / .3048);

                    tabica.LookupParameter("Comprimento tabica").Set((ele as Ceiling).LookupParameter("Perímetro").AsDouble());

                    transaction1.Commit();
                }
            }
            return Result.Succeeded;
        }
    }




}
