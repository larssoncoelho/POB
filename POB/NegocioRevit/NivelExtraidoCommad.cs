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

using Autodesk.Revit.ApplicationServices;
namespace POB.NegocioRevit
{
    class NivelExtraidoCommad
    {
        public static ResultadoExternalCommandData Execute(Document uiDoc,
            ICollection<ElementId> selecaoId, bool criarTransacao, List<Level> listaLevel)
        {
            ResultadoExternalCommandData resultado = new ResultadoExternalCommandData();
            Transaction t1 = new Transaction(uiDoc);
            if (criarTransacao) t1.Start("ttt");
            foreach (ElementId item in selecaoId)
            {
                Element ele = uiDoc.GetElement(item);
                if (ele == null) continue;
                try
                {

                    AssemblyInstance assemblyInstance = null;
                    if (ele.Category.Name == Category.GetCategory(ele.Document, BuiltInCategory.OST_Assemblies).Name)
                    {
                        assemblyInstance = ele as AssemblyInstance;
                        ele = ele.Document.GetElement((ele as Autodesk.Revit.DB.AssemblyInstance).GetMemberIds().ToList()[0]);
                    }
                    int levelId = 0;
                    Parameter par;
                    if (assemblyInstance == null)
                        par = ele.LookupParameter("NivelExtraido");
                    else par = assemblyInstance.LookupParameter("NivelExtraido");
                    Parameter tocPavimento;
                    if (assemblyInstance == null)
                        tocPavimento = ele.LookupParameter("tocPavimento");
                    else tocPavimento = assemblyInstance.LookupParameter("tocPavimento");

                    Level level = null;
                    XYZ ponto = new XYZ(0, 0, 0);
                    foreach (Solid solid in Util.GetSolids(ele))
                    {

                        if ((null != solid) && (solid.Faces.Size > 0) && (solid.Volume > 0))
                        {
                            foreach (Face f in solid.Faces)
                            {
                                if (f != null)
                                {
                                    ponto = Util.GetPointFace(f);
                                    if (ponto != null)
                                    {
                                        level = uiDoc.GetElement(Util.GetLevelMaisProximo(ponto, listaLevel)) as Level;
                                        levelId = level.Id.IntegerValue;
                                        break;
                                    }
                                }
                            }
                            resultado.Lista.Add(new ResultadoElemento
                            {
                                Element = level

                            });
                            par.Set(levelId);
                            if (tocPavimento != null)
                                tocPavimento.Set(level.Name);
                        }

                    }
                }
                catch (Exception e)
                {
                    resultado.Lista.Add(new ResultadoElemento { ElementId = ele.Id, Mensagem = e.Message });
                }
            }
            if (criarTransacao)
            {
                t1.Commit();
                t1.Dispose();
            }
            return resultado;
        }
    }
}
/*   if (ele.Location != null)
   {
       if ((ele.Location is LocationPoint))
           ponto = (ele.Location as LocationPoint).Point;
       else ponto = (ele.Location as LocationCurve).Curve.GetEndPoint(0);

       level = uiDoc.GetElement(Util.GetLevelMaisProximo(ponto, listaLevel)) as Level;
       levelId = level.Id.IntegerValue;
       resultado.Lista.Add(new ResultadoElemento{ Element = level });
       par.Set(levelId);
       t1.Commit();
       t1.Dispose();
   }
   else
   {*/
