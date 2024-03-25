using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using wf = System.Windows.Forms;
using wDrawing = System.Drawing;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using System.Xml;

using System.Reflection;
using System.Collections;

namespace POB.NegocioRevit
{
    public class GerarInclinacaoNoPisoCommand
    {
        public static ResultadoExternalCommandData Execute(Floor floor, bool criarTransacao, ExternalCommandData commandData)
        {
            ResultadoExternalCommandData resultadoExternalCommandData = new ResultadoExternalCommandData();
            Transaction transaction = new Transaction(floor.Document);
            var pontoSelecionado =Util. GetPointRevit(commandData);

            var coordenadaZMaisAlto = Util.GetPointsDoPiso(floor).Max(x=>x.Z);
            
            pontoSelecionado = new XYZ(pontoSelecionado.X, pontoSelecionado.Y,coordenadaZMaisAlto);

            var inclinacao = floor.LookupParameter("tocInclinacao").AsDouble();
            if (pontoSelecionado == null)
                return new ResultadoExternalCommandData
                {
                    ErroGlobal = true,
                    Mensagem = "Não foi obetido o ponto",
                    Resultado = Result.Cancelled
                };
            List<XYZ> listaDePontos = new List<XYZ>();



            if (criarTransacao) transaction.Start("Pontos");
            SlabShapeEditor m_slabShapeEditor = floor.SlabShapeEditor;

          /*  var openings = floor.FindInserts(true, true, true, true);
            foreach (var item in openings)
            {
                if (floor.Document.GetElement(item).Category.Name ==
                    Category.GetCategory(floor.Document, BuiltInCategory.OST_FloorOpening).Name)
                    floor.Document.Delete(item);
             }*/

           
            m_slabShapeEditor.ResetSlabShape();
            SlabProfile m_slabProfile = new SlabProfile(floor, commandData);

           var vertexInicial =  m_slabProfile.AddVertex(pontoSelecionado, false);
            transaction.Commit();
            transaction.Start("d");

            foreach (Solid solid in Util.GetSolids(floor))
            {
                if (solid != null)
                {
                   
                    var face = Util.GetTopFace(solid);
                    Util.GetPointsFace(ref listaDePontos, face);
                }
            }

            listaDePontos = listaDePontos.Distinct().ToList();

           
          

            var listaVertexExtendido = new List<ObjetoTransferenciaPOB.SlabShapeVertexExtendido>();

            foreach (var ponto in listaDePontos)
            {
                listaVertexExtendido.Add(new ObjetoTransferenciaPOB.SlabShapeVertexExtendido
                {
                    Vertex = m_slabProfile.AddVertex(ponto, false),
                    Distancia = Util.GetDistanciaEntreDoisPontos(ponto, pontoSelecionado),
                    PontoOriginal = ponto
                });
            }
            foreach (var vertexEx in listaVertexExtendido)
            {
                try
                {
                    m_slabShapeEditor.ModifySubElement(vertexEx.Vertex,vertexEx.Distancia*inclinacao);
                }
                catch (Exception e)
                {
                    System.Console.WriteLine(e.Message);
                    resultadoExternalCommandData.Lista.Add(new ResultadoElemento
                    {
                        ElementId = new ElementId(-1),
                        Mensagem = e.Message
                    });
                }
            }
            double menorDitancia = listaVertexExtendido.Min(x => x.Distancia);

#if D24
            //m_slabShapeEditor.ModifySubElement(vertexInicial, menorDitancia * inclinacao);
#else
            m_slabShapeEditor.ModifySubElement(vertexInicial, menorDitancia * inclinacao);
#endif

            if (criarTransacao) transaction.Commit();

            if (criarTransacao) transaction.Start("sd");

            CurveArray percuso = new CurveArray();
            Plane plane = Plane.CreateByNormalAndOrigin(new XYZ(0, 0, 1), pontoSelecionado);
            Arc arc1 = Arc.Create(plane,  0.1/0.3048/2 ,0, Math.PI);
            Arc arc2 = Arc.Create(plane, 0.1 / 0.3048/2, Math.PI, Math.PI * 2);
            percuso.Append(arc1);
            percuso.Append(arc2);

            Opening opening = floor.Document.Create.NewOpening(floor, percuso ,true);
            if (criarTransacao) transaction.Commit();
            return resultadoExternalCommandData;
        }
    }
}