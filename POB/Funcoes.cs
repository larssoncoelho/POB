using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using wf = System.Windows.Forms;
using wfv = System.Windows.Forms.View;
using wDrawing = System.Drawing;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using System.Xml;
using sd = System.Drawing;
using System.Reflection;
using System.Collections;
using Autodesk.Revit.DB.Mechanical;
using ObjetoDeTranferencia;
using POB.ObjetoDeTranferencia;
using System.Runtime.InteropServices;
using System.Drawing.Imaging;
using System.Drawing;
namespace POB
{


    public static class Perguntar
    {
       
        public static wf.DialogResult InputBox(string title, string promptText, ref string value)
        {
            wf.Form form = new wf.Form();
            wf.Label label = new wf.Label();
            wf.TextBox textBox = new wf.TextBox();
            wf.Button buttonOk = new wf.Button();
            wf.Button buttonCancel = new wf.Button();

            form.Text = title;
            label.Text = promptText;
            textBox.Text = value;

            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = wf.DialogResult.OK;
            buttonCancel.DialogResult = wf.DialogResult.Cancel;

            label.SetBounds(9, 20, 372, 13);
            textBox.SetBounds(12, 36, 372, 20);
            buttonOk.SetBounds(228, 72, 75, 23);
            buttonCancel.SetBounds(309, 72, 75, 23);

            label.AutoSize = true;
            textBox.Anchor = textBox.Anchor | wf.AnchorStyles.Right;
            buttonOk.Anchor = wf.AnchorStyles.Bottom | wf.AnchorStyles.Right;
            buttonCancel.Anchor = wf.AnchorStyles.Bottom | wf.AnchorStyles.Right;
            textBox.Text = "Nivel01;Nivel02";
            form.ClientSize = new sd.Size(396, 107);
            form.Controls.AddRange(new wf.Control[] { label, textBox, buttonOk, buttonCancel });
            form.ClientSize = new sd.Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
            form.FormBorderStyle = wf.FormBorderStyle.FixedDialog;
            form.StartPosition = wf.FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;

            wf. DialogResult dialogResult = form.ShowDialog();
            value = textBox.Text;
            return dialogResult;
        }
    }

    public class CmdRelationshipInverter

    {


        private Dictionary<Autodesk.Revit.DB.ElementId, List<Autodesk.Revit.DB.ElementId>> getElementIds(List<Autodesk.Revit.DB.Element> elements)
        {
            Dictionary<Autodesk.Revit.DB.ElementId, List<Autodesk.Revit.DB.ElementId>> dict = new Dictionary<Autodesk.Revit.DB.ElementId, List<Autodesk.Revit.DB.ElementId>>();

            string fmt = "{0} is hosted by {1}";

            foreach (Autodesk.Revit.DB.FamilyInstance fi in elements)
            {
                Autodesk.Revit.DB.ElementId id = fi.Id;
                Autodesk.Revit.DB.ElementId idHost = fi.Host.Id;
                if (!dict.ContainsKey(idHost))
                {
                    dict.Add(idHost, new List<Autodesk.Revit.DB.ElementId>());
                }
                dict[idHost].Add(id);
            }
            return dict;
        }

        private void dumpHostedElements(Dictionary<Autodesk.Revit.DB.ElementId, List<Autodesk.Revit.DB.ElementId>> ids)
        {
            foreach (Autodesk.Revit.DB.ElementId idHost in ids.Keys)
            {
                string s = string.Empty;

                foreach (Autodesk.Revit.DB.ElementId id in ids[idHost])
                {
                    if (0 < s.Length)
                    {
                        s += ", ";
                    }
                    //         s += ElementDescription( id );
                }


            }
        }


    }



    public static class Resultado
    {
        static public List<object> vCampo = new List<object>();
        static public void Limpar()
        {
            vCampo.Clear();
        }

    }
    public static class Mestrado
    {

        public static void DitanciaWall(Autodesk.Revit.DB.Document doc, UIDocument uidoc, Autodesk.Revit.DB.Wall wall)
        {
            Autodesk.Revit.DB.LocationCurve locCurve = wall.Location as Autodesk.Revit.DB.LocationCurve;
            Autodesk.Revit.DB.Curve curve = locCurve.Curve;
            // Find the direction of the tangent to the curve at the curve's midpoint
            // ComputeDerivates finds 3 vectors that describe the curve at the specified point
            // BasisX is the tangent vector
            // "true" means that the point on the curve will be described with a normalized value
            // 0 = one end of the curve, 0.5 = middle, 1.0 = other end of the curve
            Autodesk.Revit.DB.XYZ curveTangent = curve.ComputeDerivatives(0.5, true).BasisX;
            // Autodesk.Revit.DB.XYZ coordinate of one of the wall's endpoints
            Autodesk.Revit.DB.XYZ wallEnd = curve.GetEndPoint(0);

            // Filter to find only doors
            Autodesk.Revit.DB.ElementCategoryFilter filter = new Autodesk.Revit.DB.ElementCategoryFilter(Autodesk.Revit.DB.BuiltInCategory.OST_Doors);

            // Autodesk.Revit.DB.ReferenceIntersector uses the specified filter, find elements, in a 3d view (which is the active view for this example)
            Autodesk.Revit.DB.ReferenceIntersector intersector = new Autodesk.Revit.DB.ReferenceIntersector(filter, Autodesk.Revit.DB.FindReferenceTarget.Element, (Autodesk.Revit.DB.View3D)doc.ActiveView);

            string doorDistanceInfo = "";
            // Autodesk.Revit.DB.ReferenceIntersector.Find shoots a ray in the specified direction from the specified point
            // it returns a list of ReferenceWithContext objects
            foreach (Autodesk.Revit.DB.ReferenceWithContext refWithContext in intersector.Find(wallEnd, curveTangent))
            {
                // Each ReferenceWithContext object specifies the distance from the ray's origin to the object (proximity)
                // and the reference of the Autodesk.Revit.DB.Element hit by the ray
                double proximity = refWithContext.Proximity;
                Autodesk.Revit.DB.FamilyInstance door = doc.GetElement(refWithContext.GetReference()) as Autodesk.Revit.DB.FamilyInstance;
                doorDistanceInfo += door.Symbol.Family.Name + " - " + door.Name + " = " + proximity + "\n";
            }
            TaskDialog.Show("Autodesk.Revit.DB.ReferenceIntersector", doorDistanceInfo);

        }
    }
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int Left;   // Coordenada do lado esquerdo
        public int Top;    // Coordenada do topo
        public int Right;  // Coordenada do lado direito
        public int Bottom; // Coordenada da parte inferior
    }

    public static class Util
    {
        public static Autodesk.Revit.DB.Document uiDoc;
        public static double volumePorParametro = 0;
        public static Autodesk.Revit.DB.FilteredElementCollector selecao;
        #region Unit conversion
        const double _convertFootToMm = 12 * 25.4;
        public static double area;
        public static double v;
        public static Autodesk.Revit.DB.Level medicaoBlocoId;
        /// <summary>
        /// Convert a given length in millimetres to feet.
        /// </summary>
        /// 
        [DllImport("user32.dll")]
        public static extern IntPtr WindowFromPoint(System.Drawing.Point point);

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

        [DllImport("user32.dll")]
        public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
        public static void CaptureRevitView(RECT rect)
        {
            /*// Obter o handle da janela ativa (Revit)
            IntPtr hwnd = GetForegroundWindow();

            // Obter as dimensões da janela ativa
            System.Drawing.Rectangle rect = new System.Drawing.Rectangle();
            GetWindowRect(hwnd, ref rect);

            // Determinar a área da visualização (ajuste conforme necessário)
            int offsetX = 10; // Exemplo: ajuste para barras de rolagem
            int offsetY = 80; // Exemplo: ajuste para menus do Revit
            System.Drawing.Rectangle viewRect = new System.Drawing.Rectangle(
                rect.X + offsetX,
                rect.Y + offsetY,
                rect.Width - offsetX * 2,
                rect.Height - offsetY - 10 // Ajustar a borda inferior
            );

            // Capturar a janela de visualização
            */
            System.Drawing.Rectangle viewRect = new System.Drawing.Rectangle(rect.Left, rect.Top, rect.Right - rect.Left, rect.Bottom - rect.Top);
            using (Bitmap screenshot = new Bitmap(viewRect.Width, viewRect.Height))
            {
                using (Graphics graphics = Graphics.FromImage(screenshot))
                {
                    graphics.CopyFromScreen(viewRect.Location, System.Drawing.Point.Empty, viewRect.Size);
                }

                // Salvar e copiar para a área de transferência
                string filePath = @"d:\ExportedView.BMP";
                screenshot.Save(filePath, ImageFormat.Bmp);
                wf.Clipboard.SetImage(screenshot);

                // MessageBox.Show($"A captura foi salva em {filePath} e copiada para a área de transferência.", "Captura de Tela");
            }
        }
        public static RECT CaptureWindowRect()
        {
            RECT windowRect = new RECT();
            // Esperar o clique do mouse
            //wf.MessageBox.Show("Clique na janela que deseja capturar.", "Captura de Janela");

            // Pegar a posição atual do cursor
            System.Drawing.Point cursorPosition = wf.Cursor.Position;

            // Obter o handle da janela sob o cursor
            IntPtr hWnd = WindowFromPoint(cursorPosition);

            if (hWnd != IntPtr.Zero)
            {
                // Obter o retângulo da janela
                if (GetWindowRect(hWnd, out windowRect))
                {
                    // Calcular as dimensões da janela
                    int width = windowRect.Right - windowRect.Left;
                    int height = windowRect.Bottom - windowRect.Top;

                    // Obter o título da janela (opcional)
                    System.Text.StringBuilder title = new System.Text.StringBuilder(256);
                    GetWindowText(hWnd, title, title.Capacity);

                    // Mostrar informações da janela
                    //wf.MessageBox.Show($"Janela: {title}\nPosição: ({windowRect.Left}, {windowRect.Top})\n" +
                    //                 $"Dimensões: {width}x{height}", "Informações da Janela");
                }
                else
                {
                    wf.MessageBox.Show("Não foi possível capturar o retângulo da janela.", "Erro");
                }
            }
            else
            {
                wf.MessageBox.Show("Nenhuma janela encontrada sob o cursor.", "Erro");
            }
            return windowRect;
        }

        public static void BuscarAninhados(Autodesk.Revit.DB.ElementId eleId, Autodesk.Revit.DB.Document uiDoc, ref List<Autodesk.Revit.DB.ElementId> listaEle)
        {
            var ele = uiDoc.GetElement(eleId);
            if (ele is Autodesk.Revit.DB.FamilyInstance)
            {
                Autodesk.Revit.DB.FamilyInstance aFamilyInst = ele as Autodesk.Revit.DB.FamilyInstance;
                if (aFamilyInst.SuperComponent == null)
                {
                    var subElements = aFamilyInst.GetSubComponentIds();
                    if (subElements.Count != 0)
                    {
                        foreach (var aSubElemId in subElements)
                        {
                            var aSubElem = ele.Document.GetElement(aSubElemId);
                            if (aSubElem is Autodesk.Revit.DB.FamilyInstance)
                            {
                                listaEle.Add(aSubElem.Id);
                                BuscarAninhados(aSubElem.Id, uiDoc, ref listaEle);
                            }
                        }
                    }
                }
                else
                {
                    var subElements1 = aFamilyInst.GetSubComponentIds();
                    if (subElements1.Count != 0)
                    {
                        foreach (var aSubElemId in subElements1)
                        {
                            var aSubElem1 = ele.Document.GetElement(aSubElemId);
                            if (aSubElem1 is Autodesk.Revit.DB.FamilyInstance)
                            {
                                listaEle.Add(aSubElem1.Id);
                                BuscarAninhados(aSubElem1.Id, uiDoc, ref listaEle);
                            }
                        }
                    }
                }
            }
        }

       /* public static  void CreateSatFileFor(Autodesk.Revit.DB.Element e, IList<Autodesk.Revit.DB.Element> allElements, string filename_prefix)
        {
            Autodesk.Revit.DB.Document doc = e.Document;

            // Keep this Autodesk.Revit.DB.Element visible and 
            // hide all other model elements

            Autodesk.Revit.DB.Transaction trans = new Autodesk.Revit.DB.Transaction(doc);
            trans.Start("Hide Elements");

            // Create Autodesk.Revit.DB.Element set other than current wall/floor 

            var eset = new List<Autodesk.Revit.DB.ElementId>();

            foreach (Autodesk.Revit.DB.Element ele in allElements)
            {
                if (e.Id.IntegerValue != ele.Id.IntegerValue)
                {
                    if (ele.CanBeHidden(doc.ActiveView))
                        eset.Add(ele.Id);
                }
            }

            // Hide all elements other than current 
            // one in current view

            doc.ActiveView.HideElements(eset);

            // Commit the Autodesk.Revit.DB.Transaction so that 
            // visibility is affected

            trans.Commit();

            // Export the ActiveView containing current 
            // Autodesk.Revit.DB.Element to SAT file

            SATExportOptions satExportOptions = new SATExportOptions();

            ViewSet vset = new ViewSet();
            vset.Insert(doc.ActiveView);

            // Get the material information

        //    IEnumerator<Material> mats =  e.Materials.Cast<Material>().GetEnumerator();

            // Get the last material, as walls may 
            // have multiple materials assigned

            Material mat = null;

        /*   while (mats.MoveNext())
            {
                mat = mats.Current;
            }
         
            // Get temp folder path for saving SAT files

            string folder = System.IO.Path.GetTempPath();

            string filename = filename_prefix
              + "-" + e.Id.ToString()
              + "-" + mat.Name
              + "-" + mat.Id.ToString()
              + ".sat";

          // doc.Export(folder, filename, vset,  satExportOptions);

            // After export make all model elements visible

            trans = new Autodesk.Revit.DB.Transaction(doc);
            trans.Start("Unhide Elements");
         //   doc.ActiveView.Unhide(eset);
            trans.Commit();
        }*/
        public static List<Autodesk.Revit.DB.Element> GetAllModelElements(Autodesk.Revit.DB.Document doc)
        {
            List<Autodesk.Revit.DB.Element> elements = new List<Autodesk.Revit.DB.Element>();

            Autodesk.Revit.DB.FilteredElementCollector collector
              = new Autodesk.Revit.DB.FilteredElementCollector(doc)
                .WhereElementIsNotElementType();

            foreach (Autodesk.Revit.DB.Element e in collector)
            {
                if (null != e.Category
                  && e.Category.HasMaterialQuantities)
                {
                    elements.Add(e);
                }
            }
            return elements;
        }
        public static Autodesk.Revit.DB.XYZ GetPointRevit(ExternalCommandData revit)
        {
            Selection sel = revit.Application.ActiveUIDocument.Selection;
           return  sel.PickPoint("Selecione o ponto de queda");
        }

        public static Autodesk.Revit.DB.AssemblyInstance CreateAssembly(Autodesk.Revit.DB.Document doc, 
                                                                            ICollection<Autodesk.Revit.DB.ElementId> elementIds)
        {
            Autodesk.Revit.DB.Transaction transaction = new Autodesk.Revit.DB.Transaction(doc);
 

            Autodesk.Revit.DB.ElementId categoryId = doc.GetElement(elementIds.First()).Category.Id; // use category of one of the assembly elements
            if (Autodesk.Revit.DB.AssemblyInstance.IsValidNamingCategory(doc, categoryId, elementIds))
            {
                transaction.Start("Create Assembly Instance");
                Autodesk.Revit.DB.AssemblyInstance assemblyInstance = Autodesk.Revit.DB.AssemblyInstance.Create(doc, elementIds, categoryId);
                transaction.Commit(); // commit the Autodesk.Revit.DB.Transaction that creates the assembly instance before modifying the instance's name
                return assemblyInstance;
            }
            return null;
        }

        public static double GetDistanciaEntreDoisPontos(Autodesk.Revit.DB.XYZ ponto1, Autodesk.Revit.DB.XYZ item)
        {
            double dist = 0;
            double x2 = Math.Pow((item.X - ponto1.X), 2);
            double y2 = Math.Pow((item.Y - ponto1.Y), 2);
            dist = Math.Sqrt(x2 + y2);
            return dist;

        }
       
      public static Autodesk.Revit.DB.XYZ GetPointFace(Autodesk.Revit.DB.Face f)
        {
            List<Autodesk.Revit.DB.CurveLoop> curveLoops = f.GetEdgesAsCurveLoops().ToList();
            foreach (Autodesk.Revit.DB.CurveLoop curveLoops2 in curveLoops)
            {
                foreach (Autodesk.Revit.DB.Curve curve in curveLoops2)
                {
                    return curve.GetEndPoint(0);


                }
            }
            return null;
        }
        public static void GetPointsFace(ref List<Autodesk.Revit.DB.XYZ>  listaDePontos, Autodesk.Revit.DB.Face f)
        {

            List<Autodesk.Revit.DB.CurveLoop> curveLoops = f.GetEdgesAsCurveLoops().ToList();
            foreach (Autodesk.Revit.DB.CurveLoop curveLoops2 in curveLoops)
            {
                foreach (Autodesk.Revit.DB.Curve curve in curveLoops2)
                {
                    listaDePontos.Add(curve.GetEndPoint(0));
                    listaDePontos.Add(curve.GetEndPoint(1));
                }
            }
           
        }
        public static List<Autodesk.Revit.DB.XYZ> GetPointsDoPiso(Autodesk.Revit.DB.Floor floor)
        {
            List<Autodesk.Revit.DB.XYZ> points = new List<Autodesk.Revit.DB.XYZ>();

            foreach (var solid in Util.GetSolids(floor))
            {
                if (solid != null)
                {
                    foreach (var face in solid.Faces.Cast<Autodesk.Revit.DB.Face>().ToList())
                    {
                        if (face != null) Util.GetPointsFace(ref points, face);
                    }
                }
            }
            return points;
        }

        internal static ObjetoTransferenciaPOB.DadosChapa  GetChapa(Duct duto, List<ObjetoTransferenciaPOB.DadosChapa> lista1)
        {
            var v1 = duto.LookupParameter("Altura").AsDouble();
            var v2 = duto.LookupParameter("Largura").AsDouble();
            double v3 = 0;
            if (v1 > v2)
                v3 = v1;
            else v3 = v2;
            lista1 = lista1.OrderByDescending(x => x.MaiorLado).ToList();
            return lista1.Where(x => x.MaiorLado <= v3).FirstOrDefault();
        }

        /* public static List<Autodesk.Revit.DB.XYZ> GetPointsDaFaceSuperiorDoPiso(Autodesk.Revit.DB.Floor floor)
 {
     List<Autodesk.Revit.DB.XYZ> points = new List<Autodesk.Revit.DB.XYZ>();

     foreach (var solid in Util.GetSolids(floor))
     {
         if (solid != null)
         {
             var faceBotom = GetBottonFace(solid);
             if(faceBotom!=null) Util.GetPointsFace(ref points, faceBotom);

         }
     }
     return points;
 }*/

        public static bool GetElementLocation(out Autodesk.Revit.DB.XYZ p, Autodesk.Revit.DB.Element e)
        {
            p = Autodesk.Revit.DB.XYZ.Zero;
            bool rc = false;
            Autodesk.Revit.DB.Location loc = e.Location;
            if (null != loc)
            {
                Autodesk.Revit.DB.LocationPoint lp = loc as Autodesk.Revit.DB.LocationPoint;
                if (null != lp)
                {
                    p = lp.Point;
                    rc = true;
                }
                else
                {
                    Autodesk.Revit.DB.LocationCurve lc = loc as Autodesk.Revit.DB.LocationCurve;


                    p = lc.Curve.GetEndPoint(0);
                    rc = true;
                }
            }
            return rc;
        }

#if D23 || D24
        public static Autodesk.Revit.DB.Parameter GetParameter(Autodesk.Revit.DB.Element ele, string nome, Autodesk.Revit.DB.ForgeTypeId tipo, bool instancia, bool criarTransacao)
        {
            var categorySet = new Autodesk.Revit.DB.CategorySet();
            categorySet.Insert(ele.Category);
            try
            {
                Autodesk.Revit.DB.Parameter par = ele.LookupParameter(nome);
                if (par != null)
                {
                    if (par.Definition.GetDataType()!= tipo)
                    {
                        TaskDialog.Show("15575", "O parâmetro " + par.Definition.Name + " está cadastrado no Revit com um tipo diferente," +
                            " deletar do revit e do arquivo de parâmetro compartilhado");
                    }
                    return par;
                }
                else return InserirParametroCompartilhado(ele, nome, tipo, instancia, false, criarTransacao);
            }
            catch (Exception e)
            {
                return null;
            }
        }
        public static void InserirParametroCompartilhadoPorCategoria(Autodesk.Revit.DB.Document uiDoc, string nome, bool instancia,
                                                                                             bool elementId, bool criarTransacao, Autodesk.Revit.DB.CategorySet categorySet,
                                                                                             Autodesk.Revit.DB.ForgeTypeId parameterType)
        {


            Autodesk.Revit.DB.Transaction t = new Autodesk.Revit.DB.Transaction(uiDoc);
            if (criarTransacao) t.Start("Inserir parametro");


            try
            {
                ManipulaParametroCompartilhado.InserirParametroCompartilhadoNoProjeto(uiDoc.Application,
                                        categorySet,
                                       Autodesk.Revit.DB.BuiltInParameterGroup.PG_TEXT,
                                        instancia,
                                        "Eletrica",
                                        nome,
                                        parameterType);


                if (criarTransacao) t.Commit();
                if (criarTransacao) t.Dispose();

            }
            catch (Exception e4)
            {


                if (criarTransacao) t.RollBack();

                if (criarTransacao) t.Dispose();

            }


        }

        public static Autodesk.Revit.DB.Parameter InserirParametroCompartilhado(
            Autodesk.Revit.DB.Element ele,
            string nome,
            Autodesk.Revit.DB.ForgeTypeId tipo,
            bool instancia,
            bool elementId,
            bool criarTransacao)
        {
            Autodesk.Revit.DB.CategorySet categorySet = new Autodesk.Revit.DB.CategorySet();
            categorySet.Insert(ele.Category);


            Autodesk.Revit.DB.Transaction t = new Autodesk.Revit.DB.Transaction(ele.Document);
            if (criarTransacao) t.Start("Inserir parametro");


            try
            {
                ManipulaParametroCompartilhado.InserirParametroCompartilhadoNoProjeto(ele.Document.Application,
                                        categorySet,
                                       Autodesk.Revit.DB.BuiltInParameterGroup.PG_TEXT,
                                        instancia,
                                        "Eletrica",
                                        nome,
                                        tipo);


                if (criarTransacao) t.Commit();
                if (criarTransacao) t.Dispose();

            }
            catch (Exception e4)
            {


                if (criarTransacao) t.RollBack();

                if (criarTransacao) t.Dispose();

            }
            return ele.LookupParameter(nome);

        }


#if D24 || D23
        public static Autodesk.Revit.DB.ForgeTypeId ObterTipoDeParametro(Type tipo, bool elementId)
        {
            if (tipo == typeof(string)) return Autodesk.Revit.DB.SpecTypeId.String.Text;
            if (tipo == typeof(double)) return Autodesk.Revit.DB.SpecTypeId.Number;
            if (tipo == typeof(bool)) return Autodesk.Revit.DB.SpecTypeId.Boolean.YesNo ;
            if (tipo == typeof(int) & (!elementId)) return Autodesk.Revit.DB.SpecTypeId.Int.Integer;


            return Autodesk.Revit.DB.SpecTypeId.String.Text;
        }

#else
 public static ParameterType ObterTipoDeParametro(Type tipo, bool Autodesk.Revit.DB.ElementId)
        {
            if (tipo == typeof(string)) return ParameterType.Text;
            if (tipo == typeof(double)) return ParameterType.Number;
            if (tipo == typeof(bool)) return ParameterType.YesNo; ;
            if (tipo == typeof(int) & (!Autodesk.Revit.DB.ElementId)) return ParameterType.Integer;
            if (tipo == typeof(int) & (Autodesk.Revit.DB.ElementId)) return ParameterType.FamilyType;


            return ParameterType.Text;
        }
#endif

        public static void GetParameter(Autodesk.Revit.DB.Document uiDoc, Autodesk.Revit.DB.CategorySet filtro, string nome, Autodesk.Revit.DB.ForgeTypeId tipo, bool instancia, bool criarTransacao)
        {
            try
            {
                InserirParametroCompartilhadoPorCategoria(uiDoc, nome, instancia, false, criarTransacao, filtro, tipo);
            }
            catch (Exception e)
            {

            }
        }

#else
        public static Parameter GetParameter(Autodesk.Revit.DB.Element ele, string nome, ParameterType tipo, bool instancia, bool criarTransacao)
        {
          var  categorySet = new CategorySet();
            categorySet.Insert(ele.Category);
            try
            {
                Parameter par = ele.LookupParameter(nome);
                if (par != null)
                {
                    if (par.Definition.ParameterType != tipo)
                    {
                        TaskDialog.Show("15575", "O parâmetro " + par.Definition.Name + " está cadastrado no Revit com um tipo diferente," +
                            " deletar do revit e do arquivo de parâmetro compartilhado");
                    }
                    return par;
                }
                else return InserirParametroCompartilhado(ele, nome, tipo, instancia, criarTransacao);
            }
            catch (Exception e)
            {
                return null;
            }
        }
          /* public static void GetParameter(Autodesk.Revit.DB.Document uiDoc, CategorySet filtro, string nome, ParameterType tipo, bool instancia, bool criarTransacao)
        {
            try
            {
                InserirParametroCompartilhadoPorCategoria(uiDoc, nome,  instancia, criarTransacao, filtro, tipo);
            }
            catch (Exception e)
            {
              
            }
        }*/
           public static void InserirParametroCompartilhadoPorCategoria(Autodesk.Revit.DB.Document uiDoc,
                                                                                                          string nome, 
                                                                                                          bool instancia, 
                                                                                                          bool criarTransacao, CategorySet categorySet,
                                                                                                ParameterType parameterType)
        {
           

            Autodesk.Revit.DB.Transaction t = new Autodesk.Revit.DB.Transaction(uiDoc);
           if(criarTransacao) t.Start("Inserir parametro");


            try
            {
                ManipulaParametroCompartilhado.InserirParametroCompartilhadoNoProjeto(uiDoc.Application,
                                        categorySet,
                                        BuiltInParameterGroup.PG_TEXT,
                                        instancia,
                                        "Eletrica",
                                        nome,
                                        parameterType);


                if (criarTransacao) t.Commit();
                if (criarTransacao) t.Dispose();

            }
            catch (Exception e4)
            {


                if (criarTransacao) t.RollBack();

                if (criarTransacao) t.Dispose();

            }
          

        }

        public static Parameter InserirParametroCompartilhado(
            Autodesk.Revit.DB.Element ele, 
            string nome, 
            ParameterType tipo, 
            bool instancia, 
            bool criarTransacao)
        {
            CategorySet categorySet = new CategorySet();
            categorySet.Insert(ele.Category);


            Autodesk.Revit.DB.Transaction t = new Autodesk.Revit.DB.Transaction(ele.Document);
        if(criarTransacao)    t.Start("Inserir parametro");


            try
            {
                ManipulaParametroCompartilhado.InserirParametroCompartilhadoNoProjeto(ele.Document.Application,
                                        categorySet,
                                        BuiltInParameterGroup.PG_TEXT,
                                        instancia,
                                        "Eletrica",
                                        nome,
                                        tipo);


                if (criarTransacao) t.Commit();
                if (criarTransacao) t.Dispose();

            }
            catch (Exception e4)
            {


                if (criarTransacao) t.RollBack();

                if (criarTransacao) t.Dispose();

            }
            return ele.LookupParameter(nome);

        }



        public static ParameterType ObterTipoDeParametro(Type tipo, bool eFamilia)
        {
            if (tipo == typeof(string)) return ParameterType.Text;
            if (tipo == typeof(double)) return ParameterType.Number;
            if (tipo == typeof(bool)) return ParameterType.YesNo; ;
            if (tipo == typeof(int) & (!eFamilia)) return ParameterType.Integer;
            if (tipo == typeof(int) & (eFamilia)) return ParameterType.FamilyType;
            return ParameterType.Text;
        }
      
         public static void GetParameter(Autodesk.Revit.DB.Document uiDoc, CategorySet filtro, string nome, ParameterType tipo, bool instancia, bool criarTransacao)
        {
            try
            {
                InserirParametroCompartilhadoPorCategoria(uiDoc, nome,  instancia, criarTransacao, filtro, tipo);
            }
            catch (Exception e)
            {
              
            }
        }

#endif


        public static ICollection<Autodesk.Revit.DB.Element> FindElementsInterfering(Autodesk.Revit.DB.Element ele, List<Autodesk.Revit.DB.ElementId> excludedElements)
        {
            Autodesk.Revit.DB.FilteredElementCollector interferingCollector = new Autodesk.Revit.DB.FilteredElementCollector(ele.Document);

            interferingCollector.WhereElementIsNotElementType();
            excludedElements.Add(new Autodesk.Revit.DB.ElementId( Convert.ToInt32(ele.get_Parameter(Autodesk.Revit.DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString())));
            Autodesk.Revit.DB.ExclusionFilter exclusionFilter = new Autodesk.Revit.DB.ExclusionFilter(excludedElements);
            interferingCollector.WherePasses(exclusionFilter);
            Autodesk.Revit.DB.ElementIntersectsElementFilter intersectionFilter = new Autodesk.Revit.DB.ElementIntersectsElementFilter(ele);
            interferingCollector.WherePasses(intersectionFilter);
            return interferingCollector.ToElements();
        }

        public static List<Autodesk.Revit.DB.Curve> GetLinhaPerfilParede(Autodesk.Revit.DB.Face face)
        {
            var lista = new List<Autodesk.Revit.DB.Curve>();
            foreach (Autodesk.Revit.DB.CurveLoop curveLoop in face.GetEdgesAsCurveLoops())
            {



                Autodesk.Revit.DB.CurveLoopIterator cli = curveLoop.GetCurveLoopIterator();
            }
            return null;
 
        }
              /*  while (cli.MoveNext())
                {

                    Autodesk.Revit.DB.Curve curve = cli.curr;

                    Line linhaAtual = (cli.Current as Line);
        Autodesk.Revit.DB.XYZ p1 = linhaAtual.GetEndPoint(0);
        Autodesk.Revit.DB.XYZ p2 = linhaAtual.GetEndPoint(1);
                    if ((p1.X - p2.X) > 0)
                    {
                        m = (p1.Y - p2.Y) / (p1.X - p2.X);
                        if (m<0) return linhaAtual;
                    }
                    if (p1.X == p2.X)
                    {
                        return linhaAtual;
                    }
                }
            }

            return null;

        }
        */
        public static Autodesk.Revit.DB.SketchPlane CreateSketchPlane(Autodesk.Revit.DB.XYZ normal, Autodesk.Revit.DB.XYZ origin,
                                           Autodesk.Revit.DB.Document m_familyDocument, UIApplication uiDoc)
        {
            // First create a Geometry.Plane which need in NewSketchPlane() method
            //to-do   Plane geometryPlane = uiDoc.Application.Create.NewPlane(normal, origin);
            //to-do      if (null == geometryPlane)  // assert the creation is successful
            {
                throw new Exception("Create the geometry plane failed.");
            }
            // Then create a sketch plane using the Geometry.Plane
            //to-do     SketchPlane plane = SketchPlane.Create(m_familyDocument, geometryPlane);
            // throw exception if creation failed
            //to-do      if (null == plane)
            {
                throw new Exception("Create the sketch plane failed.");
            }
            //to-do    return plane;
            return null;
        }


        /*public static SketchPlane CreateSketchPlane(Autodesk.Revit.DB.XYZ normal, Autodesk.Revit.DB.XYZ origin,
                                          UIApplication uiDoc)
        {
            // First create a Geometry.Plane which need in NewSketchPlane() method
            //to-do    Plane geometryPlane = uiDoc.Application.Create.NewPlane(normal, origin);
            //to-do    if (null == geometryPlane)  // assert the creation is successful
            {
                throw new Exception("Create the geometry plane failed.");
            }
            // Then create a sketch plane using the Geometry.Plane
            //SketchPlane plane = SketchPlane.Create(uiDoc.ActiveUIDocument.Autodesk.Revit.DB.Document, geometryPlane);
            // throw exception if creation failed
            //to-do     if (null == plane)
            {
                throw new Exception("Create the sketch plane failed.");
            }
            //to-do     return plane;
            return null;
        }*/



        /*public static Autodesk.Revit.DB.WallType NovaPateParede(Wall wall, CompoundStructureLayer camada)
        {
            Autodesk.Revit.DB.WallType symbolNovo;

            try
            {
                Util.uiDoc = uiDoc;
                Autodesk.Revit.DB.WallType symbol = Util.FindElementByName(typeof(Autodesk.Revit.DB.WallType), wall.Name) as Autodesk.Revit.DB.WallType;
                symbolNovo = symbol.Duplicate(wall.Name + "_" + camada.Width.ToString() + "_" + camada.MaterialId.IntegerValue.ToString()) as Autodesk.Revit.DB.WallType;
                CompoundStructure structure = wall.Autodesk.Revit.DB.WallType.GetCompoundStructure();
                List<CompoundStructureLayer> lista = new List<CompoundStructureLayer>();
                lista.Add(camada);
                structure.SetLayers(lista);
                symbolNovo.SetCompoundStructure(structure);
                return symbolNovo;
            }
            catch (Exception e)
            {
                Util.uiDoc = uiDoc;

                return Util.FindElementByName(typeof(Autodesk.Revit.DB.WallType), wall.Name + "_" + camada.Width.ToString() + "_" + camada.MaterialId.IntegerValue.ToString()) as Autodesk.Revit.DB.WallType;
            }


        }
        */


      /*  public static void ProfileWall1(Autodesk.Revit.DB.Wall wall, UIApplication uiapp, UIDocument uidoc, Autodesk.Revit.UI.UIApplication.Application app)
        {


          /*  Autodesk.Revit.DB.Document doc = uidoc.Autodesk.Revit.DB.Document;
            Autodesk.Revit.DB.View view = doc.ActiveView;

            Autodesk.Revit.Creation.Application creapp  = app.Create;

            Autodesk.Revit.Creation.Autodesk.Revit.DB.Document credoc    = doc.Create;

            using (Autodesk.Revit.DB.Transaction tx = new Autodesk.Revit.DB.Transaction(doc))
            {
                tx.Start("Wall Profile");

                // Get the external wall face for the profile
                // a little bit simpler than in the last realization

                Reference sideFaceReference = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Exterior).First();

                Autodesk.Revit.DB.Face face = (Autodesk.Revit.DB.Face)wall.GetGeometryObjectFromReference(sideFaceReference);

                // The normal of the wall external face.

                Autodesk.Revit.DB.XYZ normal = wall.Orientation;

                // Offset curve copies for visibility.

                Transform offset = Transform.CreateTranslation(
                  5 * normal);

                // If the curve loop direction is counter-
                // clockwise, change its color to RED.

                Color colorRed = new Color(255, 0, 0);

                // Get edge loops as curve loops.

                IList<Autodesk.Revit.DB.CurveLoop> curveLoops = face.GetEdgesAsCurveLoops();

                foreach (var curveLoop in curveLoops)
                {
                    CurveArray curves = creapp.NewCurveArray();

                    foreach (Autodesk.Revit.DB.Curve curve in curveLoop)
                        curves.Append(curve.CreateTransformed(offset));

                    // Create model lines for an curve loop.

                    Plane plane;//to-do  = creapp.NewPlane(curves);

                    //to-do    SketchPlane sketchPlane
                    //to-do    = SketchPlane.Create(doc, plane);

                    /*    ModelCurveArray curveElements    = credoc.NewModelCurveArray(curves,   sketchPlane);

                        if (curveLoop.IsCounterclockwise(normal))
                            foreach (ModelCurve mcurve in curveElements)
                            {
                                OverrideGraphicSettings overrides
                                    = view.GetElementOverrides(
                                        mcurve.Id);

                                overrides.SetProjectionLineColor(
                                                        colorRed);

                                view.SetElementOverrides(
                                    mcurve.Id, overrides);
                            }
                        
                }
            }
        }*/
        public static DataTable ServicosSelecionados(Selection sel, Autodesk.Revit.DB.Document uiDoc)
        {
            DataTable dt = new DataTable();
            string texto = "";
            int contador = 0;
            dt.Columns.Add("tocServicoId", typeof(string));
            dt.Columns.Add("tocGrupoId", typeof(string));
            int servicoId;

            foreach (Autodesk.Revit.DB.ElementId item in sel.GetElementIds())
            {

                try
                {
                    Autodesk.Revit.DB.Element ele = uiDoc.GetElement(item);
                    servicoId = Convert.ToInt32(ele.LookupParameter("tocServicoId").AsString());
                    DataRow[] linhas = dt.Select("tocServicoId = '" + ele.LookupParameter("tocServicoId").AsString() + "'");
                    if (linhas.Length == 0)
                    {
                        DataRow dr = dt.NewRow();
                        dr["tocServicoId"] = ele.LookupParameter("tocServicoId").AsString();
                        dr["tocGrupoId"] = ele.LookupParameter("tocGrupoId").AsString();
                        dt.Rows.Add(dr);

                    }
                }
                catch
                {

                }
            }


            return dt;
        }

        public static DataTable GruposSelecionados(Selection sel, Autodesk.Revit.DB.Document uiDoc)
        {
            DataTable dt = new DataTable();
            string texto = "";
            int contador = 0;
            dt.Columns.Add("tocGrupoId", typeof(string));
            int grupoId;

            foreach (Autodesk.Revit.DB.ElementId item in sel.GetElementIds())
            {
                try
                {
                    Autodesk.Revit.DB.Element ele = uiDoc.GetElement(item);
                    grupoId = Convert.ToInt32(ele.LookupParameter("tocGrupoId").AsString());

                    if (grupoId > 0)
                    {

                        DataRow[] linhas = dt.Select("tocGrupoId = '" + ele.LookupParameter("tocGrupoId").AsString() + "'");
                        if (linhas.Length == 0)
                        {
                            DataRow dr = dt.NewRow();
                            dr["tocGrupoId"] = ele.LookupParameter("tocGrupoId").AsString();
                            dt.Rows.Add(dr);
                        }
                    }
                }
                catch
                {

                }
            }
            return dt;
        }
        public static double MmToFoot(double length)
        {
            return length / _convertFootToMm;
        }
        #endregion // Unit conversion

        #region Formatting
        /// <summary>
        /// Return an English plural suffix for the given
        /// number of items, i.e. 's' for zero or more
        /// than one, and nothing for exactly one.
        /// </summary>
        public static Autodesk.Revit.DB.Document EscolherDocumentoAtivo(Autodesk.Revit.UI.UIApplication app, string nome)
        {
            var docs = app.Application.Documents.Cast<Autodesk.Revit.DB.Document>();
            var doc = docs.Where(a => a.PathName.Contains(nome)).First();
            return app.OpenAndActivateDocument(doc.PathName).Document;
            //foreach (var d in docs) {
            //	TaskDialog.Show("_", d.PathName);
            // }

        }
        public static List<Autodesk.Revit.DB.View> GetListViewAvanco()
        {
            Autodesk.Revit.DB.FilteredElementCollector selecaoView = new Autodesk.Revit.DB.FilteredElementCollector(uiDoc).OfClass(typeof(Autodesk.Revit.DB.View));
            List<Autodesk.Revit.DB.View> vistasDeAvanco = new List<Autodesk.Revit.DB.View>();
            foreach (Autodesk.Revit.DB.View view in selecaoView)
            {
                try
                {
                    if (view.AreGraphicsOverridesAllowed())
                        if (view.LookupParameter("tocVistaAvanco").AsValueString() == "Sim")
                            vistasDeAvanco.Add(view);
                }
                catch
                {

                }

            }
            return vistasDeAvanco;
        }

        public static string PluralSuffix(int n)
        {
            return 1 == n ? "" : "s";
        }

        /// <summary>
        /// Return a dot (full stop) for zero
        /// or a colon for more than zero.
        /// </summary>
        public static string DotOrColon(int n)
        {
            return 0 < n ? ":" : ".";
        }

        #region Geometrical Comparison
        const double _eps = 1.0e-9;


        public static double AreaLajeProjecao(string markPai, string nomePai, Autodesk.Revit.DB.Element ele,
            string campoMark, string campoArea)
        {

            double areaAcumulada = 0;
            if (ele is Autodesk.Revit.DB.Architecture.BuildingPad)
                selecao = new Autodesk.Revit.DB.FilteredElementCollector(uiDoc).OfClass(typeof(Autodesk.Revit.DB.Architecture.BuildingPad));
            if (ele is Autodesk.Revit.DB.Floor)
                selecao = new Autodesk.Revit.DB.FilteredElementCollector(uiDoc).OfClass(typeof(Autodesk.Revit.DB.Floor));
            if (ele is Autodesk.Revit.DB.FootPrintRoof)
                selecao = new Autodesk.Revit.DB.FilteredElementCollector(uiDoc).OfClass(typeof(Autodesk.Revit.DB.FootPrintRoof));
            foreach (Autodesk.Revit.DB.Element ele1 in selecao)
            {
                if (ele1.LookupParameter(campoMark).AsString() == markPai)
                {
                    if (ele1.Name == nomePai)
                    {
                        areaAcumulada = areaAcumulada + ele1.LookupParameter(campoArea).AsDouble();

                    }
                }
            }





            return areaAcumulada * 0.3048 * 0.3048;
        }

       public static double GetAreaTotal(Autodesk.Revit.DB.FamilyInstance fi)
        {
            double area = 0;
            foreach (Autodesk.Revit.DB.Solid solid in GetSolids(fi))
            {
                if (solid != null)
                {
                    area = area + solid.SurfaceArea; ;
                }
            }
            return area;
        }
        internal static List<Duct> GetDutoConectados(Autodesk.Revit.DB.FamilyInstance fi)
        {

            List<Duct> dutos = new List<Duct>();
            var cm = fi.MEPModel.ConnectorManager;
            var listaConectores = cm.Connectors.Cast<Autodesk.Revit.DB.Connector>().ToList().Where(x=>x.IsConnected);
            foreach (var item in listaConectores)
            {
                var conectorDuto = item.AllRefs.Cast<Autodesk.Revit.DB.Connector>().ToList()[0];
              if(conectorDuto.Owner!=null)
                if(conectorDuto.Owner is Duct)
                    {
                        dutos.Add(conectorDuto.Owner as Duct);
                    }
            }
            return dutos;
        }

        

        public static string Retorno(string texto, string caracter, int numeroCampo)
        {
            string[] campoAlt = new string[1];
            int y = 0;

            string a = "";
            string b = "";

            for (int x = 0; x < texto.Length; x++)
            {
                a = texto.Substring(x, 1);
                if (a == caracter)
                {
                    y = y + 1;
                    Array.Resize(ref campoAlt, campoAlt.Length + 1);
                    campoAlt[y - 1] = b;
                    b = "";
                }
                else
                {
                    b = b + a;
                }
            }


            return campoAlt[numeroCampo];
        }


        static public void AlterarGraficoElemento(List<Autodesk.Revit.DB.View> selecaoView, Autodesk.Revit.DB.OverrideGraphicSettings org, Autodesk.Revit.DB.ElementId eleId)
        {
            foreach (Autodesk.Revit.DB.View view in selecaoView)
            {

                try
                {
                    if (view.AreGraphicsOverridesAllowed())
                        if (view.LookupParameter("tocVistaAvanco").AsValueString() == "Sim")
                        {
                            view.SetElementOverrides(eleId, org);
                        }
                }
                catch (Exception e)
                {

                }
            }
        }

        static public Autodesk.Revit.DB.PlanarFace GetTopFace(Autodesk.Revit.DB.Solid solid)
        {
            Autodesk.Revit.DB.PlanarFace topFace = null;
            Autodesk.Revit.DB.FaceArray faces = solid.Faces;
            foreach (Autodesk.Revit.DB.Face f in faces)
            {
                Autodesk.Revit.DB.PlanarFace pf = f as Autodesk.Revit.DB.PlanarFace;
                if (null != pf && Util.IsHorizontal(pf))
                {
                    if ((null == topFace) || (topFace.Origin.Z < pf.Origin.Z))
                    {
                        topFace = pf;
                    }
                }
            }
            return topFace;
        }




        /* double vol = 0;
         Options opt = ele.Autodesk.Revit.DB.Document.Application.Create.NewGeometryOptions();
         GeometryElement ge12 = ele.get_Geometry(opt);
         try
         {

             foreach (GeometryObject obj in ge12)
             {
                 GeometryInstance gi = obj as GeometryInstance;
                 GeometryElement gh = gi.GetInstanceGeometry();
                 foreach (GeometryObject obj1 in gh)
                 {
                     Solid solid = obj1 as Solid;
                     if (null != solid)
                     {
                         vol = vol + solid.Volume * 0.3048 * 0.3048 * 0.3048;
                     }
                 }
             }
         }
         catch
         {
             Util.InfoMsg("erro interno");
             return 0;
         }
         return vol;*/




        static public Autodesk.Revit.DB.PlanarFace GetBottonFace(Autodesk.Revit.DB.Solid solid)
        {
            Autodesk.Revit.DB.PlanarFace bottonFace = null;
            Autodesk.Revit.DB.FaceArray faces = solid.Faces;
            foreach (Autodesk.Revit.DB.Face f in faces)
            {
                Autodesk.Revit.DB.PlanarFace pf = f as Autodesk.Revit.DB.PlanarFace;
                if (null != pf
                  && Util.IsHorizontal(pf))
                {
                    if ((null == bottonFace)
                      || (bottonFace.Origin.Z > pf.Origin.Z))
                    {

                        bottonFace = pf;
                    }
                }
            }
            return bottonFace;
        }
        public static Autodesk.Revit.DB.FaceArray GetVerticalFace(Autodesk.Revit.DB.Solid solid)
        {
            //List<PlanarFace> lista = new List<PlanarFace>();
            Autodesk.Revit.DB.FaceArray faceArray = new Autodesk.Revit.DB.FaceArray();
            // lista.Clear();
            Autodesk.Revit.DB.FaceArray faces = solid.Faces;
            Autodesk.Revit.DB.PlanarFace pf;
            foreach (Autodesk.Revit.DB.Face f in faces)
            {

                pf = f as Autodesk.Revit.DB.PlanarFace;
                if (null != pf && Util.IsVertical(pf))
                {
                    faceArray.Append(pf);
                }
            }

            return faceArray;
        }
        public static List<Autodesk.Revit.DB.Face> GetVerticalFaceLista(Autodesk.Revit.DB.Solid solid)
        {
            //List<PlanarFace> lista = new List<PlanarFace>();
            List<Autodesk.Revit.DB.Face> listaFace = new List<Autodesk.Revit.DB.Face>();
            // lista.Clear();
            Autodesk.Revit.DB.FaceArray faces = solid.Faces;
            Autodesk.Revit.DB.PlanarFace pf;
            foreach (Autodesk.Revit.DB.Face f in faces)
            {

                pf = f as Autodesk.Revit.DB.PlanarFace;
                if (null != pf && Util.IsVertical(pf))
                {
                    listaFace.Add(pf);
                }
            }

            return listaFace;
        }
        public static List<Autodesk.Revit.DB.Face> GetFaceLista(Autodesk.Revit.DB.Solid solid)
        {
            //List<PlanarFace> lista = new List<PlanarFace>();
            List<Autodesk.Revit.DB.Face> listaFace = new List<Autodesk.Revit.DB.Face>();
            // lista.Clear();
            
            try
            {
                Autodesk.Revit.DB.FaceArray faces = solid.Faces;
                Autodesk.Revit.DB.PlanarFace pf;
                foreach (Autodesk.Revit.DB.Face f in faces)
                {
                    listaFace.Add((f as Autodesk.Revit.DB.PlanarFace));

                }
            }
            catch (Exception e30)
            {
                var t = e30.Message;
                return listaFace;
            }
            return listaFace;
        }
        public static Autodesk.Revit.DB.FaceArray GetHorizontalFace(Autodesk.Revit.DB.Solid solid)
        {

            Autodesk.Revit.DB.FaceArray faceArray = new Autodesk.Revit.DB.FaceArray();

            Autodesk.Revit.DB.PlanarFace horizontalFace = null;
            Autodesk.Revit.DB.FaceArray faces = solid.Faces;
            Autodesk.Revit.DB.PlanarFace pf;
            foreach (Autodesk.Revit.DB.Face f in faces)
            {
                pf = f as Autodesk.Revit.DB.PlanarFace;
                if (null != pf && Util.IsHorizontal(pf))
                {

                    faceArray.Append(pf);

                }
            }
            return faceArray;
        }
        public static double GetAreaVerical(Autodesk.Revit.DB.FaceArray faceArray)
        {
            double valor = 0;
            foreach (Autodesk.Revit.DB.PlanarFace pf in faceArray)
            {
                valor = valor + pf.Area;

            }
            return valor;
        }
        public static double GetVolumePorParametro(Autodesk.Revit.DB.Element ele2)
        {



            foreach (Autodesk.Revit.DB.Parameter param in ele2.Parameters)
            {
                try
                {
                    Util.InfoMsg(param.Definition.Name);

                    if (param.Definition.Name == "Volume")
                        volumePorParametro = param.AsDouble() * 0.3048 * 0.3048 * 0.3048;
                    continue;
                }
                catch
                {
                }

            }


            return volumePorParametro;
        }
        public static double GetAreaHorizontal(Autodesk.Revit.DB.FaceArray faceArray)
        {
            double valor = 0;
            foreach (Autodesk.Revit.DB.PlanarFace pf in faceArray)
            {
                valor = valor + pf.Area;
            }
            return valor;
        }
        public static string GetListaServicosId(ICollection<Autodesk.Revit.DB.ElementId> selecao, Autodesk.Revit.DB.Document uiDoc)
        {
            List<string> lista = new List<string>();

            foreach (Autodesk.Revit.DB.ElementId eleId in selecao)
            {

                if (!lista.Contains(uiDoc.GetElement(eleId).LookupParameter("tocServicoId").AsString()))

                    lista.Add(uiDoc.GetElement(eleId).LookupParameter("tocServicoId").AsString());




            }

            return lista.ToString();

        }
        public static double GetAreaFormaPilar(Autodesk.Revit.DB.Element ele)
        {
            double area = 0;
            Autodesk.Revit.DB.Options opt = ele.Document.Application.Create.NewGeometryOptions();
            Autodesk.Revit.DB.GeometryElement ge12 = ele.get_Geometry(opt);

            foreach (Autodesk.Revit.DB.GeometryObject obj in ge12)
            {
                Autodesk.Revit.DB.GeometryInstance gi = obj as Autodesk.Revit.DB.GeometryInstance;

                Autodesk.Revit.DB.GeometryElement gh = gi.GetInstanceGeometry();

                foreach (Autodesk.Revit.DB.GeometryObject obj1 in gh)
                {
                    Autodesk.Revit.DB.Solid solid = obj1 as Autodesk.Revit.DB.Solid;
                    if (null != solid)
                    {
                        area = area + GetAreaVerical(GetVerticalFace(solid));
                    }
                }
            }
            return area;
        }


        public static double GetAreaFormaLaje(Autodesk.Revit.DB.Element ele)
        {
            double area = 0;
            Autodesk.Revit.DB.Options opt = ele.Document.Application.Create.NewGeometryOptions();
            Autodesk.Revit.DB.GeometryElement ge12 = ele.get_Geometry(opt);

            foreach (Autodesk.Revit.DB.GeometryObject obj in ge12)
            {
                Autodesk.Revit.DB.GeometryInstance gi = obj as Autodesk.Revit.DB.GeometryInstance;

                Autodesk.Revit.DB.GeometryElement gh = gi.GetInstanceGeometry();

                foreach (Autodesk.Revit.DB.GeometryObject obj1 in gh)
                {
                    Autodesk.Revit.DB.Solid solid = obj1 as Autodesk.Revit.DB.Solid;
                    if (null != solid)
                    {
                        try
                        {

                            if (null != GetBottonFace(solid))
                                area = area + GetBottonFace(solid).Area;
                        }
                        catch (Exception e)
                        {
                            InfoMsg("GetBottonFace" + "\n" + e.Message);
                        }
                    }
                }
            }
            return area;
        }
        public static double GetElevacaoRelacaoAoZeroToFace(Autodesk.Revit.DB.Element ele)
        {
            double elevacao = 0;
            foreach (Autodesk.Revit.DB.Solid solid in Util.GetSolids(ele))
            {

                if (ele is Autodesk.Revit.DB.Plumbing.Pipe)
                {
                    Autodesk.Revit.DB.LocationCurve lc = (ele as Autodesk.Revit.DB.Plumbing.Pipe).Location as Autodesk.Revit.DB.LocationCurve;

                }
                else
                {
                    try
                    {
                        if (null != solid)
                        {

                            Autodesk.Revit.DB.PlanarFace pf = Util.GetTopFace(solid);
                            foreach (Autodesk.Revit.DB.CurveLoop cl in pf.GetEdgesAsCurveLoops())
                            {
                                Autodesk.Revit.DB.CurveLoopIterator cli = cl.GetCurveLoopIterator();
                                Autodesk.Revit.DB.Line l = (cli.Current as Autodesk.Revit.DB.Line);
                                elevacao = l.GetEndPoint(0).Z * 0.3048;
                                continue;
                            }
                        }
                    }
                    catch

                    {
                    }
                }
            }
            return elevacao;
        }


        public static double GetElevacaoRelacaoAoZero(Autodesk.Revit.DB.Element ele)
        {
            double elevacao = 0;
            foreach (Autodesk.Revit.DB.Solid solid in Util.GetSolids(ele))
            {

                if (ele is Autodesk.Revit.DB.Plumbing.Pipe)
                {
                    Autodesk.Revit.DB.LocationCurve lc = (ele as Autodesk.Revit.DB.Plumbing.Pipe).Location as Autodesk.Revit.DB.LocationCurve;

                }
                else
                {
                    try
                    {
                        if (null != solid)
                        {

                            Autodesk.Revit.DB.PlanarFace pf = Util.GetBottonFace(solid);
                            foreach (Autodesk.Revit.DB.CurveLoop cl in pf.GetEdgesAsCurveLoops())
                            {
                                Autodesk.Revit.DB.CurveLoopIterator cli = cl.GetCurveLoopIterator();
                                Autodesk.Revit.DB.Line l = (cli.Current as Autodesk.Revit.DB.Line);
                                elevacao = l.GetEndPoint(0).Z * 0.3048;
                                continue;
                            }
                        }
                    }
                    catch

                    {
                    }
                }
            }
            return elevacao;
        }



        

        public static double GetElevacaoRelacaoAoZeroTopFace(Autodesk.Revit.DB.Element ele)
        {
            double elevacao = 0;
            foreach (Autodesk.Revit.DB.Solid solid in Util.GetSolids(ele))
            {

                if (ele is Autodesk.Revit.DB.Plumbing.Pipe)
                {
                    Autodesk.Revit.DB.LocationCurve lc = (ele as Autodesk.Revit.DB.Plumbing.Pipe).Location as Autodesk.Revit.DB.LocationCurve;

                }
                else
                {
                    try
                    {
                        if (null != solid)
                        {

                            Autodesk.Revit.DB.PlanarFace pf = Util.GetTopFace(solid);
                            foreach (Autodesk.Revit.DB.CurveLoop cl in pf.GetEdgesAsCurveLoops())
                            {
                                Autodesk.Revit.DB.CurveLoopIterator cli = cl.GetCurveLoopIterator();
                                Autodesk.Revit.DB.Line l = (cli.Current as Autodesk.Revit.DB.Line);
                                elevacao = l.GetEndPoint(0).Z * 0.3048;
                                continue;
                            }
                        }
                    }
                    catch

                    {
                    }
                }
            }
            return elevacao;
        }
        public static double GetElevacaoRelacaoAoZeroFaceFeet(Autodesk.Revit.DB.Face face )
        {
            double elevacao = 0;
            try
            {
                foreach (Autodesk.Revit.DB.CurveLoop cl in face.GetEdgesAsCurveLoops())
                {
                    Autodesk.Revit.DB.CurveLoopIterator cli = cl.GetCurveLoopIterator();
                    Autodesk.Revit.DB.Line l = (cli.Current as Autodesk.Revit.DB.Line);
                    elevacao = l.GetEndPoint(0).Z;
                    continue;
                }
            }
            catch
            {

            }
            return elevacao;
        }
        public static int GetLevelMaisProximo(Autodesk.Revit.DB.Element ele, DataTable dt)

        {
            double menorLevel = -1;
            double d = Util.GetElevacaoRelacaoAoZero(ele);
            int medicaoBlocoId = 0;


            foreach (DataRow dr in dt.Rows)
            {


                if ((null == -1) ||
                    (Math.Abs(d - menorLevel * 0.3048) > Math.Abs(d - Convert.ToDouble(dr["ELEVACAO"]) * 0.3048)))
                {
                    menorLevel = Convert.ToDouble(dr["ELEVACAO"]);
                    medicaoBlocoId = Convert.ToInt32(dr["medicao_bloco_id"]);
                }
            }
            return medicaoBlocoId;
        }
        public static Autodesk.Revit.DB.ElementId GetLevelMaisProximo(Autodesk.Revit.DB.Face face, List<Autodesk.Revit.DB.Level> lista)

        {
            Autodesk.Revit.DB.Level i = lista.OrderBy(x => x.Elevation).FirstOrDefault() as Autodesk.Revit.DB.Level;
            double menorLevel = i.Elevation;
            double d = Util.GetElevacaoRelacaoAoZeroFaceFeet(face);
            Autodesk.Revit.DB.ElementId eleId = i.Id;
            foreach (Autodesk.Revit.DB.Level level in lista)
            {


                if (Math.Abs(d - menorLevel) > Math.Abs(d - level.Elevation))
                {
                    menorLevel = level.Elevation;
                    eleId = level.Id;
                }
            }
            return eleId;
        }

        public static Autodesk.Revit.DB.ElementId GetLevelMaisProximo(Autodesk.Revit.DB.XYZ p, List<Autodesk.Revit.DB.Level> lista)

        {
            
            double d = p.Z;
            double menorLevel = lista.First().Elevation;
            Autodesk.Revit.DB.ElementId eleId = lista.First().Id;

            foreach (Autodesk.Revit.DB.Level level in lista)
            {


                if ((null == -1) ||
                    (Math.Abs(d - menorLevel) > Math.Abs(d - level.Elevation)))
                {
                    menorLevel = level.Elevation;
                    eleId = level.Id;
                }
            }
            return eleId;
        }

        public static Autodesk.Revit.DB.Level GetNivelMaisProximo(Autodesk.Revit.DB.Element ele, Autodesk.Revit.DB.FilteredElementCollector filtro)

        {
            double menorLevel = -1;
            double d = Util.GetElevacaoRelacaoAoZero(ele);

            foreach (Autodesk.Revit.DB.Level level in filtro)
            {
                if ((null == -1) ||
                    (Math.Abs(d - menorLevel * 0.3048) > Math.Abs(d - level.Elevation * 0.3048)))
                {
                    menorLevel = level.Elevation;
                    medicaoBlocoId = level;
                }
            }
            return medicaoBlocoId;
        }
        public static Autodesk.Revit.DB.Level GetNivelMaisProximoToFace(Autodesk.Revit.DB.Element ele, Autodesk.Revit.DB.FilteredElementCollector filtro)

        {
            double menorLevel = -1;
            double d = Util.GetElevacaoRelacaoAoZeroToFace(ele);

            foreach (Autodesk.Revit.DB.Level level in filtro)
            {
                if ((null == -1) ||
                    (Math.Abs(d - menorLevel * 0.3048) > Math.Abs(d - level.Elevation * 0.3048)))
                {
                    menorLevel = level.Elevation;
                    medicaoBlocoId = level;
                }
            }
            return medicaoBlocoId;
        }

        public static Autodesk.Revit.DB.Color GetColorRevit(wDrawing.Color colorWindows)
        {

            return new Autodesk.Revit.DB.Color(colorWindows.R, colorWindows.G, colorWindows.B);
        }

        public static List<Autodesk.Revit.DB.Solid> GetSolids(Autodesk.Revit.DB.Element ele)
        {
            List<Autodesk.Revit.DB.Solid> lista = new List<Autodesk.Revit.DB.Solid>();
            Autodesk.Revit.DB.Options opt = ele.Document.Application.Create.NewGeometryOptions();
            opt.ComputeReferences = true;
            opt.DetailLevel = Autodesk.Revit.DB.ViewDetailLevel.Fine;
            opt.IncludeNonVisibleObjects = true;
            // opt.View = ele.Autodesk.Revit.DB.Document.ActiveView;

            Autodesk.Revit.DB.GeometryElement ge12 = ele.get_Geometry(opt);

            foreach (Autodesk.Revit.DB.GeometryObject obj in ge12)
            {
                try
                {

                    if (obj is Autodesk.Revit.DB.Solid)
                    {
                        lista.Add((obj as Autodesk.Revit.DB.Solid));
                    }
                    if (obj is Autodesk.Revit.DB.GeometryInstance)
                    {

                        Autodesk.Revit.DB.GeometryInstance gi = obj as Autodesk.Revit.DB.GeometryInstance;
                        Autodesk.Revit.DB.GeometryElement gh = gi.GetInstanceGeometry();
                        foreach (Autodesk.Revit.DB.GeometryObject obj1 in gh)
                        {
                            Autodesk.Revit.DB.Solid solid = obj1 as Autodesk.Revit.DB.Solid;
                            if (null != solid)
                            {
                                lista.Add(solid);

                            }
                        }
                    }
                }
                catch
                {

                }
            }

            return lista;
        }
        public static double GetAreaFormaEstrutura(Autodesk.Revit.DB.Element ele1, string structuralColumn, string column,
                                             string structuralFraming, string floors)
        {
            area = 0;
            try
            {

                if (ele1.Category.Name == structuralColumn)
                {
                    area = GetAreaFormaPilar(ele1);
                    //InfoMsg("Categoria: " + ele1.Category.Name+"\n"+area.ToString());
                }
                if (ele1.Category.Name == column)
                {
                    area = GetAreaFormaPilar(ele1);
                    //InfoMsg("Categoria: " + ele1.Category.Name + "\n" + area.ToString());
                }
                if (ele1.Category.Name == structuralFraming)
                {
                    area = GetAreaFormaViga(ele1) + GetAreaFormaLaje(ele1);
                    //InfoMsg("Categoria: " + ele1.Category.Name + "\n" + area.ToString());
                }
                if (ele1.Category.Name == floors)
                {
                    area = GetAreaFormaLaje(ele1);
                    //InfoMsg("Categoria: " + ele1.Category.Name + "\n" + area.ToString());
                }
            }
            catch (Exception e)
            {

            }


            return area;
        }

        public static double AreaFacesFormaViga(Autodesk.Revit.DB.FaceArray faceArray)
        {
            double area = 0;
            List<double> lista = new List<double>();

            foreach (Autodesk.Revit.DB.PlanarFace pf in faceArray)
            {
                lista.Add(pf.Area);

            }
            lista.Sort();

            if (lista.Count > 0)
            {

                area = area + lista[lista.Count - 1];
                area = area + lista[lista.Count - 2];
            }


            return area;
        }

        public static Autodesk.Revit.DB.Face GetSecaoTransversalViga(Autodesk.Revit.DB.FaceArray faceArray)
        {
            Autodesk.Revit.DB.Face secaoTransversalMenor = null;
            foreach (Autodesk.Revit.DB.PlanarFace pf in faceArray)
            {
                if (null != pf && Util.IsVertical(pf))
                {
                    if (secaoTransversalMenor == null)
                    {
                        secaoTransversalMenor = pf;
                    }
                    else
                    {
                        if (secaoTransversalMenor.Area > pf.Area)
                        {
                            secaoTransversalMenor = pf;
                        }
                    }
                }
            }
            return secaoTransversalMenor;
        }

       /* public static Line GetMenorDimensao(Autodesk.Revit.DB.Face face)
        {
            Line menorDimensao = null;
            foreach (Autodesk.Revit.DB.CurveLoop curveLoop in face.GetEdgesAsCurveLoops())
            {

                CurveLoopIterator cli = curveLoop.GetCurveLoopIterator();
                while (cli.MoveNext())
                {
                    if (menorDimensao == null)
                    {
                        menorDimensao = (cli.Current as Line);
                    }
                    else
                    {
                        if (menorDimensao.Length > (cli.Current as Line).Length)
                            menorDimensao = (cli.Current as Line);
                    }

                }
            }

            return menorDimensao;

        }*/
        public static Autodesk.Revit.DB.Line GetLinhaX(Autodesk.Revit.DB.Face face)
        {
            foreach (Autodesk.Revit.DB.CurveLoop curveLoop in face.GetEdgesAsCurveLoops())
            {
                Autodesk.Revit.DB.CurveLoopIterator cli = curveLoop.GetCurveLoopIterator();
                while (cli.MoveNext())
                {
                    double m = 0;
                    Autodesk.Revit.DB.Line linhaAtual = (cli.Current as Autodesk.Revit.DB.Line);
                    Autodesk.Revit.DB.XYZ p1 = linhaAtual.GetEndPoint(0);
                    Autodesk.Revit.DB.XYZ p2 = linhaAtual.GetEndPoint(1);
                    if ((p1.X - p2.X) > 0)
                    {
                        m = (p1.Y - p2.Y) / (p1.X - p2.X);
                        if (m > 0) return linhaAtual;
                    }
                    if (p1.Y == p2.Y)
                    {
                        return linhaAtual;
                    }
                }
            }
            return null;

        }

        public static Autodesk.Revit.DB.Line GetLinhaY(Autodesk.Revit.DB.Face face)
        {
            foreach (Autodesk.Revit.DB.CurveLoop curveLoop in face.GetEdgesAsCurveLoops())
            {
                Autodesk.Revit.DB.CurveLoopIterator cli = curveLoop.GetCurveLoopIterator();
                while (cli.MoveNext())
                {
                    double m = 0;
                    Autodesk.Revit.DB.Line linhaAtual = (cli.Current as Autodesk.Revit.DB.Line);
                    Autodesk.Revit.DB.XYZ p1 = linhaAtual.GetEndPoint(0);
                    Autodesk.Revit.DB.XYZ p2 = linhaAtual.GetEndPoint(1);
                    if ((p1.X - p2.X) > 0)
                    {
                        m = (p1.Y - p2.Y) / (p1.X - p2.X);
                        if (m < 0) return linhaAtual;
                    }
                    if (p1.X == p2.X)
                    {
                        return linhaAtual;
                    }
                }
            }

            return null;

        }
        public static Autodesk.Revit.DB.Line GetMaiorDimensao(Autodesk.Revit.DB.Face face)
        {
            Autodesk.Revit.DB.Line maiorDimensao = null;
            foreach (Autodesk.Revit.DB.CurveLoop curveLoop in face.GetEdgesAsCurveLoops())
            {

                Autodesk.Revit.DB.CurveLoopIterator cli = curveLoop.GetCurveLoopIterator();
                while (cli.MoveNext())
                {
                    if (maiorDimensao == null)
                    {
                        maiorDimensao = (cli.Current as Autodesk.Revit.DB.Line);
                    }
                    else
                    {
                        if (maiorDimensao.Length < (cli.Current as Autodesk.Revit.DB.Line).Length)
                            maiorDimensao = (cli.Current as Autodesk.Revit.DB.Line);
                    }

                }
            }

            return maiorDimensao;

        }


        public static double GetAreaFormaViga(Autodesk.Revit.DB.Element ele)
        {

            double area = 0;
            Autodesk.Revit.DB.Options opt = ele.Document.Application.Create.NewGeometryOptions();
            Autodesk.Revit.DB.GeometryElement ge12 = ele.get_Geometry(opt);

            foreach (Autodesk.Revit.DB.GeometryObject obj in ge12)
            {
                Autodesk.Revit.DB.GeometryInstance gi = obj as Autodesk.Revit.DB.GeometryInstance;

                Autodesk.Revit.DB.GeometryElement gh = gi.GetInstanceGeometry();

                foreach (Autodesk.Revit.DB.GeometryObject obj1 in gh)
                {
                    Autodesk.Revit.DB.Solid solid = obj1 as Autodesk.Revit.DB.Solid;
                    if (null != solid)
                    {
                        area = area + AreaFacesFormaViga(GetVerticalFace(solid));
                    }
                }
            }
            return area;
        }
        static public bool IsZero(double a)
        {
            return _eps > Math.Abs(a);
        }

        static public bool IsHorizontal(Autodesk.Revit.DB.XYZ v)
        {
            return IsZero(v.Z);
        }

        static public bool IsVertical(Autodesk.Revit.DB.XYZ v)
        {
            return IsZero(v.X) && IsZero(v.Y);
        }

        static public bool IsHorizontal(Autodesk.Revit.DB.Edge e)
        {
            Autodesk.Revit.DB.XYZ p = e.Evaluate(0);
            Autodesk.Revit.DB.XYZ q = e.Evaluate(1);
            return IsHorizontal(q - p);
        }

        static public bool IsHorizontal(Autodesk.Revit.DB.PlanarFace f)
        {
            return IsVertical(f.FaceNormal);
        }

        static public bool IsVertical(Autodesk.Revit.DB.PlanarFace f)
        {
            return IsHorizontal(f.FaceNormal);
        }

        static public bool IsVertical(Autodesk.Revit.DB.CylindricalFace f)
        {
            return IsVertical(f.Axis);
        }
        #endregion // Geometrical Comparison

        #region Formatting
        /// <summary>
        /// Return an English plural suffix 's' or 
        /// nothing for the given number of items.
        /// </summary>


        static public string AngleString(double a)
        {
            return RealString(a * 180 / Math.PI) + " degrees";
        }

        static public string PointString(Autodesk.Revit.DB.XYZ p)
        {
            return string.Format("({0},{1},{2})",
              RealString(p.X), RealString(p.Y),
              RealString(p.Z));
        }

        static public string TransformString(Autodesk.Revit.DB.Transform t)
        {
            return string.Format("({0},{1},{2},{3})", PointString(t.Origin),
              PointString(t.BasisX), PointString(t.BasisY), PointString(t.BasisZ));
        }
        #endregion // Formatting

        const string _caption = "The Building Coder";

        static public void InfoMsg(string msg)
        {
            //Debug.WriteLine( msg );
            System.Windows.Forms.MessageBox.Show(msg,
              _caption,
              System.Windows.Forms.MessageBoxButtons.OK,
              System.Windows.Forms.MessageBoxIcon.Information);
        }

        public static string ElementDescription(Autodesk.Revit.DB.Element e)
        {
            // for a wall, the Autodesk.Revit.DB.Element name equals the 
            // wall type name, which is equivalent to the 
            // family name ...
            Autodesk.Revit.DB.FamilyInstance fi = e as Autodesk.Revit.DB.FamilyInstance;
            string fn = (null == fi)
              ? string.Empty
              : fi.Symbol.Family.Name + " ";

            string cn = (null == e.Category)
              ? e.GetType().Name
              : e.Category.Name;

            return string.Format("{0} {1}<{2} {3}>",
              cn, fn, e.Id.IntegerValue, e.Name);
        }

        public static Autodesk.Revit.DB.Element SelectSingleElement(Autodesk.Revit.DB.Document doc, string description)
        {
            //    UIApplication UIAPP = doc.Application as UIApplication;
            // Selection sel = UIAPP.ActiveUIDocument.Selection;
            Autodesk.Revit.DB.Element e = null;
            //sel.Elements.Clear();
            //sel.StatusbarTip = "Please select " + description;
            /* if( sel.PickObject() )
             {
               ElementSetIterator elemSetItr
                 = sel.Elements.ForwardIterator();
               elemSetItr.MoveNext();
               e = elemSetItr.Current as Autodesk.Revit.DB.Element;
             }*/
            return e;
        }

        public static bool GetSelectedElementsOrAll(
          List<Autodesk.Revit.DB.Element> a,
          Autodesk.Revit.DB.Document doc,
          Type t)
        {
            // Selection sel = (doc.Application as UIApplication).ActiveUIDocument.Selection;
            /* if( 0 < sel.Elements.Size )
             {
               foreach( Autodesk.Revit.DB.Element e in sel.Elements )
               {
                 if( e.GetType().IsSubclassOf( t ) )
                 {
                   a.Add( e );
                 }
               }
             }
             else
             {
              // doc.get_Elements( t, a );
             }*/
            return 0 < a.Count;
        }

        /// <summary>
        /// Return a string for a real number
        /// formatted to two decimal places.
        /// </summary>



        public static TaskDialogResult PerguntarRevit()
        {

            // Creates a Revit task dialog to communicate information to the user.
            TaskDialog mainDialog = new TaskDialog("Hello, Revit!");
            mainDialog.MainInstruction = "Escolha uma opção";
            mainDialog.MainContent = "As opções abaixo determinam como os elementos serão deletados";

            // Add commmandLink options to task dialog
            mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1,
                     "Deletar elementos no AutoOrc");
            mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2,
                            "Não deletar elementos no AutoOrc");
            mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3,
                               "Cancelar");


            // Set common buttons and default button. If no CommonButton or CommandLink is added,
            // task dialog will show a Close button by default
            mainDialog.CommonButtons = TaskDialogCommonButtons.Close;
            mainDialog.DefaultButton = TaskDialogResult.Close;

            // Set footer text. Footer text is usually used to link to the help Autodesk.Revit.DB.Document.
            mainDialog.FooterText = "Selecione uma opção";

            TaskDialogResult tResult = mainDialog.Show();
            return tResult;
        }

        public static Autodesk.Revit.DB.Element GetInsulationPipeOuflexPipe(Autodesk.Revit.DB.Document doc, string name)
        {
            Util.uiDoc = uiDoc;

            return Util.FindElementByName(typeof(Autodesk.Revit.DB.Plumbing.PipeInsulationType), name);
        }
        public static Autodesk.Revit.DB.Element GetInsulationDuct(string name)
        {
            Util.uiDoc = uiDoc;

            return Util.FindElementByName(typeof(Autodesk.Revit.DB.Mechanical.DuctInsulationType), name);
        }

        /*public static void RetirarInsulation(Autodesk.Revit.DB.Element ele)
        {
            if (ele is Autodesk.Revit.DB.Plumbing.Pipe)
            {
                Autodesk.Revit.DB.FilteredElementCollector fec = new Autodesk.Revit.DB.FilteredElementCollector(ele.Autodesk.Revit.DB.Document).OfClass(typeof(Autodesk.Revit.DB.Plumbing.PipeInsulation));
                foreach (Autodesk.Revit.DB.Plumbing.PipeInsulation pi in fec)
                {
                    if (pi.HostElementId == ele.Id)
                    {
                        ele.Document.Delete(pi.Id);

                        break;
                    }
                }
            }
            if (ele is Autodesk.Revit.DB.Plumbing.FlexPipe)
            {
                Autodesk.Revit.DB.FilteredElementCollector fec = new Autodesk.Revit.DB.FilteredElementCollector(ele.Document.OfClass(typeof(Autodesk.Revit.DB.Plumbing.PipeInsulation));
                foreach (Autodesk.Revit.DB.Plumbing.PipeInsulation pi in fec)
                {
                    if (pi.HostElementId == ele.Id)
                    {
                        ele.Document.Delete(pi.Id);
                        break;
                    }
                }
            }
            if (ele is Autodesk.Revit.DB.Mechanical.Duct)
            {
                Autodesk.Revit.DB.FilteredElementCollector fec = new Autodesk.Revit.DB.FilteredElementCollector(ele.Document.OfClass(typeof(Autodesk.Revit.DB.Mechanical.DuctInsulation));
                foreach (Autodesk.Revit.DB.Mechanical.DuctInsulation pi in fec)
                {
                    if (pi.HostElementId == ele.Id)
                    {
                        ele.Document.Delete(pi.Id);
                        break;
                    }
                }
            }

        }
        */
        public static string RealString(double a)
        {
            return a.ToString("0.##");
        }

        /// <summary>
        /// Return a string for an Autodesk.Revit.DB.XYZ point
        /// or vector with its coordinates
        /// formatted to two decimal places.
        /// </summary>

        #endregion // Formatting

        #region Display a message
        public const string Caption = "Family API";



        public static void InfoMsg2(
          string instruction,
          string content)
        {

            TaskDialog d = new TaskDialog(Caption);
            d.MainInstruction = instruction;
            d.MainContent = content;
            d.Show();
        }

        public static void ErrorMsg(string msg)
        {

            TaskDialog d = new TaskDialog(Caption);
            d.MainIcon = TaskDialogIcon.TaskDialogIconWarning;
            d.MainInstruction = msg;
            d.Show();
        }

        /// <summary>
        /// Return a string describing the given Autodesk.Revit.DB.Element:
        /// .NET type name,
        /// category name,
        /// family and symbol name for a family instance,
        /// Autodesk.Revit.DB.Element id and Autodesk.Revit.DB.Element name.
        /// </summary>
        public static string ElementDescription1(
          Autodesk.Revit.DB.Element e)
        {
            if (null == e)
            {
                return "<null>";
            }

            // For a wall, the Autodesk.Revit.DB.Element name equals the
            // wall type name, which is equivalent to the
            // family name ...

            Autodesk.Revit.DB.FamilyInstance fi = e as Autodesk.Revit.DB.FamilyInstance;

            string typeName = e.GetType().Name;

            string categoryName = (null == e.Category)
              ? string.Empty
              : e.Category.Name + " ";

            string familyName = (null == fi)
              ? string.Empty
              : fi.Symbol.Family.Name + " ";

            string symbolName = (null == fi
              || e.Name.Equals(fi.Symbol.Name))
                ? string.Empty
                : fi.Symbol.Name + " ";

            return string.Format("{0} {1}{2}{3}<{4} {5}>",
              typeName, categoryName, familyName,
              symbolName, e.Id.IntegerValue, e.Name);
        }
        #endregion // Display a message

        public static List<Autodesk.Revit.DB.Family> FindFamilyByLevel(
          Autodesk.Revit.DB.Document doc, Autodesk.Revit.DB.Level level)
        {

            return null;
            /* new Autodesk.Revit.DB.FilteredElementCollector(doc)
               .OfClass(targetType)
               .FirstOrDefault<Autodesk.Revit.DB.Element>(
                 e => e.Name.Equals(targetName));*/
        }

        public static void DadosDeConexoeseComprimentos(Autodesk.Revit.DB.Element element)
        {
            Autodesk.Revit.DB.FamilyInstance aFamilyInst;
            var qtde = element.LookupParameter("Qtde");
            var unidade = element.LookupParameter("Unidade");

            if (element.Category.Name == "Conexões de tubo")
            {
                if ((element as Autodesk.Revit.DB.FamilyInstance).Symbol.Family.Name.Contains("PEXTubo"))
                {
                    try
                    {
                        var compCurva = element.LookupParameter("Length").AsDouble() * 0.3040;
                        unidade.Set("M");
                        qtde.Set(compCurva);
                    }
                    catch
                    {
                        unidade.Set("M");
                        qtde.Set(-1);
                    }
                    return;

                }
                else
                {
                    unidade.Set("Unid");
                    qtde.Set(1);
                }
            }
            else if (CriarMenu.ListaDeCategoriasDeItensContaveis.Contains(element.Category.Name))
            {
                /*if (CriarMenu.ListaQueTemTipoDeSistema.Contains(Autodesk.Revit.DB.Element.Category.Name))
                {
                    aFamilyInst = Autodesk.Revit.DB.Element as Autodesk.Revit.DB.FamilyInstance;
                    /*if (aFamilyInst.SuperComponent != null)
                    {
                        var p = aFamilyInst.SuperComponent.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);
                        if (p != null)
                            if (p.HasValue)
                                aFamilyInst.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).Set(p.AsElementId());
                    }
                }*/
                unidade.Set("Unid");
                qtde.Set(1);
            }
            else if (CriarMenu.ListaCategoriaItensSistemaPorMetro.Contains(element.Category.Name))
            {
                unidade.Set("m");
                qtde.Set(element.LookupParameter("Comprimento").AsDouble() * 0.3040);
            }
        }

        #region Retrieve an Autodesk.Revit.DB.Element by class and name
        /// <summary>
        /// Retrieve a database Autodesk.Revit.DB.Element 
        /// of the given type and name.
        /// </summary>
        public static Autodesk.Revit.DB.Element FindElementByName(
          Type targetType,
          string targetName)
        {
            return new Autodesk.Revit.DB.FilteredElementCollector(uiDoc)
              .OfClass(targetType)
              .FirstOrDefault<Autodesk.Revit.DB.Element>(
                e => e.Name.Equals(targetName));
        }
        #endregion // Retrieve an Autodesk.Revit.DB.Element by class and name





        static public Autodesk.Revit.DB.IndependentTag CreateIndependentTag(Autodesk.Revit.DB.Document document, Autodesk.Revit.DB.Wall wall)
        {
            // make sure active view is not a 3D view
            Autodesk.Revit.DB.View view = document.ActiveView;
            // define tag mode and tag orientation for new tag
            Autodesk.Revit.DB.TagMode tagMode = Autodesk.Revit.DB.TagMode.TM_ADDBY_CATEGORY;
            Autodesk.Revit.DB.TagOrientation tagorn = Autodesk.Revit.DB.TagOrientation.Horizontal;

            // Add the tag to the middle of the wall
            Autodesk.Revit.DB.LocationCurve wallLoc = wall.Location as Autodesk.Revit.DB.LocationCurve;
            Autodesk.Revit.DB.XYZ wallStart = wallLoc.Curve.GetEndPoint(0);
            Autodesk.Revit.DB.XYZ wallEnd = wallLoc.Curve.GetEndPoint(1);
            Autodesk.Revit.DB.XYZ wallMid = wallLoc.Curve.Evaluate(0.5, true);

  //          IndependentTag newTag = Autodesk.Revit.DB.Document.Create.NewTag(view, wall, true, tagMode, tagorn, wallMid);

        //    if (null == newTag)
            {
                throw new Exception("Create IndependentTag Failed.");
            }

            // newTag.TagText is read-only, so we change the Type Mark type parameter to
            // set the tag text. The label parameter for the tag family determines
            // what type parameter is used for the tag text.

            Autodesk.Revit.DB.WallType type = wall.WallType;

            Autodesk.Revit.DB.Parameter foundParameter = type.LookupParameter("");
            bool result = foundParameter.Set("Hello");

            // set leader mode free
            // otherwise leader end point move with elbow point

            /* newTag.LeaderEndCondition = LeaderEndCondition.Free;
             Autodesk.Revit.DB.XYZ elbowPnt = wallMid + new Autodesk.Revit.DB.XYZ(5.0, 5.0, 0.0);
             newTag.LeaderElbow = elbowPnt;
             Autodesk.Revit.DB.XYZ headerPnt = wallMid + new Autodesk.Revit.DB.XYZ(10.0, 10.0, 0.0);
             newTag.TagHeadPosition = headerPnt;

             return newTag;*/
            return null;
        }

        static public void CriarMaterial(Autodesk.Revit.DB.Document uiDoc, string nome, byte r, byte g, byte b)
        {

            Autodesk.Revit.DB.Material material = Util.FindElementByName(typeof(Autodesk.Revit.DB.Material), nome) as Autodesk.Revit.DB.Material;
            if(material != null)
            {
                Autodesk.Revit.DB.Material Previsto = Util.FindElementByName(typeof(Autodesk.Revit.DB.Material), nome) as Autodesk.Revit.DB.Material;
                Autodesk.Revit.DB.Color cor = new Autodesk.Revit.DB.Color(r, g, b);
                Previsto.Color = cor;
            #if DEBUG20201 || DEBGUG20211
                            Previsto.CutForegroundPatternColor = cor;
                            Previsto.SurfaceForegroundPatternColor = cor;
            #else
                            //    Previsto.SurfacePatternColor = cor;
                            //  Previsto.CutPatternColor = cor;
            #endif
                            cor.Dispose();

            }
            else
            {
                Autodesk.Revit.DB.Material defaultm = Util.FindElementByName(typeof(Autodesk.Revit.DB.Material), "Default") as Autodesk.Revit.DB.Material;
                Autodesk.Revit.DB.Material Previsto = defaultm.Duplicate(nome);
                Autodesk.Revit.DB.Color cor = new Autodesk.Revit.DB.Color(r, g, b);
                Previsto.Color = cor;

                //Autodesk.Revit.DB.Element preenchimento = Util.FindElementByName(typeof(Autodesk.Revit.DB.Element), "Preenchimento sólido") as Autodesk.Revit.DB.Element;
#if  DEBUG20201 || DEBGUG20211
                Previsto.CutForegroundPatternColor = cor;
                Previsto.SurfaceForegroundPatternColor = cor;
#else
                //// Previsto.SurfacePatternId = preenchimento.Id;
                //   Previsto.SurfacePatternColor = cor;
                //    Previsto.CutPatternColor = cor;
#endif


            }

        }


            static public void CriarMaterialFuncao(Autodesk.Revit.DB.Document uiDoc)
        {


            try
            {

                Autodesk.Revit.DB.Material defaultm = Util.FindElementByName(typeof(Autodesk.Revit.DB.Material), "Default") as Autodesk.Revit.DB.Material;
                Autodesk.Revit.DB.Material Previsto = defaultm.Duplicate("Previsto");
                Autodesk.Revit.DB.Color cor = new Autodesk.Revit.DB.Color(255, 128, 0);
                Previsto.Color = cor;
              
                Autodesk.Revit.DB.Element preenchimento = Util.FindElementByName(typeof(Autodesk.Revit.DB.Element), "Preenchimento sólido") as Autodesk.Revit.DB.Element;
#if DEBUG20201
#else
               //// Previsto.SurfacePatternId = preenchimento.Id;
             //   Previsto.SurfacePatternColor = cor;
            //    Previsto.CutPatternColor = cor;
#endif

                cor.Dispose();
            }
            catch
            {
                Autodesk.Revit.DB.Material Previsto = Util.FindElementByName(typeof(Autodesk.Revit.DB.Material), "Previsto") as Autodesk.Revit.DB.Material;
                Autodesk.Revit.DB.Color cor = new Autodesk.Revit.DB.Color(255, 128, 0);
                Previsto.Color = cor;
          #if DEBUG20201
#else
            //    Previsto.SurfacePatternColor = cor;
              //  Previsto.CutPatternColor = cor;
#endif
                cor.Dispose();
            }
            try
            {
                Autodesk.Revit.DB.Material previsto = Util.FindElementByName(typeof(Autodesk.Revit.DB.Material), "Default") as Autodesk.Revit.DB.Material;
                Autodesk.Revit.DB.Material Completo = previsto.Duplicate("Completo");
                Autodesk.Revit.DB.Color cor = new Autodesk.Revit.DB.Color(0, 0, 255);
                Completo.Color = cor;
                //Completo.SurfacePatternColor = cor;
                //Completo.CutPatternColor = cor;
                cor.Dispose();
            }
            catch
            {

            }







            try
            {
                Autodesk.Revit.DB.Material previsto = Util.FindElementByName(typeof(Autodesk.Revit.DB.Material), "Default") as Autodesk.Revit.DB.Material;
                Autodesk.Revit.DB.Material Completo = previsto.Duplicate("Concluido");
                Autodesk.Revit.DB.Color cor = new Autodesk.Revit.DB.Color(0, 0, 255);
                Completo.Color = cor;
               // Completo.SurfacePatternColor = cor;
               // Completo.CutPatternColor = cor;
                cor.Dispose();
            }
            catch
            {

            }
            try
            {
                Autodesk.Revit.DB.Material previsto = Util.FindElementByName(typeof(Autodesk.Revit.DB.Material), "Default") as Autodesk.Revit.DB.Material;
                Autodesk.Revit.DB.Material iniciado = previsto.Duplicate("Iniciado");
                Autodesk.Revit.DB.Color cor = new Autodesk.Revit.DB.Color(255, 255, 0);
                iniciado.Color = cor;
               // iniciado.CutPatternColor = cor;
              //  iniciado.SurfacePatternColor = cor;

                cor.Dispose();

            }
            catch
            {

            }
            try
            {
                Autodesk.Revit.DB.Material previsto = Util.FindElementByName(typeof(Autodesk.Revit.DB.Material), "Default") as Autodesk.Revit.DB.Material;
                Autodesk.Revit.DB.Material iniciado = previsto.Duplicate("Projecao");
                Autodesk.Revit.DB.Color cor = new Autodesk.Revit.DB.Color(0, 128, 0);
                iniciado.Color = cor;
              
               // iniciado.CutPatternColor = cor;
              //  iniciado.SurfacePatternColor = cor;

                cor.Dispose();

            }

            catch
            {

            }
            try
            {
                Autodesk.Revit.DB.Material previsto = Util.FindElementByName(typeof(Autodesk.Revit.DB.Material), "Default") as Autodesk.Revit.DB.Material;
                Autodesk.Revit.DB.Material iniciado = previsto.Duplicate("Projecao1");
                Autodesk.Revit.DB.Color cor = new Autodesk.Revit.DB.Color(255, 0, 255);
                iniciado.Color = cor;
#if DEBUG20201
#else
            //    iniciado.CutPatternColor = cor;
            //    iniciado.SurfacePatternColor = cor;
#endif
                cor.Dispose();

            }

            catch
            {

            }
            try
            {
                Autodesk.Revit.DB.Material previsto = Util.FindElementByName(typeof(Autodesk.Revit.DB.Material), "Default") as Autodesk.Revit.DB.Material;
                Autodesk.Revit.DB.Material iniciado = previsto.Duplicate("Projecao2");
                Autodesk.Revit.DB.Color cor = new Autodesk.Revit.DB.Color(0, 255, 255);
                iniciado.Color = cor;
               // iniciado.CutPatternColor = cor;
              //  iniciado.SurfacePatternColor = cor;

                cor.Dispose();

            }

            catch
            {

            }
            try
            {
                Autodesk.Revit.DB.Material previsto = Util.FindElementByName(typeof(Autodesk.Revit.DB.Material), "Default") as Autodesk.Revit.DB.Material;
                Autodesk.Revit.DB.Material iniciado = previsto.Duplicate("Projecao3");
                Autodesk.Revit.DB.Color cor = new Autodesk.Revit.DB.Color(128, 128, 0);
                iniciado.Color = cor;
#if DEBUG20201
#else
              //  i//niciado.CutPatternColor = cor;
              //  iniciado.SurfacePatternColor = cor;
#endif
                cor.Dispose();

            }

            catch
            {

            }

            try
            {
                Autodesk.Revit.DB.Material previsto = Util.FindElementByName(typeof(Autodesk.Revit.DB.Material), "Default") as Autodesk.Revit.DB.Material;
                Autodesk.Revit.DB.Material iniciado = previsto.Duplicate("Projecao4");
                Autodesk.Revit.DB.Color cor = new Autodesk.Revit.DB.Color(0, 128, 0);
                iniciado.Color = cor;
              //  iniciado.CutPatternColor = cor;
              //  iniciado.SurfacePatternColor = cor;
                cor.Dispose();

            }
            catch
            {

            }
            try
            {

                Autodesk.Revit.DB.Plumbing.PipeInsulationType insulacaoPrevisto =
                Util.FindElementByName(typeof(Autodesk.Revit.DB.Plumbing.PipeInsulationType), "previsto") as Autodesk.Revit.DB.Plumbing.PipeInsulationType;
                Autodesk.Revit.DB.Plumbing.PipeInsulationType insulacaoProjecao = insulacaoPrevisto.Duplicate("Projecao") as Autodesk.Revit.DB.Plumbing.PipeInsulationType;
                Autodesk.Revit.DB.Color cor = new Autodesk.Revit.DB.Color(128, 255, 0);
                insulacaoProjecao.LookupParameter("Material").Set((Util.FindElementByName(typeof(Autodesk.Revit.DB.Material), "projecao") as Autodesk.Revit.DB.Material).Id);
            }
            catch
            {

            }


            try
            {

                Autodesk.Revit.DB.Plumbing.PipeInsulationType insulacaoPrevisto =
                Util.FindElementByName(typeof(Autodesk.Revit.DB.Plumbing.PipeInsulationType), "previsto") as Autodesk.Revit.DB.Plumbing.PipeInsulationType;
                Autodesk.Revit.DB.Plumbing.PipeInsulationType insulacaoProjecao = insulacaoPrevisto.Duplicate("Completo") as Autodesk.Revit.DB.Plumbing.PipeInsulationType;
                Autodesk.Revit.DB.Color cor = new Autodesk.Revit.DB.Color(0, 0, 255);
                insulacaoProjecao.LookupParameter("Material").Set((Util.FindElementByName(typeof(Autodesk.Revit.DB.Material), "Completo") as Autodesk.Revit.DB.Material).Id);
            }
            catch
            {

            }
        }


        static public void AlterarLinhaReferenciaParede(Autodesk.Revit.DB.Element elemento, int valorNovo)
        {
            Autodesk.Revit.DB.Transaction trans = new Autodesk.Revit.DB.Transaction(elemento.Document);

            trans.Start("Lab");
            elemento.LookupParameter("").Set(valorNovo);
            trans.Commit();
            trans.Dispose();



        }

        [Obsolete]
        static public void CriarFiltros(UIDocument uiDoc)
        {
            //UIDocument uidoc = ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uiDoc.Document;
            Autodesk.Revit.DB.View view = doc.ActiveView;

            TaskDialog.Show("Boost Your BIM", "# of filters applied to this view = " + view.GetFilters().Count);

            // create list of categories that will for the filter
            IList<Autodesk.Revit.DB.ElementId> categories = new List<Autodesk.Revit.DB.ElementId>();
            categories.Add(new Autodesk.Revit.DB.ElementId(Autodesk.Revit.DB.BuiltInCategory.OST_Walls));

            // create a list of rules for the filter
            IList<Autodesk.Revit.DB.FilterRule> rules = new List<Autodesk.Revit.DB.FilterRule>();
            // This filter will have a single rule that the wall type width must be less than 1/2 foot
            Autodesk.Revit.DB.Parameter wallWidth = new Autodesk.Revit.DB.FilteredElementCollector(doc).OfClass(typeof(Autodesk.Revit.DB.WallType)).FirstElement().get_Parameter(Autodesk.Revit.DB.BuiltInParameter.WALL_ATTR_WIDTH_PARAM);
            rules.Add(Autodesk.Revit.DB.ParameterFilterRuleFactory.CreateLessRule(wallWidth.Id, 0.5, 0.001));

            Autodesk.Revit.DB.ParameterFilterElement filter = null;
            using (Autodesk.Revit.DB.Transaction t = new Autodesk.Revit.DB.Transaction(doc, "Create and Apply Filter"))
            {
                t.Start();
#if R2021
#else
                Autodesk.Revit.DB.ElementParameterFilter f = new Autodesk.Revit.DB.ElementParameterFilter(rules);
            
                filter = Autodesk.Revit.DB.ParameterFilterElement.Create(doc, "Thin Wall Filter", categories, f);
                
                view.AddFilter(filter.Id);
            #endif
                t.Commit();
            }

            string filterNames = "";
            foreach (Autodesk.Revit.DB.ElementId id in view.GetFilters())
            {
                filterNames += doc.GetElement(id).Name + "\n";
            }
            TaskDialog.Show("Boost Your BIM", "Filters applied to this view: " + filterNames);

            // Create a new OverrideGraphicSettings object and specify cut line color and cut line weight
            Autodesk.Revit.DB.OverrideGraphicSettings ogs = new Autodesk.Revit.DB.OverrideGraphicSettings();
            ogs.SetCutLineColor(new Autodesk.Revit.DB.Color(255, 0, 0));
            ogs.SetCutLineWeight(9);

            using (Autodesk.Revit.DB.Transaction t = new Autodesk.Revit.DB.Transaction(doc, "Set Override Appearance"))
            {
                t.Start();
                view.SetFilterOverrides(filter.Id, ogs);
                t.Commit();
            }
        }

        static public Autodesk.Revit.DB.FilteredElementCollector ObterLevels(Autodesk.Revit.DB.Document uiDoc)
        {

            Autodesk.Revit.DB.FilteredElementCollector selecao = new Autodesk.Revit.DB.FilteredElementCollector(uiDoc).OfClass(typeof(Autodesk.Revit.DB.Level));


            return selecao;
        }
        static public object ObterLevelsComNomeEId(Autodesk.Revit.DB.Document uiDoc)
        {

            var selecao = new Autodesk.Revit.DB.FilteredElementCollector(uiDoc).OfClass(typeof(Autodesk.Revit.DB.Level)).Cast<Autodesk.Revit.DB.Level>().ToList();
           var lista =  from Autodesk.Revit.DB.Level level in selecao
                           select new 
                           {
                               NomePavimento = level.Name,
                               Id = level.Id.IntegerValue
                           };
            return lista.ToList();
        }
        public static List<Autodesk.Revit.DB.Level> ListaDeNiveis(Autodesk.Revit.DB.Document iudoc, bool somenteQueDefinePavimento = false)
        {
            if(somenteQueDefinePavimento)
                return new Autodesk.Revit.DB.FilteredElementCollector(iudoc).OfClass(typeof(Autodesk.Revit.DB.Level)).Cast<Autodesk.Revit.DB.Level>().
                    Where(x=>x.LookupParameter("Andar da construção").AsInteger()==1).ToList<Autodesk.Revit.DB.Level>();
            else
            return new Autodesk.Revit.DB.FilteredElementCollector(iudoc).OfClass(typeof(Autodesk.Revit.DB.Level)).Cast<Autodesk.Revit.DB.Level>().ToList<Autodesk.Revit.DB.Level>();

        }
        static public Autodesk.Revit.DB.FilteredElementCollector ObterFamilyTypePorCategoria(Autodesk.Revit.DB.Document uiDoc)
        {

            Autodesk.Revit.DB.FilteredElementCollector selecao = new Autodesk.Revit.DB.FilteredElementCollector(uiDoc).OfClass(typeof(Autodesk.Revit.DB.FloorType));


            return selecao;
        }



        public static MaterialFunctionAssignment ParseEnum(string input)
        {
            // Tenta converter a string para enum
            if (Enum.TryParse(input, true, out MaterialFunctionAssignment result))
            {
                return result;
            }

            // Se a conversão para enum falhar, tenta converter para int e depois para enum
            if (int.TryParse(input, out int intValue))
            {
                if (Enum.IsDefined(typeof(MaterialFunctionAssignment), intValue))
                {
                    return (MaterialFunctionAssignment)intValue;
                }
            }

            throw new ArgumentException("Invalid input for enum conversion.");
        }

        public static FloorType DuplicarFloorType(FloorType tipoOriginal, string nomeDoTipo, List<Camada> camadaNova)
        {
            FloorType novoTipo  = Util.FindElementByName(typeof(Autodesk.Revit.DB.FloorType), nomeDoTipo) as Autodesk.Revit.DB.FloorType;
            if (novoTipo == null)  novoTipo = tipoOriginal.Duplicate(nomeDoTipo) as FloorType;
            CompoundStructure estruturaAtual = novoTipo.GetCompoundStructure();   
           var camadasDoTipo = estruturaAtual.GetLayers();
            camadasDoTipo.Clear();
            foreach (var camada in camadaNova)
            {
                Autodesk.Revit.DB.Material material = Util.FindElementByName(typeof(Autodesk.Revit.DB.Material), camada.Material) as Autodesk.Revit.DB.Material;
                var layer = new CompoundStructureLayer(camada.Espessura / 0.3048, ParseEnum(camada.TipoDeCamada), material.Id );
                
                camadasDoTipo.Add(layer);
            }
            estruturaAtual.SetLayers( camadasDoTipo );
            estruturaAtual.VariableLayerIndex = camadasDoTipo.Count-1;
            novoTipo.SetCompoundStructure(estruturaAtual);
            return novoTipo;

        }
        public static WallType DuplicarWallType(WallType tipoOriginal, string nomeDoTipo, List<Camada> camadaNova)
        {
            WallType novoTipo = Util.FindElementByName(typeof(Autodesk.Revit.DB.WallType), nomeDoTipo) as Autodesk.Revit.DB.WallType;
           if(novoTipo==null)  novoTipo = tipoOriginal.Duplicate(nomeDoTipo) as WallType;
            CompoundStructure estruturaAtual = novoTipo.GetCompoundStructure();
            var camadasDoTipo = estruturaAtual.GetLayers();
            camadasDoTipo.Clear();
            foreach (var camada in camadaNova)
            {
                Autodesk.Revit.DB.Material material = Util.FindElementByName(typeof(Autodesk.Revit.DB.Material), camada.Material) as Autodesk.Revit.DB.Material;
                var layer = new CompoundStructureLayer(camada.Espessura / 0.3048, ParseEnum(camada.TipoDeCamada), material.Id);

                camadasDoTipo.Add(layer);
            }
            estruturaAtual.SetLayers(camadasDoTipo);
            //estruturaAtual.VariableLayerIndex = camadasDoTipo.Count;
            novoTipo.SetCompoundStructure(estruturaAtual);
            return novoTipo;

        }
        static public void ObterMaterialLayerParede(Autodesk.Revit.DB.Document documento, Autodesk.Revit.DB.Element elemento)
        {

            GetFamilyTypesInFamily(documento);

            // Autodesk.Revit.DB.Element elemento;
            //elemento.
            Autodesk.Revit.DB.Wall wall = elemento as Autodesk.Revit.DB.Wall;
            Autodesk.Revit.DB.WallType aWallType = wall.WallType;
            //somente se aperede for básica

            if (Autodesk.Revit.DB.WallKind.Basic == aWallType.Kind)
            {
                //get CompoundStructure
                Autodesk.Revit.DB.CompoundStructure comStruct = aWallType.GetCompoundStructure();
                Autodesk.Revit.DB.Categories allCategories = documento.Settings.Categories;
                // Get the category OST_Walls default Material; 
                // use if that layer's default Material is <By Category>
                Autodesk.Revit.DB.Category wallCategory = allCategories.get_Item(Autodesk.Revit.DB.BuiltInCategory.OST_Walls);
                Autodesk.Revit.DB.Material wallMaterial = wallCategory.Material;

                foreach (Autodesk.Revit.DB.CompoundStructureLayer structLayer in comStruct.GetLayers())
                {
                    Autodesk.Revit.DB.Material layerMaterial =
                    documento.GetElement(structLayer.MaterialId) as Autodesk.Revit.DB.Material;

                    // If CompoundStructureLayer's Material is specified, use default
                    // Material of its Category
                    if (null == layerMaterial)
                    {
                        switch (structLayer.Function)
                        {
                            case Autodesk.Revit.DB.MaterialFunctionAssignment.Finish1:
                                layerMaterial =
                                        allCategories.get_Item(Autodesk.Revit.DB.BuiltInCategory.OST_WallsFinish1).Material;
                                break;
                            case Autodesk.Revit.DB.MaterialFunctionAssignment.Finish2:
                                layerMaterial =
                                        allCategories.get_Item(Autodesk.Revit.DB.BuiltInCategory.OST_WallsFinish2).Material;
                                break;
                            case Autodesk.Revit.DB.MaterialFunctionAssignment.Membrane:
                                layerMaterial =
                                        allCategories.get_Item(Autodesk.Revit.DB.BuiltInCategory.OST_WallsMembrane).Material;
                                break;
                            case Autodesk.Revit.DB.MaterialFunctionAssignment.Structure:
                                layerMaterial =
                                        allCategories.get_Item(Autodesk.Revit.DB.BuiltInCategory.OST_WallsStructure).Material;
                                break;
                            case Autodesk.Revit.DB.MaterialFunctionAssignment.Substrate:
                                layerMaterial =
                                        allCategories.get_Item(Autodesk.Revit.DB.BuiltInCategory.OST_WallsSubstrate).Material;
                                break;
                            case Autodesk.Revit.DB.MaterialFunctionAssignment.Insulation:
                                layerMaterial =
                                        allCategories.get_Item(Autodesk.Revit.DB.BuiltInCategory.OST_WallsInsulation).Material;
                                break;
                            default:
                                // It is impossible to reach here
                                break;
                        }
                        if (null == layerMaterial)
                        {
                            // CompoundStructureLayer's default Material is its SubCategory
                            layerMaterial = wallMaterial;
                        }
                    }
                    TaskDialog.Show("Revit", "Layer Material: " + layerMaterial);
                }
            }
        }
        static public void GetFamilyTypesInFamily(Autodesk.Revit.DB.Document familyDoc)
        {
            if (familyDoc.IsFamilyDocument == true)
            {
                Autodesk.Revit.DB.FamilyManager familyManager = familyDoc.FamilyManager;
                // get types in family
                string types = "Family Types: ";
                Autodesk.Revit.DB.FamilyTypeSet familyTypes = familyManager.Types;
                Autodesk.Revit.DB.FamilyTypeSetIterator familyTypesItor = familyTypes.ForwardIterator();
                familyTypesItor.Reset();
                while (familyTypesItor.MoveNext())
                {
                    Autodesk.Revit.DB.FamilyType familyType = familyTypesItor.Current as Autodesk.Revit.DB.FamilyType;
                    types += "\n" + familyType.Name;
                }
                wf.MessageBox.Show(types, "Revit");
            }
        }

        //inicio

        static public void UpdateLeaders(Autodesk.Revit.DB.Document doc)
        {
            //UIDocument uidoc = this.ActiveUIDocument;

            // create a collection of all the Arrowhead types
            Autodesk.Revit.DB.ElementId id = new Autodesk.Revit.DB.ElementId(Autodesk.Revit.DB.BuiltInParameter.ALL_MODEL_FAMILY_NAME);
            Autodesk.Revit.DB.ParameterValueProvider provider = new Autodesk.Revit.DB.ParameterValueProvider(id);
            Autodesk.Revit.DB.FilterStringRuleEvaluator evaluator = new Autodesk.Revit.DB.FilterStringEquals();
#if D23 || D24
            Autodesk.Revit.DB.FilterRule rule = new Autodesk.Revit.DB.FilterStringRule(provider, evaluator, "Basic Wall");
#else
            FilterRule rule = new FilterStringRule(provider, evaluator, "Basic Wall", false);
#endif

            Autodesk.Revit.DB.ElementParameterFilter filter = new Autodesk.Revit.DB.ElementParameterFilter(rule);
            Autodesk.Revit.DB.FilteredElementCollector collector2 = new Autodesk.Revit.DB.FilteredElementCollector(doc).OfClass(typeof(Autodesk.Revit.DB.ElementType)).WherePasses(filter);

            Autodesk.Revit.DB.ElementId arrowheadId = null;

            // loop through the collection of Arrowhead types and get the Id of a specific one
            foreach (Autodesk.Revit.DB.ElementId eid in collector2.ToElementIds())
            {
                if (doc.GetElement(eid).Name == "Wall 2") //change the name to the arrowhead type you want to use
                {
                    arrowheadId = eid;
                    TaskDialog.Show("Entrou", "Sim?");

                }
            }



            // create a collection of all families
            List<Autodesk.Revit.DB.Family> families = new List<Autodesk.Revit.DB.Family>(new Autodesk.Revit.DB.FilteredElementCollector(doc).OfClass(typeof(Autodesk.Revit.DB.Family)).Cast<Autodesk.Revit.DB.Family>());
            using (Autodesk.Revit.DB.Transaction t = new Autodesk.Revit.DB.Transaction(doc, "Update Leaders"))
            {
                t.Start();
                // loop through all families and get their types       
                foreach (Autodesk.Revit.DB.Family f in families)
                {
                    ISet<Autodesk.Revit.DB.ElementId> symbols = f.GetFamilySymbolIds();

                    // loop through the family types and try switching the parameter Leader Arrowhead to the one chosen above
                    // if the type doesn't have this parameter it will skip to the next family type
                    foreach (Autodesk.Revit.DB.ElementId eleId in symbols)
                    {
                        Autodesk.Revit.DB.FamilySymbol fs = doc.GetElement(eleId) as Autodesk.Revit.DB.FamilySymbol;
                        try

                        {
                            fs.LookupParameter("Name").Set(arrowheadId);
                            TaskDialog.Show("Entrou", "Sim?");
                        }
                        catch (Exception exception)
                        {
                            TaskDialog.Show("Entrou", arrowheadId.ToString());
                            TaskDialog.Show("Entrou", exception.Message);

                        }
                    }
                }
                t.Commit();
            }
        }

        internal static double GetDiametro(Autodesk.Revit.DB.Face face)
        {
            List<Autodesk.Revit.DB.XYZ> listaPontos = new List<Autodesk.Revit.DB.XYZ>();
            bool ecurvo = true;
            List<double> lista = new List<double>();
            lista.Add(0);
            List<Autodesk.Revit.DB.CurveLoop> curveLoops = face.GetEdgesAsCurveLoops().ToList();
            foreach (Autodesk.Revit.DB.CurveLoop curveLoops2 in curveLoops)
            {
                foreach (Autodesk.Revit.DB.Curve curve in curveLoops2)
                {
                    if(curve is Autodesk.Revit.DB.Arc)
                    {
                        lista.Add((curve as Autodesk.Revit.DB.Arc).Radius * 2);
                    } else
                    {
                        ecurvo = false;
                        if (!listaPontos.Contains((curve as Autodesk.Revit.DB.Line).GetEndPoint(0)))
                            listaPontos.Add((curve as Autodesk.Revit.DB.Line).GetEndPoint(0));

                        if (!listaPontos.Contains((curve as Autodesk.Revit.DB.Line).GetEndPoint(1)))
                            listaPontos.Add((curve as Autodesk.Revit.DB.Line).GetEndPoint(1));
                         

                    }

                }
            }
            if (ecurvo)
            {
                lista.Sort();
                return lista.LastOrDefault();
            }
            else return Util.GetMaiorDistancia(listaPontos);
        }

        private static double GetMaiorDistancia(List<Autodesk.Revit.DB.XYZ> listaPontos)
        {
            List<double> d = new List<double>();
            foreach (var item in listaPontos)
            {
                foreach (var item1 in listaPontos)
                {
                    d.Add(Util.GetDistanciaEntreDoisPontos(item1, item));

                }
            }
            d.Sort();
            return d.LastOrDefault();
        }

        public static void ObterAmbiente(Autodesk.Revit.DB.Document uiDoc, Autodesk.Revit.DB.ElementId item)
        {
            Autodesk.Revit.DB.Element ele = uiDoc.GetElement(item);
            Autodesk.Revit.DB.XYZ ponto = new Autodesk.Revit.DB.XYZ(0, 0, 0);
            if (ele== null) return;

            if (ele.Location == null) return;

            if ((ele.Location is Autodesk.Revit.DB.LocationPoint))
                ponto = (ele.Location as Autodesk.Revit.DB.LocationPoint).Point;
            else ponto = (ele.Location as Autodesk.Revit.DB.LocationCurve).Curve.GetEndPoint(0);

            var room = uiDoc.GetSpaceAtPoint(ponto);
            if (room != null)
            {
                try
                {
                    ele.LookupParameter("tocAmbienteSistema").Set(room.LookupParameter("Nome").AsString());
                    ele.LookupParameter("tocAmbienteSistemaId").Set(room.Id.IntegerValue);
                    ele.LookupParameter("tocAmbienteNivel").Set(room.Level.Name);
                    ele.LookupParameter("tocAmbienteNivelId").Set(room.Level.Id.IntegerValue);
                    if(room.Zone !=null)
                    {
                        ele.LookupParameter("tocZona").Set(room.Zone.Name);
                        ele.LookupParameter("tocZonalId").Set(room.Zone.Id.IntegerValue);
                    }
              
                }
                catch
                {

                }
            }
        }
        public static Autodesk.Revit.DB.Element FindElementByName(
     Type targetType,
     string targetName, Autodesk.Revit.DB.Document iuidoc)
        {
            return new Autodesk.Revit.DB.FilteredElementCollector(iuidoc)
              .OfClass(targetType)
              .FirstOrDefault<Autodesk.Revit.DB.Element>(
                e => e.Name.Equals(targetName));
        }
        public static Autodesk.Revit.DB.ElementId GetSolidFill(Autodesk.Revit.DB.Document uidoc)
        {

            Autodesk.Revit.DB.Element padrao = Util.FindElementByName(typeof(Autodesk.Revit.DB.FillPatternElement), "Preenchimento sólido", uidoc);
            if (padrao == null)
                padrao = Util.FindElementByName(typeof(Autodesk.Revit.DB.FillPatternElement), "Solid fill", uidoc);
            if (padrao == null)
                padrao = Util.FindElementByName(typeof(Autodesk.Revit.DB.FillPatternElement), "<Preenchimento sólido>", uidoc);
            return padrao.Id;

        }
        internal static IEnumerable<Autodesk.Revit.DB.ElementId> GetFilterElementByParameterId(Autodesk.Revit.DB.Document uiDoc, string paramater1, string valor1, string paramater2, string valor2 )
        {

            var p = uiDoc.ParameterBindings.ForwardIterator();
            p.Reset();
            Autodesk.Revit.DB.Definition d1 = null;
            while (p.MoveNext())
            {
                if (p.Key.Name == paramater1)
                {
                    d1 = p.Key;
                }
            }
            p.Reset();
            Autodesk.Revit.DB.Definition d2 = null;
            p = uiDoc.ParameterBindings.ForwardIterator();
            while (p.MoveNext())
            {
                if (p.Key.Name == paramater2)
                {
                    d2 = p.Key;
                }
            }
            if (d2 == null) return new List<Autodesk.Revit.DB.ElementId>();

            Autodesk.Revit.DB.ParameterValueProvider provider1 = new Autodesk.Revit.DB.ParameterValueProvider((d1 as Autodesk.Revit.DB.InternalDefinition).Id);
            Autodesk.Revit.DB.ParameterValueProvider provider2 = new Autodesk.Revit.DB.ParameterValueProvider((d2 as Autodesk.Revit.DB.InternalDefinition).Id);


            Autodesk.Revit.DB.FilterStringRuleEvaluator evaluator1 = new Autodesk.Revit.DB.FilterStringEquals();
            Autodesk.Revit.DB.FilterStringRuleEvaluator evaluator2 = new Autodesk.Revit.DB.FilterStringEquals();

#if D23 || D24
            Autodesk.Revit.DB.FilterRule rule1 = new Autodesk.Revit.DB.FilterStringRule(provider1, evaluator1, valor1);
            Autodesk.Revit.DB.FilterRule rule2 = new Autodesk.Revit.DB.FilterStringRule(provider2, evaluator2, valor2);
#else
            FilterRule rule1 = new FilterStringRule(provider1, evaluator1, valor1, false);
            FilterRule rule2 = new FilterStringRule(provider2, evaluator2, valor2, false);
#endif

            List<Autodesk.Revit.DB.FilterRule> lista = new List<Autodesk.Revit.DB.FilterRule>();
            lista.Add(rule1);
            lista.Add(rule2);

            Autodesk.Revit.DB.ElementParameterFilter filter = new Autodesk.Revit.DB.ElementParameterFilter(lista, false);

            Autodesk.Revit.DB.FilteredElementCollector collector = new Autodesk.Revit.DB.FilteredElementCollector(uiDoc).WherePasses(filter);

            return collector.ToElementIds();
        }


        internal static IEnumerable<Autodesk.Revit.DB.Element> GetFilterElementByParameter(Autodesk.Revit.DB.Document uiDoc, string paramater, int valor)
        {
         
            var p = uiDoc.ParameterBindings.ForwardIterator();
            p.Reset();
            Autodesk.Revit.DB.Definition d = null;
            while ( p.MoveNext())
            {
                if(p.Key.Name==paramater)
                {
                    d = p.Key;
                }
            }
            if (d == null) return new List<Autodesk.Revit.DB.Element>();

            Autodesk.Revit.DB.ParameterValueProvider provider = new Autodesk.Revit.DB.ParameterValueProvider((d as Autodesk.Revit.DB.InternalDefinition).Id);

            Autodesk.Revit.DB.FilterNumericRuleEvaluator evaluator = new Autodesk.Revit.DB.FilterNumericEquals();



            Autodesk.Revit.DB.FilterRule rule = new Autodesk.Revit.DB.FilterIntegerRule( provider, evaluator, valor);

            Autodesk.Revit.DB.ElementParameterFilter filter= new Autodesk.Revit.DB.ElementParameterFilter(rule);

            Autodesk.Revit.DB.FilteredElementCollector collector = new Autodesk.Revit.DB.FilteredElementCollector(uiDoc).WherePasses(filter);
         
            return collector.ToElements();
        }


        [TransactionAttribute(TransactionMode.Manual)]
        [RegenerationAttribute(RegenerationOption.Manual)]
        public class VincularTubo : Autodesk.Revit.UI.IExternalCommand
        {
            public static Autodesk.Revit.DB.Plumbing.Pipe tuboAvanco; //= new Autodesk.Revit.DB.Plumbing.Pipe();
            public static Autodesk.Revit.DB.Plumbing.Pipe tuboLevantamento; //= new Autodesk.Revit.DB.Plumbing.Pipe();
            public static string PLAN_SERVICO_AMO;
            public Result Execute(ExternalCommandData commandData, ref string message, Autodesk.Revit.DB.ElementSet elements)
            {
                UIApplication uiApp = commandData.Application;
                Autodesk.Revit.DB.Document uiDoc = uiApp.ActiveUIDocument.Document;
                // selecionar
                Selection sel = uiApp.ActiveUIDocument.Selection;
                int contador = 0;

                Autodesk.Revit.DB.Element ele;
                //FuncoesDiversas.UpdateLeaders(uiDoc);
                foreach (Autodesk.Revit.DB.ElementId eleId in sel.GetElementIds())
                {
                    ele = uiDoc.GetElement(eleId);
                    string texto = string.Format("Name: {0} | {1} | {2} \nCategoria: ", ele.Name, ele.Category.Name, ele.GetType());
                    TaskDialog.Show("Informação do elemento", texto);
                    Autodesk.Revit.DB.Transaction trans = new Autodesk.Revit.DB.Transaction(uiDoc);

                    switch (contador)

                    {
                        case 0:
                            tuboAvanco = ele as Autodesk.Revit.DB.Plumbing.Pipe;
                            break;
                        case 1:
                            tuboLevantamento = ele as Autodesk.Revit.DB.Plumbing.Pipe;
                            PLAN_SERVICO_AMO = ele.LookupParameter("Mark").AsString();
                            PLAN_SERVICO_AMO = ele.GetParameters("Mark")[0].AsValueString();
                            break;
                    }
                    contador = contador + 1;
                    Autodesk.Revit.DB.WallType symbol = Util.FindElementByName(typeof(Autodesk.Revit.DB.WallType), "Wall 2") as Autodesk.Revit.DB.WallType;

                }

                if (PLAN_SERVICO_AMO == "")
                {
                    return Result.Failed;
                }

                Autodesk.Revit.DB.Transaction trans1 = new Autodesk.Revit.DB.Transaction(uiDoc);

                trans1.Start("Lab");
                tuboAvanco.LookupParameter("Mark").SetValueString(PLAN_SERVICO_AMO);
                trans1.Commit();
                trans1.Dispose();


                return Result.Succeeded;

            }
        }

        //termino
    }







}

