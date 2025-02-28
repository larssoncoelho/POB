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
using wf = System.Windows.Forms;
using POB.ObjetoDeTranferencia;
using static System.Windows.Forms.LinkLabel;
using System.Threading;
using Newtonsoft.Json;

namespace POB
{

    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]

    public class ObterId : IExternalCommand
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        public static extern ushort GlobalAddAtom(string lpString);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        public static extern bool PostMessage(IntPtr hWnd, uint Msg, ushort wParam, ushort lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern bool SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        const int WM_USER = 0x0400;
        const uint MSG_DIRETA = WM_USER + 1;
        private ElementId baseLevelId;

        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit, ref string message, ElementSet elements)
        {
            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            Selection sel = uiApp.ActiveUIDocument.Selection;
            var itens =  sel.PickObjects(ObjectType.LinkedElement);
            string texto = "";
            List<POB.ObjetoDeTranferencia.ListaDeId> listaDeId = new List< ObjetoDeTranferencia.ListaDeId>();
            foreach (var r in itens)
            {
                var link = (uiDoc.GetElement(r.ElementId) as RevitLinkInstance).Name.Split(':')[0];
                texto = texto + '\n' + link + "\t" + r.LinkedElementId.IntegerValue.ToString();
                listaDeId.Add(new ObjetoDeTranferencia.ListaDeId
                {
                    Modelo = link,
                    Id = r.LinkedElementId.IntegerValue.ToString()
                });


            }
            var listaAgregada = listaDeId.GroupBy(o => o.Modelo).Select(group=>
            new ListaDeId {Modelo = group.Key, 
                   Id = string.Join(", ", group.Select(item=>item.Id))}).ToList();
            string texto1= "";
            foreach (var a in listaAgregada)
            {
                texto1 = texto1 + ';' + a.Modelo + ": " + a.Id;
            }
            var modeloAgregado = (from a in listaDeId
                                 group a by new
                                 {
                                     a.Modelo
                                 } into f
                                 select new ListaDeId
                                 {
                                     Modelo = f.Key.Modelo
                                 }).ToList();
            var texto2 = "";
            foreach (var a in modeloAgregado)
            {
                texto2 = texto2 + ';' + a.Modelo;
            }

            var dados = new
            {
                ids = texto,
                modelosEId = texto1,
                modelos = texto2
            };



            var json = JsonConvert.SerializeObject(dados);
            wf.Clipboard.SetText(json);
           
            File.WriteAllText(@"d:\dadosId.txt", json);
            try
            {
                ushort wParam = GlobalAddAtom("Acao:");
                ushort wValor = GlobalAddAtom("idsElementos");
                string handleHex = File.ReadLines(@"d:\handle.txt").First();
                IntPtr hWnd = (IntPtr)Convert.ToInt32(handleHex, 16);
                if (hWnd != IntPtr.Zero)
                {
                    PostMessage(hWnd, MSG_DIRETA, wParam, wValor);
                }
                else
                {
                }

            }
            catch (Exception e)
            {
            }

            return Result.Succeeded;
        }
    }
}
