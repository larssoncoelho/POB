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
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using Newtonsoft.Json;
using static Autodesk.Revit.DB.SpecTypeId;

//using LinqToExcel;

namespace POB
{

    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]

    public class DadosEixo : IExternalCommand
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

        public bool ENumero(string valor)
        {
            return int.TryParse(valor, out int result);
        }

        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit,
            ref string message, ElementSet elements)
        {
            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            var listaTexto = new List<string>();
            var listaNumero = new List<int>();

            foreach (ElementId element in uiApp.ActiveUIDocument.Selection.GetElementIds())
            {
            
                var ele = uiDoc.GetElement(element);
                if (ele is Autodesk.Revit.DB.Grid)
                {
                    if (ENumero((ele as Autodesk.Revit.DB.Grid).LookupParameter("Nome").AsString()))
                        listaNumero.Add(Convert.ToInt32((ele as Autodesk.Revit.DB.Grid).LookupParameter("Nome").AsString()));
                    else listaTexto.Add((ele as Autodesk.Revit.DB.Grid).LookupParameter("Nome").AsString());
                }
            }
            listaTexto = listaTexto.OrderBy(x=>x).ToList();
            listaNumero = listaNumero.OrderBy(x => x).ToList();
            var eixo = "";
            eixo = listaTexto[0]+"-"+listaTexto[1] +" X "+ listaNumero[0] + "-" + listaNumero[1];
            Clipboard.SetText(eixo);

            var vista = uiDoc.ActiveView;
            var nivel = "";
            if (vista is Autodesk.Revit.DB.ViewPlan)
            {
                var id = (vista as Autodesk.Revit.DB.ViewPlan).GenLevel;
                if(id!=null)  nivel= id.Name;
            }

            var dados = new
            {
                eixo = eixo,
                nivel = nivel
            };

            var json =JsonConvert.SerializeObject(dados);
            File.WriteAllText(@"d:\dadosEixo.txt", json);
            try
            {
                ushort wParam = GlobalAddAtom("Acao:");

                // Encontra a janela pelo título
                ushort wValor = GlobalAddAtom("eixo");
                //IntPtr hWnd = FindWindow(null, "Controle rev 01.24");

                string handleHex = File.ReadLines(@"d:\handle.txt").First(); // POB.Properties.Settings.Default.handleAta;

                // Converte para IntPtr
                IntPtr hWnd = (IntPtr)Convert.ToInt32(handleHex, 16);
                if (hWnd != IntPtr.Zero)
                {
                    Util.CaptureRevitView(Util.CaptureWindowRect());// (uiApp.ActiveUIDocument);
                    // Envia a mensagem para a janela encontrada
                    PostMessage(hWnd, MSG_DIRETA, wParam, wValor);
                 }
                else
                {
                }

            }
            catch (Exception e)
            {
              //  return Result.Canceled;

            }
         
            return Result.Succeeded;
        }
    }
}
