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
using wf = System.Windows.Forms;

using Newtonsoft.Json;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows;

//using LinqToExcel;

namespace POB
{
    
    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]

    public class DadosVista3DImagem : IExternalCommand
    {
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hWnd, ref System.Drawing.Rectangle rect);

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
        
       
       
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit,
            ref string message, ElementSet elements)
        {
            UIApplication uiApp = revit.Application;
            Document uiDoc = uiApp.ActiveUIDocument.Document;
            var vista = uiDoc.ActiveView;
            string dataToSend = "";
            if (vista is Autodesk.Revit.DB.View3D view3D)
            {
                ViewOrientation3D orientation = view3D.GetOrientation();
                XYZ eyePosition = orientation.EyePosition;
                XYZ upDirection = orientation.UpDirection;
                XYZ forwardDirection = orientation.ForwardDirection;

                // Monta o objeto a ser salvo
                var cameraData = new
                {
                    EyePosition = new { X = eyePosition.X, Y = eyePosition.Y, Z = eyePosition.Z },
                    UpDirection = new { X = upDirection.X, Y = upDirection.Y, Z = upDirection.Z },
                    ForwardDirection = new { X = forwardDirection.X, Y = forwardDirection.Y, Z = forwardDirection.Z }
                };

                //return Result.Succeeded;

                dataToSend = JsonConvert.SerializeObject(cameraData);
            } 
            else 
            {
                dataToSend = "";

            }
            var dados = new
            {
                vista = dataToSend
            };

                var json = JsonConvert.SerializeObject(dados);
            Clipboard.SetText(json);
            try
            {
                File.WriteAllText(@"d:\vista3dImagem.txt", json);
                Autodesk.Revit.DB.Transaction t = new Autodesk.Revit.DB.Transaction(uiDoc);

                string idAtual = File.ReadLines(@"d:\idAtual.txt").First().Split('|')[0]; // POB.Properties.Settings.Default.handleAta;
                string nomeVista = File.ReadLines(@"d:\idAtual.txt").First();
                var f = Util.FindElementByName(typeof(Autodesk.Revit.DB.View), nomeVista, uiDoc);
                if (f != null)
                {
                    t.Start("Deletar");
                    f.Pinned = false;
                    uiDoc.Delete(f.Id);
                    t.Commit();
                }

                t.Start("Duplicar");
                var id = vista.Duplicate(ViewDuplicateOption.WithDetailing);              
                var novaVista =( uiDoc.GetElement(id) as Autodesk.Revit.DB.View);
                         
                novaVista.get_Parameter(BuiltInParameter.VIEW_NAME).Set(nomeVista);
               
                t.Commit();
                var novoId = novaVista.Id;
                t.Start("Duplicar");
                var cortes = new FilteredElementCollector(uiDoc).OfClass(typeof(Autodesk.Revit.DB.ViewSection)).WhereElementIsNotElementType().ToElementIds();
                vista.HideCategoryTemporary(Category.GetCategory(uiDoc, BuiltInCategory.OST_Sections).Id);
                // vista.HideElements(cortes);
                try
                {
                   // uiDoc.GetElement(novoId).LookupParameter("tocAPTO").Set("fixo");
                }
                catch { }
                if (novaVista is Autodesk.Revit.DB.ViewSection)
                {
                   // novaVista.LookupParameter("tocAPTO").Set("fixo");
                }

                t.Commit();

            }
            catch (Exception ex) {
                //TaskDialog.Show("Contrutivel", ex.Message);
            }


            try
            {
                ushort wParam = GlobalAddAtom("Acao:");

                // Encontra a janela pelo título
                ushort wValor = GlobalAddAtom("vista3dImagem");
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
