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

namespace POB
{
    public class ResultadoElemento
    {
        public ElementId ElementId { get; set; }
        public object Objeto { get; set; }
        public string Mensagem { get; set; }
        public Element Element { get;  set; }
    }
    public class ResultadoExternalCommandData
    {
        public ResultadoExternalCommandData()
        {
            Lista = new List<ResultadoElemento>();
        }
        public List<ResultadoElemento> Lista { get; set; }
        public ResultadoElemento ResultadoElemento { get; set; }
        public bool ErroGlobal { get; set; }
        public Result Resultado { get; internal set; }
        public string Mensagem { get; internal set; }
    }
}