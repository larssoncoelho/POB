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


namespace POB.ObjetoTransferenciaPOB
{
    public class SlabShapeVertexExtendido
    {
       public SlabShapeVertex Vertex { get; set; }
       public double Distancia { get; set; }

        public XYZ PontoOriginal { get; set; }
    }
}
