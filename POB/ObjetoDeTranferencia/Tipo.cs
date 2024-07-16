using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace POB.ObjetoDeTranferencia
{
    public class Tipo
    {

      public int  id {  get; set; }
      public string Categoria { get; set; }
      public string Familia {  get; set; }      
      public string DescricaoTipo {  get; set; }
      public string ComentariosDeTipo {  get; set; }

    }
}
