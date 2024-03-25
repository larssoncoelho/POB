using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Reflection;
using wf = System.Windows.Forms;
using cl = System.Windows.Clipboard;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Data;
namespace POB.DockableDialogs
{
    [Transaction(TransactionMode.ReadOnly)]
    class CmdSelectionChanged : IExternalCommand
    {
        static UIApplication _uiapp;
        static bool _subscribed = false;

        void PanelEvent(
          object sender,
          System.ComponentModel.PropertyChangedEventArgs e)
        {
            /*Debug.Assert(sender is Autodesk.Windows.RibbonTab,
              "expected sender to be a ribbon tab");
              */
            if (e.PropertyName == "Title")
            {
                ICollection<ElementId> ids = _uiapp
                  .ActiveUIDocument.Selection.GetElementIds();

                int n = ids.Count;

                string s = (0 == n)
                  ? "<nil>"
                  : string.Join(", ",
                    ids.Select<ElementId, string>(
                      id => id.IntegerValue.ToString()));

            /*    Debug.Print(
                  "CmdSelectionChanged: selection changed: "
                  + s);*/
            }
        }

        public Result Execute( ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            _uiapp = commandData.Application;

          /*foreach (Autodesk.Windows.RibbonTab tab in
              Autodesk.Windows.ComponentManager.Ribbon.Tabs)
            {
                if (tab.Id == "Modify")
                {
                    if (_subscribed)
                    {
                        tab.PropertyChanged -= PanelEvent;
                        _subscribed = false;
                    }
                    else
                    {
                        tab.PropertyChanged += PanelEvent;
                        _subscribed = true;
                    }
                    break;
                }
            }*/

           // Debug.Print("CmdSelectionChanged: _subscribed = {0}", _subscribed);

            return Result.Succeeded;
        }
    }
}