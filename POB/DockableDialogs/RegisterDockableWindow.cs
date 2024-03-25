using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Reflection;
using wf = System.Windows.Forms;
using cl = System.Windows.Clipboard;
using System.Windows.Media.Imaging;
using Funcoes;

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

using System.Threading;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;

using Autodesk.Revit.UI.Events;

namespace POB.DockableDialogs
{
    [Transaction(TransactionMode.Manual)]
    public class RegisterDockableWindow : IExternalCommand
    {
       MainPage m_MyDockableWindow = null;

        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            DockablePaneProviderData data            = new DockablePaneProviderData();

            MainPage MainDockableWindow = new MainPage();

            m_MyDockableWindow = MainDockableWindow;

            //MainDockableWindow.SetupDockablePane(me);

            data.FrameworkElement = MainDockableWindow
              as System.Windows.FrameworkElement;

            data.InitialState = new DockablePaneState();

            data.InitialState.DockPosition            = DockPosition.Tabbed;

            //DockablePaneId targetPane;
            //if (m_targetGuid == Guid.Empty)
            //    targetPane = null;
            //else targetPane = new DockablePaneId(m_targetGuid);
            //if (m_position == DockPosition.Tabbed)

            data.InitialState.TabBehind = DockablePanes.BuiltInDockablePanes.ProjectBrowser;

            DockablePaneId dpid = new DockablePaneId( new Guid("{D7C963CE-B7CA-426A-8D51-6E8254D21157}"));

            commandData.Application.RegisterDockablePane(
                dpid, "AEC Dockable Window", MainDockableWindow     as IDockablePaneProvider);

            commandData.Application.ViewActivated         += new EventHandler<ViewActivatedEventArgs>(       Application_ViewActivated);

            return Result.Succeeded;
        }

        void Application_ViewActivated(object sender, ViewActivatedEventArgs e)
        {
            //m_MyDockableWindow.lblProjectName.Content   = e.Document.ProjectInformation.Name;
        }
    }
}
