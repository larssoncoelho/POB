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

using System.ComponentModel;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.ApplicationServices;
namespace POB
{
    public static class ManipulaParametroCompartilhado
    {
        public static Definition GetOrInsertCompartilhado(Autodesk.Revit.ApplicationServices.Application app, string name, string group,
#if D23 || D24  
            ForgeTypeId type,

#else
            ParameterType type,
#endif
            bool visible)
        {

            DefinitionFile defFile = app.OpenSharedParameterFile();
            if (defFile == null) throw new Exception("No SharedParameter File!");
            DefinitionGroup dg = defFile.Groups.FirstOrDefault(x => x.Name == group);
            if (dg == null) dg = defFile.Groups.Create(group);
            Definition def = (dg.Definitions).FirstOrDefault(d => d.Name == name);
            if (def != null) return def;
            return dg.Definitions.Create(new ExternalDefinitionCreationOptions(name, type));
        }
#if D23 || D24
        public static ExternalDefinition GetOrInsertCompartilhadoExternal(Autodesk.Revit.ApplicationServices.Application app, string name, string group, ForgeTypeId type, bool visible)
        {

            DefinitionFile defFile = app.OpenSharedParameterFile();
            if (defFile == null) throw new Exception("No SharedParameter File!");
            DefinitionGroup dg = defFile.Groups.FirstOrDefault(x => x.Name == group);
            if (dg == null) dg = defFile.Groups.Create(group);
            Definition def = (dg.Definitions).FirstOrDefault(d => d.Name == name);
            if (def != null) return def as ExternalDefinition;
            return dg.Definitions.Create(new ExternalDefinitionCreationOptions(name, type)) as ExternalDefinition;
        }
#else
        public static ExternalDefinition GetOrInsertCompartilhadoExternal(Autodesk.Revit.ApplicationServices.Application app, string name, string group, Autodesk.Revit.DB.ParameterType type, bool visible)
        {

            DefinitionFile defFile = app.OpenSharedParameterFile();
            if (defFile == null) throw new Exception("No SharedParameter File!");
            DefinitionGroup dg = defFile.Groups.FirstOrDefault(x => x.Name == group);
            if (dg == null) dg = defFile.Groups.Create(group);
            Definition def = (dg.Definitions).FirstOrDefault(d => d.Name == name);
            if (def != null) return def as ExternalDefinition;
            return dg.Definitions.Create(new ExternalDefinitionCreationOptions(name, type)) as ExternalDefinition;
        }
#endif
        public static void InserirParametroCompartilhadoNoProjeto(Autodesk.Revit.ApplicationServices.Application app,
                  CategorySet cats, BuiltInParameterGroup group, bool inst, string nomeGrupo)
        {

            DefinitionFile defFile = app.OpenSharedParameterFile();
            if (defFile == null) throw new Exception("No SharedParameter File!");
            DefinitionGroup dg = defFile.Groups.FirstOrDefault(x => x.Name == nomeGrupo);
            foreach (ExternalDefinition definicao in dg.Definitions)
            {
                Autodesk.Revit.DB.Binding binding = app.Create.NewTypeBinding(cats);
                if (inst) binding = app.Create.NewInstanceBinding(cats);
                BindingMap map = (new UIApplication(app)).ActiveUIDocument.Document.ParameterBindings;
                map.Insert(definicao, binding, group);
            }
        }
        
        public static void InserirParametroCompartilhadoNoProjeto(Autodesk.Revit.ApplicationServices.Application app,
                 CategorySet cats, BuiltInParameterGroup group, bool inst,
                 string nomeGrupo, string nomeParametro,
#if D24 || D23
                 ForgeTypeId parameterType
#else
                ParameterType parameterType
#endif
            )

        {


            ExternalDefinition definicao = GetOrInsertCompartilhado(app, nomeParametro, nomeGrupo, parameterType, true) as ExternalDefinition;

            Autodesk.Revit.DB.Binding binding = app.Create.NewTypeBinding(cats);
            if (inst) binding = app.Create.NewInstanceBinding(cats);
            BindingMap map = (new UIApplication(app)).ActiveUIDocument.Document.ParameterBindings;
            map.Insert(definicao, binding, group);
        }
    }
}
