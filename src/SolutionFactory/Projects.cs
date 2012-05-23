using System.IO;
using System.Linq;
using EnvDTE;
using EnvDTE80;
using VSLangProj;

namespace SolutionFactory
{
    public static class Projects
    {
        public static void AddProjectFromTemplate(object dte, string templateName, string projectName)
        {
            var dte1 = (DTE)dte;
            string path = Path.GetDirectoryName(dte1.Solution.FileName);
            var solution2 = ((Solution2) dte1.Solution);
            string template = solution2.GetProjectTemplate(templateName, "CSharp");
            solution2.AddFromTemplate(template, path + "\\" + projectName, projectName);
        }

        public static void AddProject(object dte, string projectName)
        {
            AddProjectFromTemplate(dte, "ClassLibrary.zip", projectName);
        }

        public static void AddReference(object dteObject, string target, string reference)
        {
            DTE dte = (DTE)dteObject;
            var targetProject = (VSProject) dte.GetProject(target).Object;
            var referenceProject = dte.GetProject(reference);
            targetProject.References.AddProject(referenceProject);
        }

        public static void AddLibraryReference(object dteObject, string target, string library)
        {
            DTE dte = (DTE)dteObject;
            var targetProject = (VSProject)dte.GetProject(target).Object;
            targetProject.References.Add(library);
        }

        public static void RemoveLibraryReference(object dteObject, string target, string library)
        {
            DTE dte = (DTE)dteObject;
            var targetProject = (VSProject)dte.GetProject(target).Object;
            var refToRemove = targetProject.References.Cast<Reference>().Where(assembly => assembly.Name.EndsWith(library, System.StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
            if (refToRemove != null)
            {
                refToRemove.Remove();
            }
        }

        private static Project GetProject(this DTE dte, string target)
        {
            return dte.Solution.Projects.Cast<Project>().Where(project => project.Name.Equals(target,System.StringComparison.InvariantCultureIgnoreCase)).First();
        }

        public static void SetNamespace(object dteObject,string project,string defaultNamespace)
        {
            var dte = (DTE)dteObject;
            dte.GetProject(project).Properties.Item("DefaultNamespace").Value = defaultNamespace;
        }
    }
}