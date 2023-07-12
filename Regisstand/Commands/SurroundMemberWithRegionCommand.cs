using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Regisstand.Core;
using System;
using System.ComponentModel.Design;
using System.Threading.Tasks;

namespace Regisstand.Commands
{
    internal sealed class SurroundMemberWithRegionCommand
    {
        public readonly Package Package;
        public const int CommandId = 257;
        public static readonly Guid CommandSet = new Guid("e73b1379-6d9d-4057-adf1-aca51e35f7c4");
        public SurroundMemberWithRegionCommand(Package package, IMenuCommandService commandService)
        {
            if (package == null)
            {
                throw new ArgumentNullException("Package");
            }

            Package = package;

            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(DoAction, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }


        public static SurroundMemberWithRegionCommand Instance
        {
            get;
            private set;
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            var svcProvider = await package.GetServiceAsync(typeof(IMenuCommandService));
            var commandService = svcProvider as OleMenuCommandService;
            PackageContext.Instance.Package = package as RegisstandPackage;
            Instance = new SurroundMemberWithRegionCommand(package, commandService);
        }


        private void DoAction(object sender, EventArgs e)
        {
            __SurroundMembersByRegions();
        }

        private void __SurroundMembersByRegions()
        {
            if (PackageContext.Instance == null || PackageContext.Instance.CurrentDocument == null)
                return;
            __SurroundAllCodeElementsWithRegion();
        }

        private void __SurroundAllCodeElementsWithRegion()
        {
            var textDocument = PackageContext.Instance.DTE.ActiveDocument;
            var codeModel = textDocument.ProjectItem.FileCodeModel;
            string latestCodeElementName;
            foreach (EnvDTE.CodeElement codeElement in codeModel.CodeElements)
            {
                try
                {
                    if (codeElement.Kind != EnvDTE.vsCMElement.vsCMElementNamespace)
                        continue;
                    EnvDTE.CodeNamespace codeNamespace = (EnvDTE.CodeNamespace)codeElement;
                    foreach (EnvDTE.CodeElement nestedCodeElement in codeNamespace.Members)
                    {
                        bool hasRegion = false;
                        //Find a way to check if region already exists
                        if (hasRegion || nestedCodeElement.Kind != vsCMElement.vsCMElementClass)
                            continue;
                        else
                        {
                            var castedCodeClass = (EnvDTE.CodeClass)nestedCodeElement;
                            foreach (var member in castedCodeClass.Members)
                            {
                                __GenerateCodeElementRegion(member as CodeElement);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                }
            }
        }

        private void __GenerateCodeElementRegion(CodeElement codeElement)
        {
            TextPoint start = codeElement.StartPoint;
            var startTabPosition = start.DisplayColumn - 1;
            TextPoint end = codeElement.EndPoint;
            var endTabPosition = end.DisplayColumn - 2;

            EditPoint startPoint = start.CreateEditPoint();
            EditPoint endPoint = end.CreateEditPoint();

            startPoint.Insert($"#region [{codeElement.Name}]\n{new string(' ', startTabPosition)}");
            endPoint.Insert($"\n{new string(' ', endTabPosition)}#endregion");
        }
    }
}
