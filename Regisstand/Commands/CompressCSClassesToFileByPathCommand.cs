using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.WindowsAPICodePack.Dialogs;
using Regisstand.Core;
using Regisstand.Extender;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Regisstand.Commands
{
    internal sealed class CompressCSClassesToFileByPathCommand
    {
        public readonly Package Package;
        public const int CommandId = 256;
        private string _SelectedDir;
        public static readonly Guid CommandSet = new Guid("e73b1379-6d9d-4057-adf1-aca51e35f7c4");
        public CompressCSClassesToFileByPathCommand(Package package, IMenuCommandService commandService)
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


        public static CompressCSClassesToFileByPathCommand Instance
        {
            get;
            private set;
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            var svcProvider = await package.GetServiceAsync(typeof(IMenuCommandService));
            var commandService = svcProvider as OleMenuCommandService;
            PackageContext.Instance.Package = package as RegisstandPackage;
            Instance = new CompressCSClassesToFileByPathCommand(package, commandService);
        }


        private void DoAction(object sender, EventArgs e)
        {
            System.Diagnostics.Debugger.Launch();
            using (var dialog = new CommonOpenFileDialog())
            {
                dialog.IsFolderPicker = true;
                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    _SelectedDir = dialog.FileName;
                    CombineFiles(PackageContext.Instance.DTE, _SelectedDir);
                }
                else
                    return;
            }
        }

        public void CombineFiles(DTE2 dte, string folderPath)
        {
            try
            {
                // Erhalte eine Referenz auf das Solution-Objekt
                Solution2 solution = (Solution2)dte.Solution;
                // Durchsuche den angegebenen Ordner nach C#-Dateien
                var allDocuments = __GetDocumentList(PackageContext.Instance.CurrentProject.ProjectItems).Where(d => d.FullName.StartsWith(folderPath));
                // Erstelle eine neue Klasse, um den kombinierten Code zu halten
                var newDocumentName = "MyCompressedClass";
                var openedFile = AddEmptyClassToCurrentProject(dte, newDocumentName, "CompressedNamespace").Document;

                // Sammle die "using"-Anweisungen aus allen Dateien
                List<string> usingStatements = new List<string>();
                foreach (var document in allDocuments)
                {
                    // Extrahiere die "using"-Anweisungen aus der Datei
                    IEnumerable<string> fileUsingStatements = document.GetDocumentContent()
                        .Where(line => line.Trim().StartsWith("using") && !line.Trim().StartsWith("using (") && !line.Trim().StartsWith("using("))
                        .Select(line => line.Trim());

                    usingStatements.AddRange(fileUsingStatements);
                }

                // Entferne doppelte "using"-Anweisungen
                usingStatements = usingStatements.Distinct().ToList();

                // Füge die kombinierten "using"-Anweisungen zur kombinierten Klasse hinzu
                foreach (string usingStatement in usingStatements)
                    __AddUsingDirectiveToClass(openedFile, usingStatement);

                // Durchlaufe alle Dateien und füge den Code zur kombinierten Klasse hinzu
                var excludedFiles = new List<string>()
            {
                "AssemblyInfo.cs",
            };
                foreach (var document in allDocuments.Where(f => !excludedFiles.Contains(f.Name)))
                {
                    var fileLines = document.GetDocumentContent();
                    var lines = new List<string>();
                    var firstBracketIndex = Array.IndexOf(fileLines, fileLines.FirstOrDefault(l => l.Trim().Equals("{")));
                    var lastBracketIndex = Array.IndexOf(fileLines, fileLines.LastOrDefault(l => l.Trim().Equals("}")));
                    foreach (var line in fileLines)
                    {
                        if ((line.Trim().StartsWith("using") && !line.Trim().StartsWith("using (") && !line.Trim().StartsWith("using(")) || line.Trim().StartsWith("namespace"))
                            continue;
                        var index = Array.IndexOf(fileLines, line);
                        if (index == firstBracketIndex || index == lastBracketIndex)
                            continue;
                        lines.Add(line);
                    }
                    // Den CodeNamespace für das TextDocument abrufen
                    CodeNamespace codeNamespace = openedFile.ProjectItem.FileCodeModel.CodeElements.OfType<CodeNamespace>().FirstOrDefault();

                    if (codeNamespace != null)
                    {
                        // Den EndPoint des CodeNamespace erhalten
                        EditPoint editPoint = codeNamespace.EndPoint.CreateEditPoint();
                        editPoint.LineUp();
                        // Den Code vor der schließenden Klammer des Namespaces einfügen
                        editPoint.Insert("\n" + string.Join(Environment.NewLine, lines.ToArray()));
                    }

                    // Speichern des Dokuments
                    openedFile.Save();
                    __DeleteDocument(document);
                }

                // Speichere die Lösung
                solution.SaveAs(solution.FullName);
            }
            catch (Exception)
            {
                System.Diagnostics.Debugger.Launch();
            }
        }

        private void __DeleteDocument(Document document)
        {
            document.ProjectItem.Remove();
            document.ProjectItem.Delete();
        }

        private List<Document> __GetDocumentList(ProjectItems projectItems)
        {
            var resultList = new List<Document>();
            if (projectItems == null)
                return resultList;

            foreach (ProjectItem projectItem in projectItems)
            {
                if (projectItem.Kind == Constants.vsProjectItemKindPhysicalFile && projectItem.FileCodeModel != null && projectItem.Name.EndsWith(".cs"))
                {
                    // Hier können Sie mit dem gewünschten Document arbeiten
                    if (projectItem.Document == null)
                        projectItem.Open();
                    if (projectItem.Document != null)
                    {
                        Document document = projectItem.Document;
                        resultList.Add(document);
                    }
                }
                else if (projectItem.Kind == Constants.vsProjectItemKindPhysicalFolder)
                {
                    // Unterverzeichnis: Die Methode rekursiv aufrufen
                    resultList.AddRange(__GetDocumentList(projectItem.ProjectItems));
                }
            }
            return resultList;
        }

        public Window AddEmptyClassToCurrentProject(DTE2 dte, string className, string namespaceName)
        {
            // Erhalte das aktive Projekt
            Project activeProject = GetActiveProject(dte);
            if (activeProject == null)
            {
                Console.WriteLine("Es ist kein aktives Projekt verfügbar.");
                return null;
            }

            var newFile = Path.Combine(_SelectedDir, className);
            newFile = Path.ChangeExtension(newFile, "cs");
            if (File.Exists(newFile))
                File.Delete(newFile);
            File.WriteAllText(newFile, $"using System;\n\nnamespace {namespaceName}\n{{\n\n\n}}");

            activeProject.ProjectItems.AddFromFile(newFile);
            var generatedFile = dte.ItemOperations.OpenFile(newFile);
            return generatedFile;
        }

        private Project GetActiveProject(DTE2 dte)
        {
            Array activeSolutionProjects = dte.ActiveSolutionProjects as Array;
            if (activeSolutionProjects != null && activeSolutionProjects.Length > 0)
            {
                return activeSolutionProjects.GetValue(0) as Project;
            }
            return null;
        }

        private void __AddUsingDirectiveToClass(Document document, string directive)
        {
            CodeElement lastUsingDirective = null;

            foreach (CodeElement ce in document.ProjectItem.FileCodeModel.CodeElements)
            {
                if (ce.Kind == vsCMElement.vsCMElementImportStmt)
                    lastUsingDirective = ce;
                else
                {
                    if (lastUsingDirective != null)
                    {
                        // insert given directive after the last one, on a new line
                        EditPoint insertPoint = lastUsingDirective.GetEndPoint().CreateEditPoint();
                        insertPoint.Insert($"\r\n{directive}");
                    }
                }
            }
            document.Save();
        }

    }
}
