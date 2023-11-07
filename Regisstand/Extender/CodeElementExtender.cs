using EnvDTE;
using Regisstand.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Regisstand.Extender
{
    public static class CodeElementExtender
    {
        public const string PRIVATEMETHODPREFIX = "__";
        public static void RenameCodeElement(this CodeElement elem, string newName)
        {
            EnvDTE80.CodeElement2 codeElement = elem as EnvDTE80.CodeElement2;
            codeElement.RenameSymbol(newName);
        }

        public static void SurroundWithRegionsInCurrentDocument()
        {
            if (PackageContext.Instance == null || PackageContext.Instance.CurrentDocument == null)
                return;
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            ActionOnAllCodeClassesInDocument((codeClass) =>
            {
                __RemoveRegionsAndEmptyLines(codeClass);
            });

            PackageContext.Instance.DTE.UndoContext.Open("Surround Members with region");

            try
            {
                ActionOnAllCodeElementsInDocument
                    ((member) =>
                    {
                        member.CheckCodeElementAndRenameIfNeeded();
                        __GenerateCodeElementRegion(member);
                    });
            }
            catch (Exception)
            { }
            finally
            {
                PackageContext.Instance.DTE.UndoContext.Close();
            }

        }

        public static string[] GetDocumentContent(this Document document)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            var textDocument = document.GetTextDocument();
            var documentText = textDocument.StartPoint.CreateEditPoint().GetText(textDocument.EndPoint);
            var fileLines = documentText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            return fileLines;
        }

        public static TextDocument GetTextDocument(this Document document)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            return (TextDocument)document.Object("TextDocument");
        }

        public static void CheckCodeElementAndRenameIfNeeded(this CodeElement elem)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            if (elem.IsPrivateMethod() && !elem.Name.StartsWith(PRIVATEMETHODPREFIX))
                elem.RenameCodeElement(PRIVATEMETHODPREFIX + elem.Name);
        }

        public static bool IsPrivateMethod(this CodeElement elem)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            if (elem.Kind == vsCMElement.vsCMElementFunction || elem.Name.Contains("."))
            {
                var function = (CodeFunction)elem;
                if (function.Access == vsCMAccess.vsCMAccessPrivate)
                    return true;
            }
            return false;
        }

        public static void ActionOnAllCodeElementsInDocument(Action<CodeElement> actionToExecute)
        {
            ActionOnAllCodeClassesInDocument((codeClass) =>
            {
                Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
                foreach (var member in codeClass.Members)
                    actionToExecute?.Invoke(member as CodeElement);
            });
        }

        public static void ActionOnAllCodeClassesInDocument(Action<CodeClass> actionToExecute)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            var textDocument = PackageContext.Instance.DTE.ActiveDocument;
            var codeModel = textDocument.ProjectItem.FileCodeModel;
            foreach (CodeElement codeElement in codeModel.CodeElements)
            {
                try
                {
                    if (codeElement.Kind != vsCMElement.vsCMElementNamespace)
                        continue;
                    CodeNamespace codeNamespace = (CodeNamespace)codeElement;
                    foreach (CodeElement nestedCodeElement in codeNamespace.Members)
                    {
                        if (nestedCodeElement.Kind != vsCMElement.vsCMElementClass)
                            continue;
                        else
                        {
                            var castedCodeClass = (CodeClass)nestedCodeElement;
                            actionToExecute?.Invoke(castedCodeClass);
                        }
                    }
                }
                catch (Exception)
                { }
            }
        }

        public static void RegexReplace(this CodeClass codeClass, string pattern, string replacement)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            var dte = PackageContext.Instance.DTE;
            var textDocument = dte.ActiveDocument.GetTextDocument();

            if (textDocument == null)
                return;

            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            var selection = textDocument.Selection;
            selection.StartOfDocument();
            selection.EndOfDocument(true);
            var replaceResult = Regex.Replace(selection.Text, pattern, replacement, RegexOptions.Multiline);
            selection.Delete();
            selection.Collapse();
            selection.TopPoint.CreateEditPoint().Insert(replaceResult);
        }

        public static void SortClassMembers(this CodeClass codeClass)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            //TODO: WHEN MOVE MEMBERS CHECK COMMENTS, REGIONS, ATTRIBUTES and MOVE TOO

            string elementBuffer = "";

            var sortedMembers = new List<CodeElement>();

            //Add private fields or properties
            foreach (CodeElement member in codeClass.Members)
            {
                if (IsPrivateFieldOrProperty(member))
                    __AddCodeElementIfIsNot(sortedMembers, member);
            }

            //Add constructors
            foreach (CodeElement member in codeClass.Members)
            {
                if (member.Kind == vsCMElement.vsCMElementFunction)
                {
                    var func = member as CodeFunction;
                    if (func.FunctionKind == vsCMFunction.vsCMFunctionConstructor)
                        __AddCodeElementIfIsNot(sortedMembers, member);
                }
            }

            //Add public Properties
            foreach (CodeElement member in codeClass.Members)
            {
                if (IsPublicProperty(member))
                    __AddCodeElementIfIsNot(sortedMembers, member);
            }

            //Add public Methods
            foreach (CodeElement member in codeClass.Members)
            {
                if (IsPublicMethod(member))
                    __AddCodeElementIfIsNot(sortedMembers, member);
            }

            //Add private Methods
            foreach (CodeElement member in codeClass.Members)
            {
                if (IsPrivateMethod(member))
                    __AddCodeElementIfIsNot(sortedMembers, member);
            }

            //Add members that do not match any type
            foreach (CodeElement member in codeClass.Members)
            {
                __AddCodeElementIfIsNot(sortedMembers, member);
            }

            if (!sortedMembers.Any())
                return;
            elementBuffer = string.Join("\n\n", sortedMembers.Select(m => m.__GetCodeElementText()));

            try
            {
                PackageContext.Instance.DTE.UndoContext.Open("Delete all classmember");
                //Delete all class members
                foreach (CodeElement member in codeClass.Members)
                    codeClass.RemoveMember(member);
            }
            catch (Exception)
            { }
            finally
            {
                PackageContext.Instance.DTE.UndoContext.Close();
            }

            //Remove regions and empty lines
            __RemoveRegionsAndEmptyLines(codeClass);

            try
            {
                PackageContext.Instance.DTE.UndoContext.Open("Insert ordered classmember");
                //Insert all ordered back
                var editPoint = codeClass.StartPoint.CreateEditPoint();
                editPoint.LineDown(2);
                editPoint.Insert(elementBuffer);
            }
            catch (Exception)
            { }
            finally
            {
                PackageContext.Instance.DTE.UndoContext.Close();
            }
        }

        private static void __AddCodeElementIfIsNot(List<CodeElement> list, CodeElement element)
        {
            if (list == null || element == null)
                return;

            if (!list.Contains(element))
                list.Add(element);
        }

        private static void __RemoveRegionsAndEmptyLines(CodeClass codeClass)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            PackageContext.Instance.DTE.UndoContext.Open("Remove Regions and empty lines");
            try
            {
                var regionRegex = @"^[ \t]*\#[ \t]*(region|endregion).*\n";
                var emptyLineRegex = @"^(?([^\r\n])\s)*\r?\n(?([^\r\n])\s)*\r?\n";
                codeClass.RegexReplace(regionRegex, "");
                codeClass.RegexReplace(emptyLineRegex, "");
            }
            catch (Exception)
            { }
            finally
            {
                PackageContext.Instance.DTE.UndoContext.Close();
            }
        }

        private static string __GetCodeElementText(this CodeElement codeElement)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                //Get comments?????????

                var buffer = PackageContext.Instance.CurrentDocument.GetDocumentContent();
                var startPoint = codeElement.StartPoint;
                var endPoint = codeElement.EndPoint;

                var codeElementLines = new List<string>();
                for (int i = startPoint.Line - 1; i < endPoint.Line; i++)
                    codeElementLines.Add(buffer[i]);
                return string.Join(Environment.NewLine, codeElementLines);
            }
            catch (Exception)
            {
                return "";
            }
        }

        private static bool IsPrivateFieldOrProperty(CodeElement codeElement)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            return codeElement.Kind == vsCMElement.vsCMElementVariable;
        }

        private static bool IsPublicProperty(CodeElement codeElement)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            return codeElement.Kind == vsCMElement.vsCMElementProperty;
        }

        private static bool IsPublicMethod(CodeElement codeElement)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            return codeElement.Kind == vsCMElement.vsCMElementFunction;
        }

        private static void __GenerateCodeElementRegion(CodeElement codeElement)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            //Only surround methods and properties (Have I forgotten any useful code element?)
            if (codeElement.Kind != vsCMElement.vsCMElementFunction && codeElement.Kind != vsCMElement.vsCMElementProperty)
                return;

            //Dont surround private properties
            if (codeElement.Kind == vsCMElement.vsCMElementProperty && (codeElement as CodeProperty).Access == vsCMAccess.vsCMAccessPrivate)
                return;

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
