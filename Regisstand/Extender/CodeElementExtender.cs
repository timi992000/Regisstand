using EnvDTE;
using Regisstand.Core;
using System;
using System.Collections.Generic;
using System.Linq;

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

        public static string[] GetDocumentContent(this Document document)
        {
            var textDocument = (TextDocument)document.Object("TextDocument");
            var documentText = textDocument.StartPoint.CreateEditPoint().GetText(textDocument.EndPoint);
            var fileLines = documentText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            return fileLines;
        }

        public static void CheckCodeElementAndRenameIfNeeded(this CodeElement elem)
        {
            if (elem.IsPrivateMethod() && !elem.Name.StartsWith(PRIVATEMETHODPREFIX))
                elem.RenameCodeElement(PRIVATEMETHODPREFIX + elem.Name);
        }

        public static bool IsPrivateMethod(this CodeElement elem)
        {
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
                foreach (var member in codeClass.Members)
                    actionToExecute?.Invoke(member as CodeElement);
            });
        }

        public static void ActionOnAllCodeClassesInDocument(Action<CodeClass> actionToExecute)
        {
            var textDocument = PackageContext.Instance.DTE.ActiveDocument;
            var codeModel = textDocument.ProjectItem.FileCodeModel;
            foreach (EnvDTE.CodeElement codeElement in codeModel.CodeElements)
            {
                try
                {
                    if (codeElement.Kind != EnvDTE.vsCMElement.vsCMElementNamespace)
                        continue;
                    EnvDTE.CodeNamespace codeNamespace = (EnvDTE.CodeNamespace)codeElement;
                    foreach (EnvDTE.CodeElement nestedCodeElement in codeNamespace.Members)
                    {
                        if (nestedCodeElement.Kind != vsCMElement.vsCMElementClass)
                            continue;
                        else
                        {
                            var castedCodeClass = (EnvDTE.CodeClass)nestedCodeElement;
                            actionToExecute?.Invoke(castedCodeClass);
                        }
                    }
                }
                catch (Exception)
                {
                }
            }
        }

        public static void SortClassMembers(this CodeClass codeClass)
        {
            //TODO: WHEN MOVE MEMBERS CHECK COMMENTS, REGIONS, ATTRIBUTES and MOVE TOO
             

            // Erzeuge eine neue geordnete Liste der Klassenmitglieder
            List<CodeElement> sortedMembers = new List<CodeElement>();

            // Füge private Felder oder Properties hinzu
            foreach (CodeElement member in codeClass.Members)
            {
                if (IsPrivateFieldOrProperty(member))
                    sortedMembers.Add(member);
            }

            // Füge Konstruktoren hinzu
            foreach (CodeElement member in codeClass.Members)
            {
                if (member.Kind == vsCMElement.vsCMElementFunction)
                {
                    var func = member as CodeFunction;
                    if (func.FunctionKind == vsCMFunction.vsCMFunctionConstructor)
                        sortedMembers.Add(member);
                }
            }

            // Füge public Properties hinzu
            foreach (CodeElement member in codeClass.Members)
            {
                if (IsPublicProperty(member))
                    sortedMembers.Add(member);
            }

            // Füge public Methoden hinzu
            foreach (CodeElement member in codeClass.Members)
            {
                if (IsPublicMethod(member))
                    sortedMembers.Add(member);
            }

            // Füge private Methoden hinzu
            foreach (CodeElement member in codeClass.Members)
            {
                if (IsPrivateMethod(member))
                    sortedMembers.Add(member);
            }

            var elementBuffer = string.Join("\n\n", sortedMembers.Select(m => m.__GetCodeElementText()));

            // Lösche die vorhandenen Klassenmitglieder
            foreach (CodeElement member in codeClass.Members)
                codeClass.RemoveMember(member);

            var editPoint = codeClass.StartPoint.CreateEditPoint();
            editPoint.LineDown(2);
            editPoint.Insert(elementBuffer);
        }

        private static string __GetCodeElementText(this CodeElement codeElement)
        {
            try
            {
                var buffer = PackageContext.Instance.CurrentDocument.GetDocumentContent();
                var startPoint = codeElement.StartPoint;
                var endPoint = codeElement.EndPoint;

                var codeElementLines = new List<string>();
                for (int i = startPoint.Line - 1; i < endPoint.Line; i++)
                    codeElementLines.Add(buffer[i]);
                return string.Join(Environment.NewLine, codeElementLines);
            }
            catch (Exception ex)
            {
                return "";
            }
        }

        private static bool IsPrivateFieldOrProperty(CodeElement codeElement)
        {
            return codeElement.Kind == vsCMElement.vsCMElementVariable;
        }

        private static bool IsPublicProperty(CodeElement codeElement)
        {
            return codeElement.Kind == vsCMElement.vsCMElementProperty;
        }

        private static bool IsPublicMethod(CodeElement codeElement)
        {
            return codeElement.Kind == vsCMElement.vsCMElementFunction; // Beispiel: Annahme, dass es sich um eine öffentliche Methode handelt
        }

        public static void RemoveLinesWithWord(this Document document, string word)
        {
            TextSelection selection = document.Selection as TextSelection;
            selection.StartOfDocument();
            while (true)
            {
                if (selection.FindText(word, (int)vsFindOptions.vsFindOptionsMatchWholeWord))
                {
                    EditPoint start = selection.AnchorPoint.CreateEditPoint();
                    start.StartOfLine();
                    EditPoint end = selection.AnchorPoint.CreateEditPoint();
                    end.EndOfLine();
                    start.Delete(end);
                }
                else
                    break;
            }
        }
    }
}
