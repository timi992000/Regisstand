﻿using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Regisstand.Core;
using Regisstand.Extender;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Threading.Tasks;

namespace Regisstand.Commands
{
    internal sealed class RenameFileTypesWithPrefixCommand
    {
        public readonly Package Package;
        public const int CommandId = 258;
        public static readonly Guid CommandSet = new Guid("e73b1379-6d9d-4057-adf1-aca51e35f7c4");
        public RenameFileTypesWithPrefixCommand(Package package, IMenuCommandService commandService)
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


        public static RenameFileTypesWithPrefixCommand Instance
        {
            get;
            private set;
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            var svcProvider = await package.GetServiceAsync(typeof(IMenuCommandService));
            var commandService = svcProvider as OleMenuCommandService;
            PackageContext.Instance.Package = package as RegisstandPackage;
            Instance = new RenameFileTypesWithPrefixCommand(package, commandService);
        }


        private void DoAction(object sender, EventArgs e)
        {
            __RenameTypes();
        }

        private void __RenameTypes()
        {
            if (PackageContext.Instance == null || PackageContext.Instance.CurrentDocument == null)
                return;
            __RenameTypesInCurrentDocument();
        }


        private void __RenameTypesInCurrentDocument()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            __IterateThroughClasses();
        }

        private void __IterateThroughClasses()
        {
            var prefix = __GetPrefix();
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
                        var newPrefix = prefix;
                        if (nestedCodeElement.Kind == vsCMElement.vsCMElementEnum)
                            newPrefix = "e" + prefix;
                        else if (nestedCodeElement.Kind == vsCMElement.vsCMElementInterface)
                            newPrefix = "I" + prefix;
                        var newName = newPrefix + nestedCodeElement.Name;
                        if (nestedCodeElement.Name.StartsWith(newPrefix))
                            continue;
                        else
                        {
                            nestedCodeElement.RenameCodeElement(newName);
                        }
                    }
                }
                catch (Exception)
                {
                }
            }
        }

        private string __GetPrefix()
        {
            using (var InputBox = new InputBox("Please type your prefix", new Dictionary<int, string> { { 1, "Prefix" } }, new List<int> { 1 }))
            {
                if (InputBox.OpenInputDialog())
                {
                    InputBox.OutputValues.TryGetValue(1, out string prefix);
                    return prefix;
                }
                return "";
            }
        }
    }
}
