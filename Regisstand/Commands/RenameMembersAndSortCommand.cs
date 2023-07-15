using Microsoft.VisualStudio.Shell;
using Regisstand.Core;
using Regisstand.Extender;
using System;
using System.ComponentModel.Design;
using System.Threading.Tasks;

namespace Regisstand.Commands
{
    internal sealed class RenameMembersAndSortCommand
    {
        public readonly Package Package;
        public const int CommandId = 259;
        public static readonly Guid CommandSet = new Guid("e73b1379-6d9d-4057-adf1-aca51e35f7c4");
        public RenameMembersAndSortCommand(Package package, IMenuCommandService commandService)
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


        public static RenameMembersAndSortCommand Instance
        {
            get;
            private set;
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            var svcProvider = await package.GetServiceAsync(typeof(IMenuCommandService));
            var commandService = svcProvider as OleMenuCommandService;
            PackageContext.Instance.Package = package as RegisstandPackage;
            Instance = new RenameMembersAndSortCommand(package, commandService);
        }


        private void DoAction(object sender, EventArgs e)
        {
            __RenameMembers();
            __SortMembers();
        }

        private void __RenameMembers()
        {
            CodeElementExtender.ActionOnAllCodeElementsInDocument
            ((member) =>
            {
                member.CheckCodeElementAndRenameIfNeeded();
            });
        }

        private void __SortMembers()
        {
            CodeElementExtender.ActionOnAllCodeClassesInDocument
            ((codeClass) =>
            {
                codeClass.SortClassMembers();
            });
        }

    }
}
