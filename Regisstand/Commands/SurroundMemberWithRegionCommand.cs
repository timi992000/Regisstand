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
      __SurroundAllMethods();
      __SurroundAllProperties();
    }


    private void __SurroundAllMethods()
    {
      ThreadHelper.ThrowIfNotOnUIThread();
      TextSelection ts = PackageContext.Instance.DTE.ActiveDocument.Selection as TextSelection;
      if (ts == null)
        return;

      CodeClass c = ts.ActivePoint.CodeElement[vsCMElement.vsCMElementClass] as CodeClass;
      if (c == null)
        return;

      foreach (CodeElement e in c.Members)
      {
        if (e.Kind == vsCMElement.vsCMElementFunction)
        {
          __GenerateCodeElementRegion(e);
        }
      }
    }

    private void __GenerateCodeElementRegion(CodeElement codeElement)
    {
      TextPoint start = codeElement.StartPoint;
      TextPoint end = codeElement.EndPoint;

      EditPoint startPoint = start.CreateEditPoint();
      EditPoint endPoint = end.CreateEditPoint();

      startPoint.Insert($"#region [{codeElement.Name}]\n");
      endPoint.Insert("\n#endregion\n");
    }

    private void __SurroundAllProperties()
    {
      ThreadHelper.ThrowIfNotOnUIThread();
      TextSelection ts = PackageContext.Instance.DTE.ActiveDocument.Selection as TextSelection;
      if (ts == null)
        return;

      CodeClass c = ts.ActivePoint.CodeElement[vsCMElement.vsCMElementClass] as CodeClass;
      if (c == null)
        return;

      foreach (CodeElement e in c.Members)
      {
        if (e.Kind == vsCMElement.vsCMElementProperty)
        {
          __GenerateCodeElementRegion(e);
        }
      }
    }
  }
}
