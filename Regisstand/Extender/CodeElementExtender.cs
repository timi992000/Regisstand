using EnvDTE;

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


  }
}
