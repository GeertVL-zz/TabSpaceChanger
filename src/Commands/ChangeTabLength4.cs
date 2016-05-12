using System;
using System.ComponentModel.Design;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Settings;

namespace TabSpaceChanger
{
  internal sealed class ChangeTabLength4
  {
    private const string CollectionPath = @"Text Editor\CSharp";

    public const int CommandId = 0x0110;
    public static readonly Guid CommandSet = new Guid("343219ff-df97-47c2-94c8-ae391c5f1fdc");
    private readonly Package package;
    private readonly DTE2 _dte2;

    private ChangeTabLength4(Package package)
    {
      if (package == null)
      {
        throw new ArgumentNullException(nameof(package));
      }

      this.package = package;

      OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof (IMenuCommandService)) as OleMenuCommandService;
      if (commandService != null)
      {
        var menuCommandID = new CommandID(CommandSet, CommandId);
        var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
        commandService.AddCommand(menuItem);
      }
      _dte2 = (DTE2)ServiceProvider.GetService(typeof(DTE));
    }

    public static ChangeTabLength4 Instance { get; private set; }

    /// <summary>
    /// Gets the service provider from the owner package.
    /// </summary>
    private IServiceProvider ServiceProvider
    {
      get { return this.package; }
    }

    public static void Initialize(Package package)
    {
      Instance = new ChangeTabLength4(package);
    }

    private void MenuItemCallback(object sender, EventArgs e)
    {
      _dte2.Properties["TextEditor", "CSharp"].Item("TabSize").Value = 4;
      _dte2.Properties["TextEditor", "CSharp"].Item("IndentSize").Value = 4;
      _dte2.Commands.Raise(VSConstants.CMDSETID.StandardCommandSet2K_string, (int)VSConstants.VSStd2KCmdID.FORMATDOCUMENT, null, null);
    }
  }
}