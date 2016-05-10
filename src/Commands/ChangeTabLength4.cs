using System;
using System.ComponentModel.Design;
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
      var settingsManager = new ShellSettingsManager(ServiceProvider);
      var settingsStore = settingsManager.GetWritableSettingsStore(SettingsScope.UserSettings);
      var tabSize = settingsStore.GetInt32(CollectionPath, "Tab Size", -1);
      var indentSize = settingsStore.GetInt32(CollectionPath, "Indent Size", -1);
      if (tabSize != -1 && indentSize != -1)
      {
        settingsStore.SetInt32(CollectionPath, "Tab Size", 4);
        settingsStore.SetInt32(CollectionPath, "Indent Size", 4);
      }
    }
  }
}