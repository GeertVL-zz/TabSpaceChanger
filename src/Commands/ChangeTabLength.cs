//------------------------------------------------------------------------------
// <copyright file="ChangeTabLength.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Settings;

namespace TabSpaceChanger
{
  /// <summary>
  /// Command handler
  /// </summary>
  internal sealed class ChangeTabLength
  {
    private const string CollectionPath = @"Text Editor\CSharp";
    private const string TabSize = "Tab Size";
    private const string IndentSize = "Indent Size";

    /// <summary>
    /// Command ID.
    /// </summary>
    public const int CommandId = 0x0100;

    /// <summary>
    /// Command menu group (command set GUID).
    /// </summary>
    public static readonly Guid CommandSet = new Guid("343219ff-df97-47c2-94c8-ae391c5f1fdc");

    /// <summary>
    /// VS Package that provides this command, not null.
    /// </summary>
    private readonly Package package;

    /// <summary>
    /// Initializes a new instance of the <see cref="ChangeTabLength"/> class.
    /// Adds our command handlers for menu (commands must exist in the command table file)
    /// </summary>
    /// <param name="package">Owner package, not null.</param>
    private ChangeTabLength(Package package)
    {
      if (package == null)
      {
        throw new ArgumentNullException(nameof(package));
      }

      this.package = package;

      OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
      if (commandService != null)
      {
        var menuCommandID = new CommandID(CommandSet, CommandId);
        var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
        commandService.AddCommand(menuItem);
      }
    }

    /// <summary>
    /// Gets the instance of the command.
    /// </summary>
    public static ChangeTabLength Instance
    {
      get;
      private set;
    }

    /// <summary>
    /// Gets the service provider from the owner package.
    /// </summary>
    private IServiceProvider ServiceProvider
    {
      get
      {
        return this.package;
      }
    }

    /// <summary>
    /// Initializes the singleton instance of the command.
    /// </summary>
    /// <param name="package">Owner package, not null.</param>
    public static void Initialize(Package package)
    {
      Instance = new ChangeTabLength(package);
    }

    /// <summary>
    /// This function is the callback used to execute the command when the menu item is clicked.
    /// See the constructor to see how the menu item is associated with this function using
    /// OleMenuCommandService service and MenuCommand class.
    /// </summary>
    /// <param name="sender">Event sender.</param>
    /// <param name="e">Event args.</param>
    private void MenuItemCallback(object sender, EventArgs e)
    {
      try
      {
        var settingsManager = new ShellSettingsManager(ServiceProvider);
        var settingsStore = settingsManager.GetWritableSettingsStore(SettingsScope.UserSettings);
        if (settingsStore.PropertyExists(CollectionPath, TabSize))
        {
          settingsStore.SetUInt32(CollectionPath, TabSize, 2);
        }
        if (settingsStore.PropertyExists(CollectionPath, IndentSize))
        {
          settingsStore.SetUInt32(CollectionPath, IndentSize, 2);
        }
      }
      catch
      {
        // ignore
      }
    }
  }
}
