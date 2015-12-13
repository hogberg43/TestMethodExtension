//------------------------------------------------------------------------------
// <copyright file="CreateTestMethod.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Text;
using EnvDTE;

using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace TestMethodEx
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class CreateTestMethod
    {
        private DTE dte;
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("5bec2e5e-2036-4bb2-bcef-16472fe08292");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="CreateMethodFromText"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private CreateTestMethod(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
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
        public static CreateTestMethod Instance
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
            Instance = new CreateTestMethod(package);
        }

        string CreateMethodFromText(string selection)
        {
            if (!selection.Contains("\r\n") && selection.Length > 0)
            {
                var sb = new StringBuilder();
                sb.Append("[TestMethod]\r\n");
                sb.Append("public void ");
                sb.Append(selection.Replace(" ", "_").Replace("\"","").Replace("//",""));
                sb.Append("() \r\n{\r\n");
                sb.Append("// Arrange\r\n\r\n\r\n");
                sb.Append("// Act\r\n\r\n\r\n");
                sb.Append("// Assert\r\n");
                //sb.Append("}");

                selection = sb.ToString();
            }
            return selection;
        }

        private void MenuItemCallback(object sender, EventArgs e)
        {
            this.dte = (DTE)ServiceProvider.GetService(typeof(DTE));
            if (dte.ActiveDocument != null)
            {
                var doc = (TextDocument)dte.ActiveDocument.Object("TextDocument");
                if (doc != null)
                {
                    var startPoint = doc.StartPoint.CreateEditPoint();
                    var isTestClass = startPoint.GetText(doc.EndPoint).Contains("[TestClass]");
                    if (isTestClass)
                    {
                        doc.Selection.Text = CreateMethodFromText(doc.Selection.Text);
                    }
                }
            }
        }
    }
}
