using System;
using System.ComponentModel.Design;
using System.Text;
using EnvDTE;

using Microsoft.VisualStudio.Shell;

namespace TestMethodEx
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class CreateTestMethod
    {
        private DTE dte;

        #region VS template code
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
        #endregion

        /// <summary>
        /// Creates the test method to replace the text being converted.
        /// </summary>
        /// <param name="selection">Text selection in the current document</param>
        /// <returns></returns>
        string CreateMethodFromText(string selection)
        {
            if (!selection.Contains("\r\n") && selection.Length > 0)
            {
                var sb = new StringBuilder();
                sb.Append("[TestMethod]\r\n");
                sb.Append("public void ");
                sb.Append(selection.Trim().Replace(" ", "_").Replace("\"","").Replace("//",""));
                sb.Append("() \r\n{\r\n");
                sb.Append("// Arrange\r\n\r\n\r\n");
                sb.Append("// Act\r\n\r\n\r\n");
                sb.Append("// Assert\r\n");
                sb.Append("throw new NotImplementedException();\r\n"); // ensure that new test method fails (Red, Green, Refactor)
                selection = sb.ToString();
            }
            return selection;
        }

        /// <summary>
        /// Custom command implementation
        /// </summary>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            // Initialize COM library containing references to the current text editor
            dte = (DTE)ServiceProvider.GetService(typeof(DTE));

            // ensure there is an active document open before processing
            if (dte.ActiveDocument != null)
            {
                var doc = (TextDocument)dte.ActiveDocument.Object("TextDocument");
                if (doc != null)
                {

                    // Ensure that the current document is an MSTest class
                    var startPoint = doc.StartPoint.CreateEditPoint();
                    var isTestClass = startPoint.GetText(doc.EndPoint).Contains("[TestClass]");
                    if (isTestClass)
                    {
                        // Test for a current selection
                        var selectedText = doc.Selection.Text;
                        if (selectedText.Length == 0)
                        {

                            // use cursor location to select current line and create method
                            // if no selection has been made
                            var line_index = doc.Selection.ActivePoint.Line;
                            var edit_point = doc.CreateEditPoint();
                            var line = edit_point.GetLines(line_index, line_index + 1).Trim();


                            // select the entire line and replace with the method stub
                            doc.Selection.StartOfLine(0);
                            doc.Selection.EndOfLine(true);
                            doc.Selection.Text = CreateMethodFromText(line); 

                        }
                        else
                        {
                            // use the current selection to create a new method or methods if more than one line is selected
                            var lines = doc.Selection.Text.Split(new[] {"\r\n"}, StringSplitOptions.None);
                            var output = "";
                            foreach (var line in lines)
                            {
                                output += CreateMethodFromText(line) + "}\r\n";
                            }
                            doc.Selection.Text = output;
                        }
                    }
                }
            }
        }
    }
}
