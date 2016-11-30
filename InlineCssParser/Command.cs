//------------------------------------------------------------------------------
// <copyright file="Command.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;
using System.Linq;
using System.Collections.Generic;

namespace InlineCssParser
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class Command
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;
        List<HtmlElement> list = new List<HtmlElement>();

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("82b0ea61-76c4-4c2c-bbf1-03ec5f8523c3");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="Command"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private Command(Package package)
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
        public static Command Instance
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
            Instance = new Command(package);
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
            DTE dte = Package.GetGlobalService(typeof(DTE)) as DTE;
            Document doc = dte.ActiveDocument;
            TextDocument txtDoc = doc.Object() as TextDocument;
            var text = txtDoc.CreateEditPoint(txtDoc.StartPoint).GetText(txtDoc.EndPoint);

            if (txtDoc.Language == "HTMLX" || txtDoc.Language == "HTML")
            {
                //// Show a message box to prove we were here
                //VsShellUtilities.ShowMessageBox(
                //    this.ServiceProvider,
                //    text,
                //    "File content",
                //    OLEMSGICON.OLEMSGICON_INFO,
                //    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                //    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

                var parsed = ParseHtml(text);
                var cssFileContent = string.Empty;

                foreach (var item in list)
                {
                    item.Id = string.IsNullOrEmpty(item.Id) ? "x" : item.Id;

                    var replaceText = string.Format("{0} id=\"{1}\" class=\"{2}\"", item.Name, item.Id, item.Id);
                    parsed = parsed.Replace(item.Guid, replaceText);
                    cssFileContent += string.Format(".{0}{{{1}}}\n\n", item.Id, "\n" + item.Style);
                }

                cssFileContent = cssFileContent.Replace(";", ";\n");

                //existing html file
                TextSelection txtSelHtml = (TextSelection)doc.Selection;
                txtSelHtml.SelectAll();
                txtSelHtml.Delete();
                txtSelHtml.Insert(parsed);

                //newly created css file
                string solutionDir = System.IO.Path.GetDirectoryName(dte.Solution.FullName);
                dte.ItemOperations.NewFile(@"General\Text File", "thereYouGo.css", EnvDTE.Constants.vsViewKindTextView);
                TextSelection txtSelCss = (TextSelection)dte.ActiveDocument.Selection;
                txtSelCss.SelectAll();
                txtSelCss.Delete();
                txtSelCss.Insert(cssFileContent);
            }
            else
            {
                VsShellUtilities.ShowMessageBox(
                    this.ServiceProvider,
                    "Invalid file!",
                    "ops!",
                    OLEMSGICON.OLEMSGICON_INFO,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
        }

        private string ParseHtml(string text)
        {
            int pointer = 0;
            var startTagIndex = 0;
            var endTagIndex = 0;

            while (text.Contains("; ")) //style tag i içerisindeki boşluklar trim ediliyor. alttaki split'i bozmasın diye
            {
                text = text.Replace("; ", ";");
            }

            while (pointer < text.Length || startTagIndex == -1 || endTagIndex == -1)
            {
                startTagIndex = text.IndexOf('<', pointer);
                endTagIndex = text.IndexOf('>', pointer);
                var elementText = text.Substring(startTagIndex + 1, (endTagIndex - (startTagIndex + 1)));

                if (elementText.Contains("style"))
                {
                    var parsedElement = elementText.Split(' ');

                    var elementName = parsedElement[0];
                    var elementId = string.Empty;
                    var elementStyle = string.Empty;
                    var guid = Guid.NewGuid().ToString();

                    var idAttr = parsedElement.FirstOrDefault(q => q.Contains("id"));
                    if (idAttr != null)
                    {
                        elementId = idAttr.Replace("id=", string.Empty).Replace("\"", string.Empty);
                    }

                    var styleAttr = parsedElement.FirstOrDefault(q => q.Contains("style"));
                    if (styleAttr != null)
                    {
                        elementStyle = styleAttr.Replace("style=", string.Empty).Replace("\"", string.Empty);
                    }

                    list.Add(new HtmlElement
                    {
                        Id = elementId,
                        Name = elementName,
                        Style = elementStyle,
                        Guid = guid
                    });

                    text = text.Replace(elementText, guid);
                    pointer = text.IndexOf('>', text.IndexOf(guid)) + 1;
                }
                else
                {
                    pointer = endTagIndex + 1;
                }
            }
            return text;
        }

        public class HtmlElement
        {
            public string Id { get; set; }
            public string Name { get; set; }
            public string Style { get; set; }
            public string Guid { get; set; }
        }
    }
}
