using System;
using System.ComponentModel.Design;
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
                var elementList = new List<HtmlElement>();
                var parsed = ParseHtml(text, elementList, txtDoc);
                var cssFileContent = string.Empty;

                if (elementList.Any())
                {
                    foreach (var item in elementList)
                    {
                        var cssClass = string.IsNullOrEmpty(item.Class) ? (string.IsNullOrEmpty(item.Id) ? CreateUniqueElementKey(item.Name, item.LineNumber) : item.Id) : item.Class;

                        var idAttr = string.IsNullOrEmpty(item.Id) ? string.Empty : string.Format("id=\"{0}\"", item.Id);

                        var replaceText = string.Format("{0} {1} class=\"{2}\"", item.Name, idAttr, cssClass);

                        parsed = parsed.Replace(item.Guid, replaceText);

                        cssFileContent += string.Format(".{0}{{{1}}}\n\n", cssClass, "\n" + item.Style);
                    }

                    //css file beautification
                    cssFileContent = cssFileContent.Replace(";", ";\n");

                    //existing html file
                    TextSelection txtSelHtml = (TextSelection)doc.Selection;
                    txtSelHtml.SelectAll();
                    txtSelHtml.Delete();
                    txtSelHtml.Insert(parsed);

                    //newly created css file
                    var docName = doc.Name.Substring(0, doc.Name.IndexOf('.'));
                    docName = string.Format("{0}.css", docName);
                    string solutionDir = System.IO.Path.GetDirectoryName(dte.Solution.FullName);
                    dte.ItemOperations.NewFile(@"General\Text File", docName, EnvDTE.Constants.vsViewKindTextView);
                    TextSelection txtSelCss = (TextSelection)dte.ActiveDocument.Selection;
                    txtSelCss.SelectAll();
                    txtSelCss.Delete();
                    txtSelCss.Insert(cssFileContent);
                }
                else
                {
                    VsShellUtilities.ShowMessageBox(this.ServiceProvider, "Not found inline css!", "Oops!",
                        OLEMSGICON.OLEMSGICON_INFO,
                        OLEMSGBUTTON.OLEMSGBUTTON_OK,
                        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                }
            }
            else
            {
                VsShellUtilities.ShowMessageBox(this.ServiceProvider, "This is not a html file!", "Oops!",
                    OLEMSGICON.OLEMSGICON_WARNING,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
        }

        private string CreateUniqueElementKey(string name, int lineNumber)
        {
            return string.Format("{0}_Line{1}", name, lineNumber);
        }

        private string ParseHtml(string text, List<HtmlElement> elementList, TextDocument txtDoc)
        {
            int pointer = 0;
            var startTagIndex = 0;
            var endTagIndex = 0;

            while (text.Contains("; ") || text.Contains(": ")) //style tag i içerisindeki boşluklar trim ediliyor. alttaki split'i bozmasın diye
            {
                text = text.Replace("; ", ";").Replace(": ", ":");
            }

            startTagIndex = text.IndexOf('<', pointer);
            endTagIndex = text.IndexOf('>', pointer);

            try
            {
                while (pointer < text.Length && startTagIndex != -1 || endTagIndex != -1)
                {
                    //txtDoc.Selection.MoveToPoint();

                    var elementText = text.Substring(startTagIndex + 1, (endTagIndex - (startTagIndex + 1)));

                    if (elementText.Contains("style=")) // '=' is very important
                    {
                        var parsedElement = elementText.Split(' ');
                        var elementName = parsedElement[0];
                        var elementId = string.Empty;
                        var elementStyle = string.Empty;
                        var elementClass = string.Empty;
                        var guid = Guid.NewGuid().ToString();

                        #region checking id attr

                        var idAttr = parsedElement.FirstOrDefault(q => q.Contains("id"));
                        if (idAttr != null)
                        {
                            elementId = idAttr.Replace("id=", string.Empty).Replace("\"", string.Empty);
                        }
                        #endregion

                        #region checking style attr

                        var styleAttr = parsedElement.FirstOrDefault(q => q.Contains("style"));
                        if (styleAttr != null)
                        {
                            elementStyle = styleAttr.Replace("style=", string.Empty).Replace("\"", string.Empty);
                        }
                        #endregion

                        #region checking class attr

                        var classAttr = parsedElement.FirstOrDefault(q => q.Contains("class"));
                        if (classAttr != null)
                        {
                            elementClass = classAttr.Replace("class=", string.Empty).Replace("\"", string.Empty);
                        }

                        #endregion

                        elementList.Add(new HtmlElement
                        {
                            Id = elementId,
                            Name = elementName,
                            Style = elementStyle,
                            Class = elementClass,
                            Guid = guid,
                            LineNumber = txtDoc.Selection.CurrentLine
                        });

                        text = text.Replace(elementText, guid);
                        pointer = text.IndexOf('>', text.IndexOf(guid)) + 1;
                    }
                    else
                    {
                        pointer = endTagIndex + 1;
                    }

                    startTagIndex = text.IndexOf('<', pointer);
                    endTagIndex = text.IndexOf('>', pointer);
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return text;
        }
    }
}