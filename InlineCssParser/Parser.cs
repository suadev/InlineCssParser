using EnvDTE;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Linq;

namespace InlineCssParser
{
    public class Parser
    {
        public string ParseHtml(string text, List<HtmlElement> elementList, TextDocument txtDoc, IVsStatusbar bar, ref uint cookie)
        {
            int pointer = 0;
            var startTagIndex = 0;
            var endTagIndex = 0;
            var endTagBefore = 0;
            uint complete = 0;
            uint total = (uint)text.Count(q => q == '<');

            while (text.Contains("; ") || text.Contains(": ")) //style tag i içerisindeki boşluklar trim ediliyor. alttaki split'i bozmasın diye
            {
                text = text.Replace("; ", ";").Replace(": ", ":");
            }

            startTagIndex = text.IndexOf('<', pointer);
            endTagIndex = text.IndexOf('>', pointer);

            try
            {
                txtDoc.Selection.StartOfDocument();
                while (pointer < text.Length && startTagIndex != -1 || endTagIndex != -1)
                {
                    complete++;
                    bar.Progress(ref cookie, 1, "", complete, total);
                    bar.SetText("Extracting inline styles ...");

                    //current line ı bulabilmek için cursoru dolastırıyoruz. text üzerinde bir değişiklik yapılmıyor
                    txtDoc.Selection.CharRight(false, endTagIndex - endTagBefore);
                    endTagBefore = endTagIndex;

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

                        var idAttr = parsedElement.FirstOrDefault(q => q.Contains("id="));
                        if (idAttr != null)
                        {
                            elementId = idAttr.Replace("id=", string.Empty).Replace("\"", string.Empty);
                        }
                        #endregion

                        #region checking style attr

                        var styleAttr = parsedElement.FirstOrDefault(q => q.Contains("style="));
                        if (styleAttr != null)
                        {
                            elementStyle = styleAttr.Replace("style=", string.Empty).Replace("\"", string.Empty);
                        }
                        #endregion

                        #region check important style
                        var importantAttr = parsedElement.FirstOrDefault(q => q.Contains("!important"));
                        if (importantAttr != null)
                        {
                            elementStyle += string.Format(" {0}", importantAttr.Replace("\"", string.Empty).Trim());
                        }
                        #endregion

                        #region checking class attr

                        var classAttr = parsedElement.FirstOrDefault(q => q.Contains("class="));
                        if (classAttr != null)
                        {
                            elementClass = classAttr.Replace("class=", string.Empty).Replace("\"", string.Empty);
                        }

                        #endregion

                        text = text.Replace(elementText, guid);
                        pointer = text.IndexOf('>', text.IndexOf(guid)) + 1;

                        elementList.Add(new HtmlElement
                        {
                            Id = elementId,
                            Name = elementName,
                            Style = elementStyle,
                            Class = elementClass,
                            Guid = guid,
                            LineNumber = txtDoc.Selection.CurrentLine
                        });
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