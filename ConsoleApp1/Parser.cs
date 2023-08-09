using System;
using System.Text.RegularExpressions;
using Xceed.Words.NET;

public class WordToHtmlParser
{
    public static string ConvertToHtml(string filePath)
    {
        string styleHead = Class1.Head();
        using (DocX doc = DocX.Load(filePath))
        {
            string html = styleHead;
            string[] patterns = new string[]
            {
                @"_{2,}",   // Паттерн для поиска подчеркиваний (более 2 символов)
                @"[^_]"    // Паттерн для поиска строки, не содержащей подчеркивание
            };

            foreach (var paragraph in doc.Paragraphs)
            {
                string paragraphText = HtmlEncode(paragraph.Text);
                string formattedHtml = "";
                bool shouldAddBlock = false; // Флаг для определения, добавлять ли блок

                foreach (var pattern in patterns)
                {
                    try
                    {
                        Match match = Regex.Match(paragraphText, pattern);
                        if (match.Success)
                        {
                            switch (pattern)
                            {
                                case @"_{2,}":
                                    string replace = match.Value.Replace("_", ""); // Удаление подчеркиваний
                                    formattedHtml +=
                                        $"<table width=\"100%\">\r\n" +
                                        $"  <tr>\r\n" +
                                        $"    <td width=\"60%\">\r\n" +
                                        $"    </td>\r\n" +
                                        $"    <td class=\"under_line\" width=\"40%\">\r\n" +
                                        $"      <span style=\"border-bottom: 3px solid white;\">{replace}</span><br></td>\r\n" +
                                        $"  </tr>\r\n" +
                                        $"</table>";

                                    shouldAddBlock = true;
                                    break;

                                case @"[^_]":
                                    formattedHtml +=
                                        $"<table width=\"100%\">\r\n" +
                                        $"  <tr>\r\n" +
                                        $"    <td width=\"5\"></td>\r\n" +
                                        $"    <td width=\"95%\" style=\"text-indent:20px;\">{paragraphText}</td>\r\n" +
                                        $"  </tr>\r\n" +
                                        $"</table>";

                                    shouldAddBlock = true;
                                    break;
                            }

                            break; // Выход из цикла, так как найдено совпадение
                        }
                    }
                    catch (RegexParseException ex)
                    {
                        Console.WriteLine("Ошибка в регулярном выражении: " + ex.Message);
                    }
                }

                if (shouldAddBlock)
                {
                    html += formattedHtml;
                }
            }

            html += "</body></html>";
            return html;
        }
    }

    private static string HtmlEncode(string input)
    {
        return System.Web.HttpUtility.HtmlEncode(input);
    }
}
