using System.Text.RegularExpressions;
using System.Web;
using Xceed.Words.NET;

namespace Parser;

public static class WordToHtmlParser
{
    public static string ConvertToHtml(string filePath)
    {
        using (DocX doc = DocX.Load(filePath))
        {
            Regex patternGreaterThan20WithUnderscore =
                new Regex(@"^(?=.*_).{21,}$"); //символов > 20 и есть подчеркивание:
            Regex patternAnyWithUnderscore = new Regex(@"^.*_.*$"); //есть подчеркивание:
            Regex patternLessThan20WithUnderscore =
                new Regex(@"^(?=.*_).{1,19}$"); //символов < 20 и есть подчеркивание:

            Regex
                patternGreaterThan20WithoutUnderscore =
                    new Regex(@"^(?!.*_).{105,}$"); // символов > 105 и нет подчеркивания:

            Regex patternOnlyUnderscore = new Regex(@"^_+$"); //есть только подчеркивани

            var html = HtmlStrings.Head();


            foreach (var paragraph in doc.Paragraphs)
            {
                var paragraphText = HtmlEncode(paragraph.Text);
                var formattedLine = "";
                var replace = "";
                if (
                    //символов > 20 и есть подчеркивание:
                    patternGreaterThan20WithUnderscore.IsMatch(paragraphText))
                {
                    replace = paragraphText.Replace("_", ""); // Удаление подчеркиваний
                    formattedLine +=
                        $"<table width=\"100%\">\r\n" +
                        $"  <tr>\r\n" +
                        $"    <td width=\"60%\">\r\n" +
                        $"    </td>\r\n" +
                        $"    <td class=\"under_line\" width=\"40%\">\r\n" +
                        $"      <span style=\"border-bottom: 3px solid white;\">{replace}</span><br></td>\r\n" +
                        $"  </tr>\r\n" +
                        $"</table>";
                    html += formattedLine;
                }
                //есть только подчеркивание
                else if (patternOnlyUnderscore.IsMatch(paragraphText))
                {
                    replace = paragraphText.Replace("_", "");
                    formattedLine += $@"
                                <table width=""100%"">
                                    <tr>
                                        <td width=""2%""></td>
                                        <td width=""98%"" style=""text-indent: 20px;"" class=""under_line"">
                                            {replace}
                                        </td>
                                    </tr>
                                </table>";
                    html += formattedLine;
                }
                // символов больше 20 и нет подчеркивания:
                else if (patternGreaterThan20WithoutUnderscore.IsMatch(paragraphText))
                {
                    //TODO
                    formattedLine += $@"
                                    <table width=""100%"">
                                        <tr>
                                            <td width=""100%"" style=""text-indent:20px;"">{paragraphText}</td>
                                        </tr>
                                    </table>";
                    html += formattedLine;
                }
                ////символов < 20 и есть подчеркивание:
                else if (patternLessThan20WithUnderscore.IsMatch(paragraphText))
                {
                    //TODO
                }


                html += formattedLine;
            }

            html += "</body></html>";
            return html;
        }
    }

    private static string HtmlEncode(string input)
    {
        return HttpUtility.HtmlEncode(input);
    }
}