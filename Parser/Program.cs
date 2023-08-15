using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;

class Program
{
    static void Main(string[] args)
    {
        string inputPath = "new.docx";
        string outputPath = "output.html";

        using (WordprocessingDocument doc = WordprocessingDocument.Open(inputPath, false))
        {
            // Извлечение содержимого из документа Word
            var content = ExtractContent(doc.MainDocumentPart.Document.Body);

            // Создание HTML-кода из содержимого
            string html = GenerateHtmlFromContent(content);

            // Сохранение HTML-кода в файл
            File.WriteAllText(outputPath, html);
        }
    }

    // Рекурсивная функция для извлечения содержимого из элементов Open XML
    static string ExtractContent(OpenXmlElement element)
    {
        string content = "";

        foreach (var childElement in element.Elements())
        {
            if (childElement is Text)
            {
                content += childElement.InnerText;
            }
            else if (childElement is Run)
            {
                content += ExtractContent(childElement);
            }
            else if (childElement is Paragraph)
            {
                content += "<p>" + ExtractContent(childElement) + "</p>";
            }
            else if (childElement is Table)
            {
                content += "<table>" + ExtractTableContent(childElement as Table) + "</table>";
            }
            // Добавьте обработку других типов элементов по аналогии
        }

        return content;
    }

    // Извлечение содержимого таблицы
    static string ExtractTableContent(Table table)
    {
        string tableContent = "";

        foreach (var row in table.Elements<TableRow>())
        {
            tableContent += "<tr>";
            foreach (var cell in row.Elements<TableCell>())
            {
                tableContent += "<td>" + ExtractContent(cell) + "</td>";
            }
            tableContent += "</tr>";
        }

        return tableContent;
    }

    // Генерация HTML-кода на основе извлеченного содержимого
    static string GenerateHtmlFromContent(string content)
    {
        return $"<html><head><title>Converted Word to HTML</title></head><body>{content}</body></html>";
    }
}
