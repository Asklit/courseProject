using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Xml;
using System.Xml.Linq;

namespace course_project
{
    internal class FileProcessing
    {
        /// <summary>
        /// Путь до файла
        /// </summary>
        const string path = @"..\\..\\..\\WordForTest.docx";

        static void Main() 
        {
            // Создание потока и записи файла по пути path
            Stream stream = File.Open(path, FileMode.Open);

            OpenAndAddToWordprocessingStream(stream);

            // Закрытие файла. PS Важно закрыть файл, так как иначе он повредится и открыть его будет невозвожно!
            stream.Close();
        }

        static void OpenAndAddToWordprocessingStream(Stream stream)
        {
            // Открытие WordProcessingDocument из потока.
            WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(stream, true);

            // Присвоить ссылку на существующий текст документа.
            MainDocumentPart mainDocumentPart = wordprocessingDocument.MainDocumentPart ?? wordprocessingDocument.AddMainDocumentPart();
            mainDocumentPart.Document ??= new Document();
            Body body = mainDocumentPart.Document.Body ?? mainDocumentPart.Document.AppendChild(new Body());

            // Блок парсинга текста документа (не уверен что этот блок нужен будет на релизе, так как стили содержатся в другом месте)
            if (body != null)
            {
                foreach (var element in body.Elements())
                {
                    if (element is Paragraph paragraph)
                    {
                        Console.WriteLine(paragraph.InnerText); // Вывод текста в консоль для проверки работы
                    }
                }
            }

            // Проверка на пустоту стилей
            if (mainDocumentPart is null || mainDocumentPart.StyleDefinitionsPart is null || mainDocumentPart.StylesWithEffectsPart is null)
            {
                // StylesWithEffectsPart из за чего то равен null, поэтому пока закоментил эту проверку
                // throw new ArgumentNullException("MainDocumentPart and/or one or both of the Styles parts is null.");
            }

            XDocument? styles = null;
            StylesPart? stylesPart = null;

            // Я пока не знаю какой метод из двух использовать нужно будет, посмотрим позже PS StylesWithEffectsPart из за чего то равен null,
            // поэтому как использовать пока не разобрался 
            // stylesPart = mainDocumentPart.StylesWithEffectsPart;
            stylesPart = mainDocumentPart.StyleDefinitionsPart;

            using (var reader = XmlNodeReader.Create(
              stylesPart.GetStream(FileMode.Open, FileAccess.Read)))
            {
                // Создание XDocument.
                styles = XDocument.Load(reader);
            }

            // Вывод стилей
            IEnumerable<XElement> elements =
                from el in styles.Elements()
                select el;
            foreach (XElement el in elements)
                Console.WriteLine(el);

            // Dispose the document handle. PS Не понимаю что делает эта строчка, но оставил ее
            wordprocessingDocument.Dispose();
        }

        static void TextParser(Paragraph paragraph)
        {
            // Функция обработки текста
        }
    }
}