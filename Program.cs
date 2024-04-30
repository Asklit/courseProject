using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

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

            // Блок парсинга документа
            if (body != null)
            {
                foreach (var element in body.Elements())
                {
                    if (element is Paragraph paragraph)
                    {
                        Console.WriteLine(paragraph.InnerText); // Вывод текста в консоль для проверки работы
                        TextParser(paragraph); // Запуск обработчика текста
                    }
                }
            }

            // Dispose the document handle. PS Не понимаю что делает эта строчка, но оставил ее
            wordprocessingDocument.Dispose();
        }

        static void TextParser(Paragraph paragraph)
        {
            // Функция обработки текста
        }
    }
}