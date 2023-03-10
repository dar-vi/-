using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Serialization;
using OfficeOpenXml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentConverter
{
    public partial class MainForm : Form
    {

        private OpenFileDialog openFileDialog1 = new OpenFileDialog();
        private SaveFileDialog saveFileDialog1 = new SaveFileDialog();
        public class Document
        {
            public string Name { get; set; }
            public string StartDate { get; set; }
            public string DeadLine { get; set; }
        }
        public static class ReadDataFromCsv
        {
            public static List<Document> Read(string fileName)
            {
                using (var reader = new StreamReader(fileName))
                using (var csv = new CsvHelper.CsvReader(reader, System.Globalization.CultureInfo.InvariantCulture))
                {
                    return csv.GetRecords<Document>().ToList();
                }
            }
        }

        public static class ReadDataFromXml
        {
            public static List<Document> Read(string fileName)
            {
                var serializer = new XmlSerializer(typeof(List<Document>));
                using (var reader = new StreamReader(fileName))
                {
                    return (List<Document>)serializer.Deserialize(reader);
                }
            }
        }
        public static class ReadDataFromJson
        {
            public static List<Document> Read(string fileName)
            {
                using (var reader = new StreamReader(fileName))
                {
                    string json = reader.ReadToEnd();
                    return JsonConvert.DeserializeObject<List< Document>>(json);
                }
            }
        }

        public static class ReadDataFromExcel
        {
            public static List<Document> Read(string fileName)
            {
                var documents = new List<Document>();
                using (var package = new ExcelPackage(new FileInfo(fileName)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)
                    {
                        var doc = new Document
                        {
                            Name = worksheet.Cells[i, 1].Value.ToString().Trim(),
                            StartDate = (DateTime)worksheet.Cells[i, 2].Value,
                            Deadline = (DateTime)worksheet.Cells[i, 3].Value
                        };
                        documents.Add(doc);
                    }
                }
                return documents;
            }
        }
        List<Document> documents = new List<Document>();

        private void MainForm_Load(object sender, EventArgs e)
        {
            // Настройка диалоговых окон
            openFileDialog1.Filter = "JSON files (*.json)|*.json|CSV files (*.csv)|*.csv|XML files (*.xml)|*.xml|XLSX files (*.xlsx)|*.xlsx";
            saveFileDialog1.Filter = "JSON files (*.json)|*.json|CSV files (*.csv)|*.csv|XML files (*.xml)|*.xml|XLSX files (*.xlsx)|*.xlsx";
        }


        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            // Открытие диалогового окна выбора файла
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                // Чтение данных из исходного файла "Document"
                string fileName = openFileDialog1.FileName;
                string extension = Path.GetExtension(fileName);

                switch (extension)
                {
                    case ".json":
                        documents = ReadDataFromJson(fileName);
                        break;
                    case ".csv":
                        documents = ReadDataFromCsv(fileName);
                        break;
                    case ".xml":
                        documents = ReadDataFromXml(fileName);
                        break;
                    case ".xlsx":
                        documents = ReadDataFromExcel(fileName);
                        break;
                    default:
                        MessageBox.Show("Unsupported file format");
                        return;
                }
            }
        }
               
        private void btnConvert_Click(object sender, EventArgs e)
        {
            // Проверка наличия данных
            if (documents.Count == 0)
            {
                MessageBox.Show("No data to convert");
                return;
            }

            // Открытие диалогового окна настроек конвертации
            ConversionSettingsForm settingsForm = new ConversionSettingsForm();
            DialogResult result = settingsForm.ShowDialog();

            if (result == DialogResult.OK)
            {
                // Получение выбранных настроек конвертации
                ConversionSettings settings = settingsForm.GetSettings();

                // Открытие диалогового окна выбора места сохранения
                DialogResult saveResult = saveFileDialog1.ShowDialog();
                if (saveResult == DialogResult.OK)
                {
                    // Конвертация и сохранение данных в выбранный формат
                    string fileName = saveFileDialog1.FileName;
                    string extension = Path.GetExtension(fileName);
                    var format = "";
                    switch (extension)
                    {
                        case ".json":
                            format = "json";
                            break;
                        case ".csv":
                            format = "csv";
                            break;
                        case ".xml":
                            format = "xml";
                            break;
                        case ".xlsx":
                            format = "xlsx";
                            break;

                    }
                    if (format == "json")
                    {
                        var jsonSerializerSettings = new JsonSerializerSettings
                        {
                            ContractResolver = new CamelCasePropertyNamesContractResolver(),
                            Formatting = Formatting.Indented
                        };

                        var jsonData = JsonConvert.SerializeObject(documents, jsonSerializerSettings);

                        using (var writer = new StreamWriter(saveFileDialog.FileName))
                        {
                            writer.Write(jsonData);
                        }
                    }
                    else if (format == "csv")
                    {
                        using (var writer = new StreamWriter(saveFileDialog.FileName))
                        using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                        {
                            csv.WriteRecords(documents);
                        }
                    }
                    else if (format == "xml")
                    {
                        var xmlSerializer = new XmlSerializer(typeof(List<Document>));

                        using (var writer = new StreamWriter(saveFileDialog.FileName))
                        {
                            xmlSerializer.Serialize(writer, documents);
                        }
                    }
                    else if (format == "xlsx")
                    {
                        using (var workbook = new XLWorkbook())
                        {
                            var worksheet = workbook.Worksheets.Add("Documents");
                            worksheet.Cell(1, 1).Value = "Name";
                            worksheet.Cell(1, 2).Value = "StartDate";
                            worksheet.Cell(1, 3).Value = "Deadline";

                            for (int i = 0; i < documents.Count; i++)
                            {
                                worksheet.Cell(i + 2, 1).Value = documents[i].Name;
                                worksheet.Cell(i + 2, 2).Value = documents[i].StartDate;
                                worksheet.Cell(i + 2, 3).Value = documents[i].Deadline;
                            }

                            workbook.SaveAs(saveFileDialog.FileName);
                        }
                    }
                }
                MessageBox.Show("Конвертация завершена успешно!");
            }
        }
        private void openFileDialog1(object sender, EventArgs e)
        {

        }

        private void saveFileDialog1(object sender, EventArgs e)
        {

        }
    }
}





