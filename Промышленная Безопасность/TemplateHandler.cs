using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using System.Windows.Forms;

namespace Промышленная_Безопасность
{
    /// <summary>
    /// Класс шаблона документа
    /// </summary>
    internal class Template
    {
        /// <summary>
        /// Объект шаблона
        /// </summary>
        public FileInfo File { get; set; }
        /// <summary>
        /// Таблица из трех столбцов (тег, описание, значение)
        /// </summary>
        public List<(string Tag, string Descriptoin, string Value)> Table { get; set; } 
        /// <summary>
        /// Добавил ли пользователь файл себе в активную таблицу
        /// </summary>
        public bool IsSelectetByUser { get; set; } 

        public Template(FileInfo file)
        {
            this.File = file;
            this.Table = new List<(string Tag, string Descriptoin, string Value)>();
        }

        /// <summary>
        /// Метод генерации данных для последующей работы класса по замене тегов на значения
        /// </summary>
        /// <returns>Пары тег-значение</returns>
        public Dictionary<string, string> GetTagValue()
        {
            var res = new Dictionary<string, string>();
            foreach (var line in Table)
            {
                res.Add(line.Tag, line.Value);
            }

            return res;
        }
    }

    /// <summary>
    /// Загружает и хранит все шаблоны документов
    /// </summary>
    internal class TemplateHandler
    {
        private DirectoryInfo _templateFolder;
        private FileInfo _csvFile;

        /// <summary>
        /// Словарь объектов шаблонов документов
        /// </summary>
        public Dictionary<string, Template> Templates { get; private set; } = new Dictionary<string, Template>();

        public TemplateHandler()
        {
            try
            {
                XmlDocument xmlDcoument = new XmlDocument();
                xmlDcoument.Load(@"AppSettings.xml");
                _templateFolder = new DirectoryInfo(xmlDcoument.SelectSingleNode("Settings").SelectSingleNode("TemplateFolder").InnerText);
                if (!_templateFolder.Exists)
                {
                    MessageBox.Show($"Не удалось найти папку с шаблонами по пути {_templateFolder}");
                    return;
                }

                _csvFile = _templateFolder.GetFiles().FirstOrDefault(x => x.Name == "Документы.txt");
                if (_csvFile is null)
                {
                    MessageBox.Show($"Не удалось найти csv файл 'Документы.csv' в папке с шаблонами по пути {_templateFolder}");
                    return;
                }

                using (StreamReader reader = new StreamReader(_csvFile.OpenRead(), Encoding.UTF8))
                {
                    bool err = false;
                    string currentFileName = "!nofile";
                    for (string line = reader.ReadLine(); line is not null; line = reader.ReadLine())
                    {
                        if (line.StartsWith('!'))
                        {
                            currentFileName = line[1..^1];
                            FileInfo currentFile = _templateFolder.GetFiles().FirstOrDefault(x => x.Name == currentFileName);
                            if (currentFile is null)
                            {
                                MessageBox.Show($"Не найден файл {currentFileName}");
                                err = true;
                                break;
                            }
                            else
                            {
                                this.Templates.Add(currentFileName, new Template(currentFile));
                            }
                        }
                        else
                        {
                            var cols = line.Split(';');
                            this.Templates[currentFileName].Table.Add((cols[0], cols[1], ""));
                        }
                    }
                    if (err)
                    {
                        MessageBox.Show("Ошибка в поиске шаблонов файлов");
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
