using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Промышленная_Безопасность
{
    internal enum ТипОбучения
    {
        кцн = 0,
        подг = 1,
        переподг = 2,
        пов = 3,
    }
    internal enum ТипПрофессии
    {
        КЦН_физ = 11,
        КЦН_юр = 12,
        Рабоч_физ = 21,
        Рабоч_юр = 21,
    }

    /// <summary>
    /// Хранит информацию обо всех документах
    /// </summary>
    class DirectoryMatcher
    {
        public DirectoryInfo Workdir { get; protected set; }
        public string DirectoryNameTemplate { get; protected set; }
        public string ProffecionName { get; protected set; }
        public FileInfo Форма { get; protected set; }
        public FileInfo Договор { get; protected set; }
        public FileInfo Приложение { get; protected set; }
        public FileInfo ПриказЗач { get; protected set; }
        public FileInfo ПриказОтч { get; protected set; }
        public FileInfo Протокол { get; protected set; }
        public FileInfo Карточка { get; protected set; }
        public FileInfo Счет { get; protected set; }

        private string GetDefaultPath()
        {
            //return @"F:\2023\"
            //return @"C:\Users\Yura\source\repos\testWord\testWord\bin\Debug\net5.0\documents\"
            //взять из настроек
            XmlDocument xmlDcoument = new XmlDocument();
            xmlDcoument.Load(@"AppSettings.xml");
            return xmlDcoument.SelectSingleNode("Settings").SelectSingleNode("DefaultPath").InnerText;
        }

        private string PathCombine(string filename) =>
            Path.Combine(Workdir.FullName, $"{filename}.docx");

        private void Initialize()
        {
            this.Форма = new FileInfo(PathCombine("Форма обложки дела"));
            this.Карточка = new FileInfo(PathCombine("Карточка"));
            this.Протокол = new FileInfo(PathCombine("Протокол"));
            this.Договор = new FileInfo(PathCombine("Договор"));
            this.Приложение = new FileInfo(PathCombine("Приложение к договору"));
            this.ПриказЗач = new FileInfo(PathCombine("Приказ о зачислении"));
            this.ПриказОтч = new FileInfo(PathCombine("Приказ об отчислении"));
            this.Счет = new FileInfo(PathCombine("Счет"));
        }

        /// <summary>
        /// На случай нового документа
        /// </summary>
        /// <param name="filename">Имя файла без расширения .docx</param>
        /// <returns>Новый FileInfo документа</returns>
        public FileInfo NewDocument(string filename) => new FileInfo(PathCombine(filename));

        public DirectoryMatcher(ТипПрофессии профессия, string название, int номер, DateTime дата, ТипОбучения? тип, int разряд, string организация)
        {
            название = string.Concat(название.Split('\"'));
            организация = string.Concat(организация.Split('\"'));

            string defaultPath = this.GetDefaultPath(); // "F:\2023\"
            this.ProffecionName = название.Trim();      // "Слесарь-ремонтник" или "Проведение работ на высоте"
            if (профессия != ТипПрофессии.КЦН_физ && профессия != ТипПрофессии.КЦН_юр)
            {
                // "7 (21.04-) пов 5р АК Скидельский"
                this.DirectoryNameTemplate = $"{номер} ({дата:dd.MM}-) {тип} {разряд}р {организация}";
            }
            else
            {
                this.DirectoryNameTemplate = $"{номер} ({дата:dd.MM}-) {организация}"; // "1 (20.04-) Барановичиэнергострой"
            }

            // "F:\2023\Слесарь-ремонтник\7 (21.04-) пов 5р АК Скидельский\"
            this.Workdir = new DirectoryInfo(Path.Combine(defaultPath, $"{ProffecionName}\\", $"{DirectoryNameTemplate}\\"));

            if (!this.Workdir.Exists)
            {
                this.Workdir.Create();
            }

            this.Initialize();
        }
    }
}
