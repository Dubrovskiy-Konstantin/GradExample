using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Промышленная_Безопасность
{
    /// <summary>
    /// Создает документ Word меняя тэги из шаблона на нужные значения
    /// </summary>
    class WordCreator
    {
        private readonly FileInfo _fileTemplateInfo;
        private readonly FileInfo _fileResultInfo;

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="fileTemplate">Имя файла шаблона документа Word</param>
        /// <param name="fileResult">Имя готового документа Word. (Если существует, будет пересоздан)</param>
        public WordCreator(FileInfo fileTemplate, FileInfo fileResult)
        {
            if (fileTemplate.Exists)
            {
                _fileTemplateInfo = fileTemplate;
                _fileResultInfo = fileResult;
                if (fileResult.Exists)
                {
                    _fileResultInfo.Delete();
                }
            }
            else
            {
                throw new FileNotFoundException($"Не найден файл шаблона по пути \"{fileTemplate.FullName}\"");
            }
        }

        /// <summary>
        /// Заменить все тэги внутри шаблона Word на необходимые значения
        /// </summary>
        /// <param name="items">Пара ключ-значение. Ключом является тэг внутри шаблона Word. Значение - то, что будет вставлено вместо тэга </param>
        /// <returns>true если успешно, false если ошибка</returns>
        public bool Process(Dictionary<string, string> items)
        {
            Word.Application app = null;
            try
            {
                app = new Word.Application();
                Object file = _fileTemplateInfo.FullName;
                Object missing = Type.Missing;

                app.Documents.Open(file);
                foreach (var item in items)
                {
                    Word.Find find = app.Selection.Find;
                    find.Text = item.Key;
                    find.Replacement.Text = item.Value;

                    Object wrap = Word.WdFindWrap.wdFindContinue;
                    Object replace = Word.WdReplace.wdReplaceAll;

                    find.Execute(FindText: Type.Missing,
                        MatchCase: false,
                        MatchWholeWord: false,
                        MatchWildcards: false,
                        MatchSoundsLike: missing,
                        MatchAllWordForms: false,
                        Forward: true,
                        Wrap: wrap,
                        Format: false,
                        ReplaceWith: missing,
                        Replace: replace
                        );
                }

                Object resultFileName = _fileResultInfo.FullName;
                app.ActiveDocument.SaveAs2(resultFileName);
                app.ActiveDocument.Close();

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
            finally
            {
                app?.Quit();
            }
        }
    }
}
