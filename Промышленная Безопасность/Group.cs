using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Промышленная_Безопасность
{
    internal class Group
    {
        public string Группа;
        public string Шифр;
        public string Название;
        public string Организация;
        public ТипОбучения типОбучения;
        public ТипПрофессии типПрофессии;
        public int ПорядковыйНомерГруппы;
        public int Разряд;
        public string Год;
        public string ДатаОбучС;
        public string ДатаОбучПо;
        public string ДатаПрактС;
        public string ДатаПрактПо;
        public string ДатаКонсультации;
        public string ДатаЭкзамена;
        public string НомерПротокола;

        public Group() { }

        public Group(DataGridViewRow row)
        {
            try
            {
                Шифр = row.Cells["Шифр"].Value.ToString().Trim();
                Название = row.Cells["Название профессии"].Value.ToString();
                Организация = row.Cells["Организация"].Value.ToString();
                типОбучения = (ТипОбучения)int.Parse(row.Cells["Код занятия"].Value.ToString());
                if(string.IsNullOrEmpty(Организация) || Организация == "NULL")
                {
                    типПрофессии = типОбучения == ТипОбучения.кцн ? ТипПрофессии.КЦН_физ : ТипПрофессии.Рабоч_физ;
                }
                else
                {
                    типПрофессии = типОбучения == ТипОбучения.кцн ? ТипПрофессии.КЦН_юр : ТипПрофессии.Рабоч_юр;
                }
            
                ПорядковыйНомерГруппы = int.Parse(row.Cells["Порядковый номер группы"].Value.ToString());
                Разряд = int.Parse(row.Cells["Разряд"].Value.ToString());
                Год = row.Cells["Год"].Value.ToString();
                ДатаОбучС = row.Cells["Дата теории с"].Value.ToString();
                ДатаОбучПо = row.Cells["Дата теории по"].Value.ToString();
                ДатаПрактС = row.Cells["Дата практики с"].Value.ToString();
                ДатаПрактПо = row.Cells["Дата практики по"].Value.ToString();
                ДатаКонсультации = row.Cells["Дата консультации"].Value.ToString();
                ДатаЭкзамена = row.Cells["Дата экзамена"].Value.ToString();
                НомерПротокола = row.Cells["Номер протокола"].Value.ToString();

                Группа = типОбучения == ТипОбучения.кцн ? 
                    $"{Шифр}.{ПорядковыйНомерГруппы}.{Год}" : 
                    $"{Шифр}.{(int)типОбучения}.{ПорядковыйНомерГруппы}.{Год}";
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }
    }
}
