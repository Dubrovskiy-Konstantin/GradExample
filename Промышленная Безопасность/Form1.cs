using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Промышленная_Безопасность
{
    public partial class Form1 : Form
    {
        private string _accountName;
        private bool _adminPremissions;
        private DBContext _dbContext;
        public delegate void ShowAncestor();
        private ShowAncestor _showAncestor;
        private Dictionary<string, string> _querries;
        private DataTable _currentTable;
        private List<Group> _groups = new List<Group>();
        private TemplateHandler _templateHandler;
        private DirectoryMatcher _directoryMatcher;
        private List<(DateTime Date, string Event)> _monthEvents = new();
        private CultureInfo _currentCulture = new CultureInfo("RU-ru", false);

        public Form1(ShowAncestor showAncestor, string accountName, bool admin = false)
        {
            this._accountName = accountName;
            this._adminPremissions = admin;
            this._showAncestor = showAncestor;
            this._dbContext = new DBContext();
            this._querries = new Dictionary<string, string>()
            {
                { "Группы", "select [Профессия].[Название] as [Название профессии], [Профессия].[Шифр], " +
                                "[Код занятия], [Организация].[Название] as [Организация], [Разряд], [Порядковый номер группы], " +
                                "[Год], [Дата теории с], [Дата теории по], [Дата практики с],[Дата практики по], " +
                                "[Дата консультации], [Дата экзамена], [Комиссия].[Дата] as [Дата комиссии], " +
                                "[Комиссия].[Номер протокола] from [Группа] " +
                                "left join[Профессия] on[Профессия].[Id] =[Группа].[Id Профессии] " +
                                "left join[Организация] on[Организация].[Id]=[Группа].[Id Организации] " +
                                "left join[Комиссия] on[Комиссия].[Id] =[Группа].[Id Комиссии] " },
                { "Слушатели",  "select [ФИО], pr.[Шифр], gr.[Код занятия], gr.[Порядковый номер группы], " +
                                    "gr.[Год], [Год рождения], [Номер паспорта], [Образование], " +
                                    "[Номер бланка об образовании] from [Слушатель] person " +
                                    "left join [Группа] gr on gr.[Id]=person.[Id Группы] " +
                                    "left join [Профессия] pr on gr.[Id Профессии]=pr.Id" },
                { "Профессии", "select [Название], [Шифр] from [Профессия]" },
                { "Организации", "select [Название] from [Организация]" }
            };
            InitializeComponent();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            _showAncestor.Invoke();
        }

        private void открытьОкноНастроекToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void включитьСпециальныеВозможностиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new Spec().Show();
        }

        private void информацияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!_adminPremissions)
            {
                MessageBox.Show("Окно специальных возможностей не доступно для вашего аккаунта");
            }
            else
            {
                MessageBox.Show("Просмотр информации об аккаунтах");
            }
        }

        #region Данные дизайн
        private void SetData(string sql)
        {            
            if (!string.IsNullOrEmpty(sql))
            {
                _currentTable = _dbContext.GetDataSet(sql).Tables[0];
                dataGridView1.Columns.Clear();
                dataGridView1.DataSource = _currentTable;
                if (dataGridView1.Columns.Count > 0) 
                    dataGridView1.Columns[0].Width = 300;
                if(dataGridView1.Rows.Count > 0 && dataGridView1.Columns.Count > 0 
                    && dataGridView1.Columns.Contains("Шифр")
                    && dataGridView1.Columns.Contains("Код занятия")
                    && dataGridView1.Columns.Contains("Порядковый номер группы")
                    && dataGridView1.Columns.Contains("Год")
                    ) //Значит таблица Группы/Слушатели
                {
                    //Вставить полный номер группы
                    DataGridViewColumn col;
                    try
                    {
                        DataGridViewCell cell = dataGridView1.Rows[0].Cells[0].Clone() as DataGridViewCell;
                        col = new DataGridViewColumn(cell);
                        col.Name = "Номер группы";
                        col.HeaderText = "Номер группы";
                        col.Width = 200;
                    }
                    catch (Exception ex) 
                    { 
                        MessageBox.Show(ex.Message); 
                        return; 
                    }

                    try
                    {
                        dataGridView1.Columns.Insert(0, col);
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            string a, b, c, d;
                            a = dataGridView1.Rows[i].Cells["Шифр"].Value.ToString();
                            b = dataGridView1.Rows[i].Cells["Код занятия"].Value.ToString();
                            c = dataGridView1.Rows[i].Cells["Порядковый номер группы"].Value.ToString();
                            d = dataGridView1.Rows[i].Cells["Год"].Value.ToString();
                            dataGridView1.Rows[i].Cells[0].Value = $"{a}.{b}.{c}.{d}";
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        dataGridView1.Columns.RemoveAt(0);
                    }
                }
            }                
        }

        private void comboBoxВыбратьТаблицу_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBoxВыбратьТаблицу.SelectedIndex)
            {
                case 0: //Слушатели
                    {
                        SetData(_querries["Слушатели"]);
                        comboBoxПоискПоПолю.Items.Clear();
                        comboBoxПоискПоПолю.SelectedIndex = -1;
                        comboBoxПоискПоПолю.Text = "";
                        textBoxПоискПоЗначениюПоля.Text = "";
                        comboBoxПоискПоПолю.Items.AddRange(new string[] {
                            "Полный номер группы",
                            "ФИО",
                            "Шифр",
                            "Номер паспорта",
                            "Номер бланка об образовании",
                        });
                        break;
                    }
                case 1: //Организации
                    {
                        SetData(_querries["Организации"]);
                        comboBoxПоискПоПолю.Items.Clear();
                        comboBoxПоискПоПолю.SelectedIndex = -1;
                        comboBoxПоискПоПолю.Text = "";
                        textBoxПоискПоЗначениюПоля.Text = "";
                        comboBoxПоискПоПолю.Items.AddRange(new string[] {
                            "Название",
                        });
                        break;
                    }
                case 2: //Профессии
                    {
                        SetData(_querries["Профессии"]);
                        comboBoxПоискПоПолю.Items.Clear();
                        comboBoxПоискПоПолю.SelectedIndex = -1;
                        comboBoxПоискПоПолю.Text = "";
                        textBoxПоискПоЗначениюПоля.Text = "";
                        comboBoxПоискПоПолю.Items.AddRange(new string[] {
                            "Название",
                            "Шифр",
                        });
                        break;
                    }
                case 3: //Группы
                    {
                        SetData(_querries["Группы"]);
                        comboBoxПоискПоПолю.Items.Clear();
                        comboBoxПоискПоПолю.SelectedIndex = -1;
                        comboBoxПоискПоПолю.Text = "";
                        textBoxПоискПоЗначениюПоля.Text = "";
                        comboBoxПоискПоПолю.Items.AddRange(new string[] {
                            "Полный номер группы",
                            "Название профессии",
                            "Шифр",
                            "Организация",
                            "Дата теории с",
                            "Дата теории по",
                            "Дата практики с",
                            "Дата практики по",
                            "Дата консультации",
                            "Дата экзамена",
                            "Номер протокола",
                        });
                        break;
                    }
                case -1:
                default:
                    {
                        comboBoxПоискПоПолю.Items.Clear();
                        break;
                    }
            }
        }

        #endregion

        #region Слушатель дизайн
        private void buttonНоваяГруппа_Click(object sender, EventArgs e)
        {
            this.groupBoxГруппа.Visible = false;
            this.groupBoxНоваяГруппа.Visible = true;
        }

        private void buttonСуществующаяГруппа_Click(object sender, EventArgs e)
        {
            this.groupBoxГруппа.Visible = false;
            this.groupBoxСуществующаяГруппа.Visible = true;
        }

        private void buttonСуществующаяГруппаНазад_Click(object sender, EventArgs e)
        {
            this.groupBoxГруппа.Visible = true;
            this.groupBoxСуществующаяГруппа.Visible = false;
        }

        private void buttonНоваяГруппаНазад_Click(object sender, EventArgs e)
        {
            this.groupBoxГруппа.Visible = true;
            this.groupBoxНоваяГруппа.Visible = false;
        }

        private void ПроверкаСлушателя(object sender, EventArgs e)
        {
            if (this.groupBoxСуществующаяГруппа.Visible)
            {
                this.textBoxПроверкаСлушателя.Text = $"Слушатель {textBoxФИО.Text} из организации {textBoxОрганизация.Text} добавляется в группу {textBoxСуществующаяГруппа.Text}";
            }
            else if (this.groupBoxНоваяГруппа.Visible)
            {
                string a = "0000-000";
                int b = 0, c = 0, d = 0;

                b = comboBoxЗанятия.SelectedIndex + 1;
                c = Convert.ToInt32(numericUpDownПорядкоывйНомерГруппы.Value);
                d = Convert.ToInt32(numericUpDownГод.Value);

                string prof = comboBoxПрофессии.SelectedItem.ToString();
                var find = b == 4 ?
                    InitialData.KCN.Find(x => x.name == prof):
                    InitialData.Professions.Find(x => x.name == prof);

                a = find.chiper;

                string group = b == 4 ? $"{a}.{c}.{d}" : $"{a}.{b}.{c}.{d}"; //"5414-008.3.7.2023"
                this.textBoxПроверкаСлушателя.Text = $"Слушатель {textBoxФИО.Text} из организации {textBoxОрганизация.Text} добавляется в группу {group}";
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                textBoxОрганизация.Text = "";
                textBoxОрганизация.ReadOnly = true;
            }
            else
            {
                textBoxОрганизация.ReadOnly = false;
            }
        }

        private void buttonСуществующаяГруппаПроверить_Click(object sender, EventArgs e)
        {
            var group = textBoxСуществующаяГруппа.Text.Trim();
            string a = group.Split('.')[0];
            string b = group.Split('.')[1];
            string c = group.Split('.')[2];
            string d = group.Split('.')[3];

            var table = _dbContext.GetDataSet($"select * from ({_querries["Группы"]}) " +
                                                $"where [Шифр] like N'{a}' " +
                                                $"and  [Код занятия] like N'{b}' " +
                                                $"and [Порядковый номер группы] like N'{c}' " +
                                                $"and [Год] like N'{d}';").Tables[0];
            DataRow row = null;
            if (table.Rows.Count == 0)
            {
                MessageBox.Show("Не найдено группы");
            }
            else
            {
                if (MessageBox.Show("Группа существует\nДобавить в эту группу?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    try
                    {
                        
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }

        }

        #endregion

        private void buttonДобавитьСлушателя_Click(object sender, EventArgs e)
        {
            string a = "0000-000";
            int b = 0, c = 0, d = 0;

            b = comboBoxЗанятия.SelectedIndex + 1;
            c = Convert.ToInt32(numericUpDownПорядкоывйНомерГруппы.Value);
            d = Convert.ToInt32(numericUpDownГод.Value);
            string prof = comboBoxПрофессии.SelectedItem.ToString();
            var find = b == 4 ?
                InitialData.KCN.Find(x => x.name == prof) :
                InitialData.Professions.Find(x => x.name == prof);

            a = find.chiper;
            _dbContext.Execute($"EXEC InsertNew  @Org = N'{textBoxОрганизация.Text}'," +
                           $"@Fio = N'{textBoxФИО.Text}', @DoB = N'{numericUpDownГодРождения.Value}', @Educ = N'{textBoxОбразование.Text}', @Pass = N'{textBoxНомерПаспорта.Text}', " +
                           $"@Chiper = N'{a}', @Cod = {b}, @Lvl = {numericUpDownБудущийРазряд.Value}, @Nom = {numericUpDownПорядкоывйНомерГруппы.Value}, " +
                           $"@Year = N'{numericUpDownГод.Value}', @DTF = N'{dateTimePickerТеорияС.Value:dd.MM.yyyy}', " +
                           $"@DTT = N'{dateTimePickerТеорияПо.Value:dd.MM.yyyy}', @DPF = N'{dateTimePickerПрактикаС.Value:dd.MM.yyyy}', " +
                           $"@DPT = N'{dateTimePickerПрактикаПо.Value:dd.MM.yyyy}', @DCn = N'{dateTimePickerКонсультащия.Value:dd.MM.yyyy}', " +
                           $"@DEx = N'{dateTimePickerЭкзамен.Value:dd.MM.yyyy}'; ");

            MessageBox.Show("Успешно");
        }

        private void groupBoxСуществующаяГруппа_Enter(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            _templateHandler = new TemplateHandler();
            foreach (var template in _templateHandler.Templates)
            {
                comboBoxДокументы.Items.Add(template.Key);
            }
            InitialData.RefreshInitial();
            foreach (var it in InitialData.Professions) comboBoxПрофессии.Items.Add(it.name);
            foreach (var it in InitialData.KCN) comboBoxПрофессии.Items.Add(it.name);
            LoadDates();
        }

        private void buttonОбновитьТаблицу_Click(object sender, EventArgs e)
        {
            if (comboBoxВыбратьТаблицу.SelectedIndex != -1)
                comboBoxВыбратьТаблицу_SelectedIndexChanged(sender, e);
        }

        private void buttonПоиск_Click(object sender, EventArgs e)
        {
            string toFind = textBoxПоискПоЗначениюПоля.Text;
            if (string.Equals(comboBoxВыбратьТаблицу.SelectedItem, "Слушатели") && string.Equals(comboBoxПоискПоПолю.SelectedItem, "Полный номер группы") && !string.IsNullOrEmpty(toFind))
            {
                // 7215-002.1.7.2023
                // КЦН 1.61.4.2023
                var parts = toFind.Split('.');
                string sql;
                if (parts.Length == 4)
                {
                    if (parts[0].Contains("КЦН"))
                    {
                        sql = $"select * from ({_querries["Слушатели"]}) sel " +
                            $"where sel.[Шифр] like N'%{parts[0]}.{parts[1]}%' " +
                            $"and sel.[Порядковый номер группы]={parts[2]} " +
                            $"and sel.[Год] like N'{parts[3]}' ";
                    }
                    else
                    {
                        sql = $"select * from ({_querries["Слушатели"]}) sel " +
                            $"where sel.[Шифр] like N'%{parts[0]}%' " +
                            $"and sel.[Код занятия]={parts[1]} " +
                            $"and sel.[Порядковый номер группы]={parts[2]} " +
                            $"and sel.[Год] like N'{parts[3]}' ";
                    }

                    //MessageBox.Show(sql);
                    SetData(sql);
                }

                return;
            }

            if (string.Equals(comboBoxВыбратьТаблицу.SelectedItem, "Группы") && string.Equals(comboBoxПоискПоПолю.SelectedItem, "Полный номер группы") && !string.IsNullOrEmpty(toFind))
            {
                // 7215-002.1.7.2023
                // КЦН 1.61.4.2023
                var parts = toFind.Split('.');
                string sql;
                if (parts.Length == 4)
                {
                    if (parts[0].Contains("КЦН"))
                    {
                        sql = $"select * from ({_querries["Группы"]}) sel " +
                            $"where sel.[Шифр] like N'%{parts[0]}.{parts[1]}%' " +
                            $"and sel.[Порядковый номер группы]={parts[2]} " +
                            $"and sel.[Год] like N'{parts[3]}'; ";
                    }
                    else
                    {
                        sql = $"select * from ({_querries["Группы"]}) sel " +
                            $"where sel.[Шифр] like N'%{parts[0]}%' " +
                            $"and sel.[Код занятия]={parts[1]} " +
                            $"and sel.[Порядковый номер группы]={parts[2]} " +
                            $"and sel.[Год] like N'{parts[3]}'; ";
                    }

                    //MessageBox.Show(sql);
                    SetData(sql);
                }

                return;
            }

            if (comboBoxПоискПоПолю.SelectedIndex != -1 && !string.IsNullOrEmpty(toFind))
            {
                string sql = $"select * from ({_querries[comboBoxВыбратьТаблицу.SelectedItem.ToString()]}) sel " +
                             $"where sel.[{comboBoxПоискПоПолю.SelectedItem.ToString()}] like N'%{toFind}%' ";
                SetData(sql);
            }
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int idcol = dataGridView1.CurrentCell.ColumnIndex;
            int idrow = dataGridView1.CurrentCell.RowIndex;
            string value = dataGridView1.CurrentCell.Value.ToString();

            if (MessageBox.Show($"Вы действительно хотите изменить значение на {value}", 
                                "Изменение", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                //string sql = $"update {comboBoxВыбратьТаблицу.SelectedItem} set " +
                //    $"{}={value}"
            }
            else
            {
                dataGridView1.CancelEdit();
            }
        }

        private void buttonУдалитьВыделенное_Click(object sender, EventArgs e)
        {
            var rows = dataGridView1.SelectedRows;
            List<string> sql = new List<string>();

            for (int i = 0; i < rows.Count; i++)
            {
                switch (comboBoxВыбратьТаблицу.SelectedItem.ToString())
                {
                    case "Профессии": 
                        { 
                            MessageBox.Show("Профессии и КЦН нельзя удалять"); 
                            break; 
                        }
                    case "Организации":
                        {
                            int j = FindHeaderId(dataGridView1.Columns, "Название");
                            if (j == -1) break;
                            sql.Add($"delete from [Организация] where [Название] like N'{rows[i].Cells[j].Value}'; ");
                            break;
                        }
                    case "Слушатели":
                        {
                            int j = FindHeaderId(dataGridView1.Columns ,"ФИО");
                            if (j == -1) break;
                            sql.Add($"delete from [Слушатель] where [ФИО] like N'{rows[i].Cells[j].Value}'; ");
                            break;
                        }
                    case "Группы":
                        {
                            int j1 = FindHeaderId(dataGridView1.Columns, "Дата теории с");
                            if (j1 == -1) break;
                            int j2 = FindHeaderId(dataGridView1.Columns, "Дата теории по");
                            if (j2 == -1) break;
                            int j3 = FindHeaderId(dataGridView1.Columns, "Порядковый номер группы");
                            if (j3 == -1) break;
                            sql.Add($"delete from [Группа] where [Дата теории с] like N'%{rows[i].Cells[j1].Value}%' " +
                                                           $"and [Дата теории по] like N'%{rows[i].Cells[j2].Value}%' " +
                                                           $"and [Порядковый номер группы]={rows[i].Cells[j3].Value}; ");
                            break;
                        }
                    default:
                        {
                            break;
                        }
                }

                if (sql.Count != 0)
                {
                    //_dbContext.Execute(sql.ToArray())
                    MessageBox.Show(sql[0]);
                }
            }
        }

        int FindHeaderId(DataGridViewColumnCollection cols, string header)
        {
            for(int i = 0; i < cols.Count; i++)
            {
                if(cols[i].HeaderText == header)
                {
                    return i;
                }
            }

            return -1;
        }

        private void buttonСоставитьДокументы_Click(object sender, EventArgs e)
        {
            if (comboBoxВыбратьТаблицу.SelectedIndex == -1)
            {
                return;
            }

            if (comboBoxВыбратьТаблицу.SelectedItem.ToString() != "Группы")
            {
                MessageBox.Show("Составить документы можно только по группе");
                return;
            }

            var rows = dataGridView1.SelectedRows;
            if (rows.Count == 0)
            {
                MessageBox.Show("Выберите строки (для выделения всей строки нажмите в самую правую ячейку)");
                return;
            }

            for (int i = 0; i < rows.Count; i++)
            {
                _groups.Add(new Group(rows[i]));
            }
            var strings = new string[rows.Count];
            for (int i = 0; i < rows.Count; i++)
            {
                strings[i] = _groups[i].Группа;
            }
            richTextBoxВыбранныеГруппы.Lines = strings;
            tabControl1.SelectedTab = tabPageDocs;
        }

        private void DocumentAutoFill()
        {
            if (dataGridViewДокументы.RowCount <= 0) return;
            if (_groups.Count <= 0) return;

            for (int i = 0; i < dataGridViewДокументы.RowCount; i++)
            {
                if (string.IsNullOrEmpty(dataGridViewДокументы.Rows[i].Cells["ColumnValue"].Value.ToString()))
                {
                    dataGridViewДокументы.Rows[i].Cells["ColumnValue"].Value = dataGridViewДокументы.Rows[i].Cells["ColumnTag"].Value.ToString() switch
                    {
                        "<Наз.Проф>" => _groups[0].Название,
                        "<Наз.Курса>" => _groups[0].Название,
                        "<Обуч.С>" => _groups[0].ДатаОбучС,
                        "<Обуч.По>" => _groups[0].ДатаОбучПо,
                        "<Практ.С>" => _groups[0].ДатаПрактС,
                        "<Практ.По>" => _groups[0].ДатаПрактПо,
                        "<Конс>" => _groups[0].ДатаКонсультации,
                        "<Экз>" => _groups[0].ДатаЭкзамена,
                        "<Группа>" => _groups[0].Группа,
                        "<Орг>" => _groups[0].Организация,
                        "<Разряд>" => _groups[0].Разряд,
                        "<Колич.Чел>" => "одного человека",
                        _ => "",
                    };

                    if (dataGridViewДокументы.Rows[i].Cells["ColumnTag"].Value.ToString() == "<Разряд.Букв>")
                    {
                        dataGridViewДокументы.Rows[i].Cells["ColumnValue"].Value = _groups[0].Разряд switch
                        {
                            1 => "первый",
                            2 => "второй",
                            3 => "третий",
                            4 => "четверный",
                            5 => "пятый",
                            6 => "шестой",
                            _ => "",
                        };
                    }
                    if (dataGridViewДокументы.Rows[i].Cells["ColumnTag"].Value.ToString() == "<Подг.Пере.Пов>")
                    {
                        dataGridViewДокументы.Rows[i].Cells["ColumnValue"].Value = _groups[0].типОбучения switch
                        {
                            ТипОбучения.подг => "профессиональной подготовки",
                            ТипОбучения.переподг => "профессиональной переподготовки",
                            ТипОбучения.пов => "повышения квалификации",
                            _ => "",
                        };
                    }
                }
            }

        }

        private void buttonОчиститьСлушателей_Click(object sender, EventArgs e)
        {
            richTextBoxВыбранныеГруппы.Clear();
            _groups.Clear();
        }

        private void buttonДобавитьДокумент_Click(object sender, EventArgs e)
        {
            if (comboBoxДокументы.SelectedIndex == -1)
            {
                MessageBox.Show("Выберите документ");
                return;
            }

            if (_groups.Count == 0)
            {
                MessageBox.Show("Выберите группы");
                return;
            }

            Group group = _groups[0];
            _directoryMatcher = new DirectoryMatcher(
                group.типПрофессии, group.Название,
                group.ПорядковыйНомерГруппы, 
                DateTime.ParseExact(group.ДатаОбучС, "dd.MM.yyyy", _currentCulture), 
                group.типОбучения, 
                group.Разряд, 
                group.Организация);
            textBoxПутьДокументов.Text = _directoryMatcher.Workdir.FullName;
            if (_templateHandler.Templates[comboBoxДокументы.SelectedItem.ToString()].IsSelectetByUser)
            {
                MessageBox.Show("Документ уже в таблице");
                return;
            }

            _templateHandler.Templates[comboBoxДокументы.SelectedItem.ToString()].IsSelectetByUser = true;
            Template template = _templateHandler.Templates[comboBoxДокументы.SelectedItem.ToString()];

            //bool alreadyInTable = false;
            //for (int i = 0; i < dataGridViewДокументы.Rows.Count && !alreadyInTable; i++)
            //{
            //    alreadyInTable = dataGridViewДокументы.Rows[i].Cells[0].Value.ToString() == template.File.Name;
            //}

            //if (alreadyInTable)
            //{
            //    MessageBox.Show("Документ уже в таблице");
            //    return;
            //}

            foreach (var row in template.Table)
            {
                dataGridViewДокументы.Rows.Add(template.File.Name, row.Tag, row.Descriptoin, row.Value);
            }

            var distinctIndexes = new List<int>();
            
            for (int i = 0; i < dataGridViewДокументы.Rows.Count; i++)
            {
                bool unique = true;
                for (int j = 0; j < i; j++)
                {
                    unique &= dataGridViewДокументы.Rows[i].Cells[1].Value.ToString() != dataGridViewДокументы.Rows[j].Cells[1].Value.ToString();
                }

                if (unique) 
                { 
                    distinctIndexes.Add(i); 
                }
            }

            for (int i = dataGridViewДокументы.Rows.Count - 1; i >= 0; i--)
            { 
                if (!distinctIndexes.Contains(i))
                {
                    dataGridViewДокументы.Rows.RemoveAt(i);
                }
            }

            DocumentAutoFill();

            //работает
            //Color color = 
            //    dataGridViewДокументы.Rows.Count == 0 || 
            //    dataGridViewДокументы.Rows[^1].DefaultCellStyle.BackColor == Color.LightGray ?
            //    Color.White : Color.LightGray;
            //foreach (var row in template.Table)
            //{
            //    dataGridViewДокументы.Rows.Add(template.File.Name, row.Tag, row.Descriptoin, row.Value);
            //    dataGridViewДокументы.Rows[^1].DefaultCellStyle.BackColor = color;
            //}
        }

        private void buttonОчиститьДокументы_Click(object sender, EventArgs e)
        {
            dataGridViewДокументы.Rows.Clear();
            foreach (var template in _templateHandler.Templates)
                template.Value.IsSelectetByUser = false;
        }

        private void buttonСгенерироватьДокументы_Click(object sender, EventArgs e)
        {
            if (dataGridViewДокументы.Rows.Count == 0) return;
            for (int i = 0; i < dataGridViewДокументы.Rows.Count; i++)
            {
                if (string.IsNullOrWhiteSpace(dataGridViewДокументы.Rows[i].Cells[3].Value.ToString()))
                {
                    DialogResult dialogResult = MessageBox.Show(
                        "Не все данные в таблице заполнены. Данные в шаблонах могут быть заменены на пустые строки.\n " +
                        "Продолжить?", "Предупреждение", MessageBoxButtons.YesNo);
                    if (dialogResult != DialogResult.Yes) return;
                    break;
                }
            }

            List<Template> selectedTemplates = (from template in _templateHandler.Templates
                                                 where template.Value.IsSelectetByUser
                                                 select template.Value).ToList();
            
            Dictionary<string, string> TagValuePairs = new();
            for (int i = 0; i < dataGridViewДокументы.Rows.Count; i++)
            {
                string tag = dataGridViewДокументы.Rows[i].Cells[1].Value.ToString();
                string value = dataGridViewДокументы.Rows[i].Cells[3].Value.ToString();
                TagValuePairs.Add(tag, value);
            }
            
            foreach (var template in selectedTemplates)
            {
                WordCreator word = new(template.File,
                    new System.IO.FileInfo(System.IO.Path.Combine(_directoryMatcher.Workdir.FullName, template.File.Name)));
                word.Process(TagValuePairs);
            }

            //Dictionary<string, string> TagValuePairs = new();

            //for (int i = 0; i < dataGridViewДокументы.Rows.Count; i++)
            //{
            //    string templateName = dataGridViewДокументы.Rows[i].Cells[0].Value.ToString();
            //    string tag = dataGridViewДокументы.Rows[i].Cells[1].Value.ToString();
            //    string value = dataGridViewДокументы.Rows[i].Cells[3].Value.ToString();
            //    TagValuePairs.Add(tag, value);
            //    i++;
            //    while (i < dataGridViewДокументы.Rows.Count && dataGridViewДокументы.Rows[i].Cells[0].Value.ToString() == templateName)
            //    {
            //        tag = dataGridViewДокументы.Rows[i].Cells[1].Value.ToString();
            //        value = dataGridViewДокументы.Rows[i].Cells[3].Value.ToString();
            //        TagValuePairs.Add(tag, value);
            //        i++;
            //    }

            //    WordCreator word = new(_templateHandler.Templates[templateName].File,
            //        new System.IO.FileInfo(System.IO.Path.Combine(_directoryMatcher.Workdir.FullName, templateName)));
            //    word.Process(TagValuePairs);

            //    TagValuePairs.Clear();
            //}
            MessageBox.Show("Успешно сгенерировано");
        }

        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            var selected = monthCalendar1.SelectionStart;
            labelВыбраннаяДата.Text = selected.ToString("dd.MM.yyyy");
            var all = _monthEvents.FindAll(x => x.Date == selected.Date);
            richTextBoxСобытияВыбраннойДаты.Clear();
            foreach (var date in all)
            {
                richTextBoxСобытияВыбраннойДаты.Text += $"{date.Date} - {date.Event};\n";
            }
        }

        private void LoadDates()
        {
            var table = _dbContext.GetDataSet(_querries["Группы"]).Tables[0];
            foreach (DataRow row in table?.Rows)
            {
                try
                {
                    DateTime t_from = DateTime.Parse(row?["Дата теории с"].ToString(), _currentCulture),
                            //t_to = DateTime.Parse(row?["Дата теории по"].ToString(), _currentCulture),
                            p_from = DateTime.Parse(row?["Дата практики с"].ToString(), _currentCulture),
                            //p_to = DateTime.Parse(row?["Дата практики по"].ToString(), _currentCulture),
                            cons = DateTime.Parse(row?["Дата консультации"].ToString(), _currentCulture),
                            exs = DateTime.Parse(row?["Дата экзамена"].ToString(), _currentCulture);
                    string group = $"{row?["Шифр"].ToString()}.{row?["Код занятия"].ToString()}.{row?["Порядковый Номер Группы"].ToString()}.{row?["Год"].ToString()}";

                    if (exs.Date >= DateTime.Today) _monthEvents.Add((exs, $"Экзамен для группы {group}"));
                    if (cons.Date >= DateTime.Today) _monthEvents.Add((cons, $"Консультация для группы {group}"));
                    //if (p_to.Date >= DateTime.Today) _monthEvents.Add((p_to, $"Конец практики для группы {group}"));
                    if (p_from.Date >= DateTime.Today) _monthEvents.Add((p_from, $"Начало практики для группы {group}"));
                    //if (t_to.Date >= DateTime.Today) _monthEvents.Add((t_to, $"Конец теор. обуч. для группы {group}"));
                    if (t_from.Date >= DateTime.Today) _monthEvents.Add((t_from, $"Начало теор. обуч. для группы {group}")); 
                }
                finally { }
            }

            _monthEvents.Sort( (x , y) =>
            {
                if (x.Date > y.Date) return 1;
                else if (x.Date < y.Date) return -1;
                else return 0;
            });

            List<DateTime> dates = new();
            foreach(var date in _monthEvents)
            {
                dates.Add(date.Date);
            }
            monthCalendar1.BoldedDates = dates.ToArray();
            if (_monthEvents.Count == 0) return;

            var first = _monthEvents.First();
            var all = _monthEvents.FindAll(x => x.Date == first.Date);
            labelБлижайшееСобытие.Text = first.Date.ToString("dd.MM.yyyy");
            foreach (var date in all)
            {
                richTextBoxБлижайшееСобытие.Text += $"{date.Date} - {date.Event};\n";
            }
            
            labelВыбраннаяДата.Text = DateTime.Today.ToString("dd.MM.yyyy");
            all = _monthEvents.FindAll(x => x.Date == DateTime.Today.Date);
            foreach (var date in all)
            {
                richTextBoxСобытияВыбраннойДаты.Text += $"{date.Date} - {date.Event};\n";
            }
        }

        private void buttonОбновитьКалендарь_Click(object sender, EventArgs e)
        {
            _monthEvents.Clear();
            richTextBoxБлижайшееСобытие.Clear();
            richTextBoxСобытияВыбраннойДаты.Clear();
            LoadDates();
        }
    }
}
