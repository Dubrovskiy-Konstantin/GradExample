using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Промышленная_Безопасность
{
    public partial class Spec : Form
    {
        DBContext _dbContext;
        public Spec()
        {
            _dbContext = new DBContext();
            InitializeComponent();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            var sql = "Select * from [dbo].[Аккаунты]";
            var _currentTable = _dbContext.GetDataSet(sql).Tables[0];
            dataGridView1.DataSource = _currentTable;
        }
    }
}
