using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
/*
0  TABLE_CATALOG
1  TABLE_SCHEMA
2  TABLE_NAME
3  COLUMN_NAME
4  COLUMN_GUID
5  COLUMN_PROPID
6  ORDINAL_POSITION
7  COLUMN_HASDEFAULT
8  COLUMN_DEFAULT
9  COLUMN_FLAGS
10 IS_NULLABLE
11 DATA_TYPE
12 TYPE_GUID
13 CHARACTER_MAXIMUM_LENGTH
14 CHARACTER_OCTET_LENGTH
15 NUMERIC_PRECISION
16 NUMERIC_SCALE
17 DATETIME_PRECISION
18 CHARACTER_SET_CATALOG
19 CHARACTER_SET_SCHEMA
20 CHARACTER_SET_NAME
21 COLLATION_CATALOG
22 COLLATION_SCHEMA
23 COLLATION_NAME
24 DOMAIN_CATALOG
25 DOMAIN_SCHEMA
26 DOMAIN_NAME
27 DESCRIPTION
 */
namespace DataBaseAssistant
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            //pars(@"E:\Users\abataloff\Dropbox\TaxiData\Тестирование.mdb");
        }

        private void pars(string a_fileName)
        {
            var _connection = new OleDbConnection(string.Format("Provider=Microsoft.Jet.OLEDB.4.0;" +
                                                                "Data Source={0};" +
                                                                "Persist Security Info=True;",
                                                                a_fileName));
            var _result = new StringBuilder();
            _connection.Open();
            var _tables = _connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] {null, null, null, "TABLE"});
            foreach (DataRow _itemTable in _tables.Rows)
            {
                var _tableName = (string)_itemTable["TABLE_NAME"];
                var _columns = _connection.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object[] { null, null, _tableName, null });
                foreach (DataRow _itemColumn in _columns.Rows)
                {
                    var _columnName = _itemColumn["COLUMN_NAME"];
                    var _type =(OleDbType) Convert.ToInt32( _itemColumn["DATA_TYPE"]);
                    _result.AppendLine(string.Format("{0}:{1} {2}",_tableName,_columnName,_type));
                }
            }
            _connection.Close();
            richTextBox1.Text = _result.ToString();
            var _outFile = Path.ChangeExtension(a_fileName, "txt");
            TextWriter _tw = new StreamWriter(_outFile);
            _tw.Write(_result.ToString());
            _tw.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var _ofd = new OpenFileDialog();
            if (_ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                pars(_ofd.FileName);
            }
        }
    }
}
