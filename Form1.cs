using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;
using System.Configuration;

namespace AtualizaBaseEtiquetas
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            AtualizaBase();
        }

        void AtualizaBase()
        {
            try
            {

                System.IO.File.Delete(@"C:\Opticon\ESL Server\Output\dbase.csv");
                var oracleHost = ConfigurationSettings.AppSettings["OracleHost"];
                var oraclePort = ConfigurationSettings.AppSettings["OraclePort"];
                var bDInstance = ConfigurationSettings.AppSettings["BDInstance"];
                var bDUser = ConfigurationSettings.AppSettings["BDUser"];
                var bDPassword = ConfigurationSettings.AppSettings["BDPassword"];

                DataTable Tbl = new DataTable();
                Tbl.Columns.Add("FIELD_0", typeof(string));
                Tbl.Columns.Add("FIELD_1", typeof(string));
                Tbl.Columns.Add("FIELD_2", typeof(string));
                Tbl.Columns.Add("FIELD_3", typeof(string));
                Tbl.Columns.Add("FIELD_4", typeof(string));
                Tbl.Columns.Add("FIELD_5", typeof(string));
                Tbl.Columns.Add("FIELD_6", typeof(string));
                DataRow Linha;

                OracleConnection conn2 = new OracleConnection("Data Source=(DESCRIPTION="
                                                            + "(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=" + oracleHost + ")(PORT=" + oraclePort + ")))"
                                                            + "(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=" + bDInstance + ")));"
                                                            + "User Id=" + bDUser + ";Password=" + bDPassword + "");
                System.IO.StreamWriter log = new System.IO.StreamWriter(@"C:\Opticon\ESL Server\AtualizaBase\bin\log\log" + DateTime.Now.ToString("dd-MM-yyyy") + ".txt", true);
                log.WriteLine(DateTime.Now.ToString() + " >>> Sessão aberta no BD");
                log.Close();
                OracleCommand oCmd2 = new OracleCommand();
                string query2 = "select 'I' as field_0, " +
           "M.TMER_CODIGO_BARRAS_UKN as field_1, " +
           "b.tmer_codigo_barras_alter_pk as field_2, " +
           "m.tmer_nome as field_3, " +
           "M.TMER_UNIDADE_FISICA_FK as field_4, " +
           "e.tmer_preco_venda as field_5, " +
           "p.tfid_preco_venda as field_6 " +
      "from tmer_mercadoria m " +
     "inner " +
      "join tmer_codigo_barras b " +
    "on b.tmer_codigo_barras_ean_fkn = m.tmer_codigo_barras_ukn " +
    "left " +
      "join tfid_promocao p " +
    "on p.tfid_codigo_pri_fk_pk = m.tmer_codigo_pri_pk " +
    "and p.tfid_codigo_sec_fk_pk = m.tmer_codigo_sec_pk " +
       "and p.tfid_tipo_preco = 1 " +
       "and p.tfid_data_fim >= trunc(sysdate) " +
       "and p.tfid_unidade_fk_pk = 04 " +
     "inner join tmer_estoque e " +
        "on e.tmer_codigo_pri_fk_pk = m.tmer_codigo_pri_pk " +
       "and e.tmer_codigo_sec_fk_pk = m.tmer_codigo_sec_pk " +
       "and e.tmer_unidade_fk_pk = 04 " +
     "inner join tloja_sap s " +
        "on e.tmer_unidade_fk_pk = s.loja_proton_uk " +
     "where s.loja_sap_pk = 1006 " +
       "and M.TMER_CODIGO_BARRAS_UKN LIKE '5%' " +
     "ORDER BY 4, 3";
                oCmd2.CommandText = query2;
                oCmd2.CommandType = CommandType.Text;
                oCmd2.Connection = conn2;
                conn2.Open();
                OracleDataReader ler1 = oCmd2.ExecuteReader();
                while (ler1.Read())
                {
                    Linha = Tbl.NewRow();
                    Linha["FIELD_0"] = ler1.GetValue(0).ToString();
                    Linha["FIELD_1"] = ler1.GetValue(1).ToString();
                    Linha["FIELD_2"] = ler1.GetValue(2).ToString();
                    Linha["FIELD_3"] = ler1.GetValue(3).ToString();
                    Linha["FIELD_4"] = ler1.GetValue(4).ToString();
                    Linha["FIELD_5"] = ler1.GetValue(5).ToString();
                    Linha["FIELD_6"] = ler1.GetValue(6).ToString();
                    Tbl.Rows.Add(Linha);
                }
                conn2.Close();
                dataGridView1.DataSource = Tbl;
                System.IO.StreamWriter log1 = new System.IO.StreamWriter(@"C:\Opticon\ESL Server\AtualizaBase\bin\log\log" + DateTime.Now.ToString("dd-MM-yyyy") + ".txt", true);
                log1.WriteLine(DateTime.Now.ToString() + " >>> Dados recuperados.");
                log1.Close();
                StringBuilder sb = new StringBuilder();

                IEnumerable<string> columnNames = Tbl.Columns.Cast<DataColumn>().
                                                  Select(column => column.ColumnName);
                //sb.AppendLine(string.Join(";", columnNames)); //Cabeçalho do CSV

                foreach (DataRow row in Tbl.Rows)
                {
                    IEnumerable<string> fields = row.ItemArray.Select(field => field.ToString());
                    sb.AppendLine(string.Join(";", fields));
                }
                System.IO.StreamWriter log2 = new System.IO.StreamWriter(@"C:\Opticon\ESL Server\AtualizaBase\bin\log\log" + DateTime.Now.ToString("dd-MM-yyyy") + ".txt", true);
                log2.WriteLine(DateTime.Now.ToString() + " >>> Criando arquivo csv.");
                log2.Close();
                System.IO.File.WriteAllText(@"C:\Opticon\ESL Server\Input\dbase.csv", sb.ToString());
                System.IO.File.WriteAllText(@"C:\Opticon\ESL Server\AtualizaBase\bin\files\dbase.csv", sb.ToString());
                System.IO.StreamWriter log3 = new System.IO.StreamWriter(@"C:\Opticon\ESL Server\AtualizaBase\bin\log\log" + DateTime.Now.ToString("dd-MM-yyyy") + ".txt", true);
                log3.WriteLine(DateTime.Now.ToString() + " >>> Arquivo csv criado e depositado na pasta Input do ESL Server.");
                log3.WriteLine(DateTime.Now.ToString() + " >>> Processo finalizado.");
                log3.Close();
                Application.Exit();
            }
            catch (Exception ex)
            {
                System.IO.StreamWriter log2 = new System.IO.StreamWriter(@"C:\Opticon\ESL Server\AtualizaBase\bin\log\log" + DateTime.Now.ToString("dd-MM-yyyy") + ".txt", true);
                log2.WriteLine(DateTime.Now.ToString() + " Ocorreu uma exceção.");
                log2.WriteLine(DateTime.Now.ToString() + " " + ex.Message);
                log2.Close();
                Application.Exit();
            }
    }
    }
}
