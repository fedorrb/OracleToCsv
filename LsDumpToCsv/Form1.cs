using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Oracle.DataAccess.Client;

namespace LsDumpToCsv
{
    public partial class Form1 : Form
    {
        OracleConnection connection = new OracleConnection(@"Data Source=SPP181;User Id=IKIS;Password=ikis;");
        //OracleConnection connection = new OracleConnection();
        OracleCommand command = new OracleCommand();
        String selectSQL = "SELECT LS_NLS FROM LS_LS_DATA_UPLOAD WHERE LS_RAJ = 3001 ORDER BY LS_NLS";
        private Dictionary<string, string> dict = new Dictionary<string, string>();
        private string Raj = string.Empty;
        private bool nsi = false;
        private stPath stPathIni;
        private string filepath = string.Empty;
        private List<string> listRaj = new List<string>();

        public Form1()
        {
            //not using
        }

        public Form1(string args0)
        {
            InitializeComponent();
            stPathIni = new stPath();
            stPathIni.LoadIniFile();
            //if (stPathIni.connectionString != "")
                //connection.ConnectionString = stPathIni.connectionString;
            if (string.Compare(args0, "NSI", true) == 0)
                this.nsi = true;
            this.Raj = "1";
            label1.Text = "";
            label2.Text = "";
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            progressBar1.Value = 0;
            stPathIni.LoadIniFile();
            command.Connection = connection;
            GetListRaj(ref listRaj);
            progressBar1.Maximum = 30 * listRaj.Count + 31;
            filepath = Path.Combine(stPathIni.lsPath, "KL");

            if (Directory.Exists(filepath) == false)
            {
                try
                {
                    Directory.CreateDirectory(filepath);
                }
                catch (Exception exception)
                {
                    MessageBox.Show(exception.Message);
                    Application.Exit();
                }
            }
            else
            {
                try
                {
                    //Delete all files from dir
                    System.IO.Directory.Delete(filepath, true);
                    Directory.CreateDirectory(filepath);
                }
                catch (Exception exception)
                {
                    MessageBox.Show(exception.Message);
                    Application.Exit();
                }
            }

            DovidnikiToCSV();
            if (this.nsi == false)
            {
                foreach (string r in listRaj)
                {
                    Raj = r;
                    if (Raj.Length > 0)
                    {
                        label1.Text = "Район: - " + Raj;
                        //if (stPathIni.connectionString != "")
                        //connection.ConnectionString = stPathIni.connectionString;
                        filepath = Path.Combine(stPathIni.lsPath, Raj);
                        if (Directory.Exists(filepath) == false)
                        {
                            try
                            {
                                Directory.CreateDirectory(filepath);
                            }
                            catch (Exception exception)
                            {
                                MessageBox.Show(exception.Message);
                                Application.Exit();
                            }
                        }
                        else
                        {
                            try
                            {
                                //Delete all files from dir
                                System.IO.Directory.Delete(filepath, true);
                                Directory.CreateDirectory(filepath);
                            }
                            catch (Exception exception)
                            {
                                MessageBox.Show(exception.Message);
                                Application.Exit();
                            }
                        }
                        RegionToCSV();
                    }
                }
            }
            //stPathIni.SaveIniFile();
            Application.Exit();
        }

        protected void RegionToCSV()
        {
            {
                //формування CSV файлів по всіх базах для особового рахунку lsNls
                Write_B_LS("LS_SHIFR_DATA_UPLOAD", "shifr.txt", "ls_raj, ls_nls, shifr_datec");
                Write_B_LS("LS_CHE_DATA_UPLOAD", "che.txt", "ls_raj, ls_nls, che_nomig");
                Write_B_LS("LS_CHEZP_DATA_UPLOAD", "chezp.txt", "ls_raj, ls_nls, chezp_nomig");
                Write_B_LS("LS_KOR_DATA_UPLOAD", "kor.txt", "ls_raj, ls_nls, kor_datsm, kor_nomig");
                Write_B_LS("LS_IGD_DATA_UPLOAD", "igd.txt", "ls_raj, ls_nls, igd_dusn, igd_nomig");
                Write_B_LS("LS_NP_DATA_UPLOAD", "np.txt", "ls_raj, ls_nls, np_kfn, np_dnprav");
                Write_B_LS("LS_PER_DATA_UPLOAD", "per.txt", "ls_raj, ls_nls, per_kfn, per_dnpen");
                Write_B_LS("LS_INV_DATA_UPLOAD", "inv.txt", "ls_raj, ls_nls, inv_nomig, inv_dnpi");
                Write_B_LS("LS_NAZN_DATA_UPLOAD", "nazn.txt", "ls_raj, ls_nls, nazn_op, nazn_dnaz");
                Write_B_LS("LS_NAC_DATA_UPLOAD", "nac.txt", "ls_raj, ls_nls, nac_god, nac_mec, nac_npp, nac_op, nac_kfn");
                Write_B_LS("LS_OSOB_DATA_UPLOAD", "osob.txt", "ls_raj, ls_nls, osob_code");
                Write_B_LS("LS_STG_DATA_UPLOAD", "stg.txt", "ls_raj, ls_nls, stg_nomig, stg_kods");
                Write_B_LS("LS_STGP_DATA_UPLOAD", "stgp.txt", "ls_raj, ls_nls, stgp_nomig, stgp_dbeg");
                Write_B_LS("LS_SV1_DATA_UPLOAD", "sv1.txt", "ls_raj, ls_nls");
                Write_B_LS("LS_SV2_DATA_UPLOAD", "sv2.txt", "ls_raj, ls_nls");
                Write_B_LS("LS_ATST_DATA_UPLOAD", "atst.txt", "ls_raj, ls_nls, atst_type, atst_num");
                Write_B_LS("LS_UMER_DATA_UPLOAD", "umer.txt", "ls_raj, ls_nls, umer_nomig");
                Write_B_LS("LS_ZP_DATA_UPLOAD", "zp.txt", "ls_raj, ls_nls, zp_dateb");
                Write_B_LS("LS_ZRB_DATA_UPLOAD", "zrb.txt", "ls_raj, ls_nls, zrb_nr, zrb_nrs, zrb_dbeg");
                Write_B_LS("LS_ZRBS_DATA_UPLOAD", "zrbs.txt", "ls_raj, ls_nls, zrbs_nr, zrbs_nrs");
                Write_B_LS("LS_ISPL_DATA_UPLOAD", "ispl.txt", "ls_raj, ls_nls, ispl_kud, ispl_num");
                Write_B_LS("LS_UD_DATA_UPLOAD", "ud.txt", "ls_raj, ls_nls, ispl_kud, ispl_num");
                Write_B_LS("LS_UD2_DATA_UPLOAD", "ud2.txt", "ls_raj, ls_nls, ispl_kud, ispl_num");
                Write_B_LS("LS_UD3_DATA_UPLOAD", "ud3.txt", "ls_raj, ls_nls, ispl_kud, ispl_num");
                Write_B_LS("LS_DET_DATA_UPLOAD", "det.txt", "ls_raj, ls_nls, ispl_kud, ispl_num, det_datar");
                Write_B_LS("LS_NUDR_DATA_UPLOAD", "nudr.txt", "ls_raj, ls_nls, ispl_kud, ispl_num, nudr_god, nudr_mec");
                Write_B_LS("LS_PPL_DATA_UPLOAD", "ppl.txt", "ls_raj, ls_nls, ppl_prizn, ppl_dateps");
                Write_B_LS("B_TOTAL_UPLOAD", "total.txt");
                Write_B_LS_LS("LS_LS_DATA_UPLOAD", "ls.txt", "ls_nls");
                Write_B_KLUL("KLUL", "klul.txt", "CODE");
            }
        }

        protected void DovidnikiToCSV()
        {
            if(this.nsi) {
                Write_NSI("nsi_building_upload", "nsibuilding.txt");
                Write_NSI("nsi_delivery_hist_upload", "nsideliveryhist.txt");
                Write_NSI("nsi_dlh2bld_upload", "dlh2bld.txt");
                Write_NSI("nsi_enterprise_upload", "nsienterprise.txt");
            }
            Write_NSI("NSI_AREA_UPLOAD", "nsiarea.txt", "ar_id");
            Write_NSI("NSI_INDEX_UPLOAD", "nsiindex.txt", "ind_ind");
            Write_NSI("NSI_UL_UPLOAD", "nsiul.txt", "ul_id");
            Write_NSI("NSI_UL_KAT_UPLOAD", "nsiulkat.txt", "kul_code");
            Write_NSI("nsi_chernobyl_cat_upload", "nsichernkat.txt");
            Write_NSI("nsi_chernobyl_resident_upload ", "nsichernres.txt", "chr_id");
            Write_NSI("nsi_com_center_upload", "nsicomcentr.txt");
            Write_NSI("nsi_delivery_upload", "nsidelivery.txt");
            Write_NSI("nsi_district_upload", "nsidistrict.txt");
            Write_NSI("nsi_dlh2bld_mask_upload", "dlh2bldmask.txt");
            Write_NSI("nsi_earning_tp_upload", "nsiearningtp.txt", "zv_id");
            Write_NSI("nsi_familiarity_upload", "nsifamiliarity.txt", "fml_id");
            Write_NSI("NSI_INDIVIDUALITY_UPLOAD", "nsiINDIVIDUALITY.txt", "nind_code");
            Write_NSI("NSI_IND2STLM_UPLOAD", "nsiIND2STLM.txt");
            Write_NSI("NSI_INVALID_REASON_UPLOAD", "nsiINVALIDREASON.txt", "invr_id");
            Write_NSI("NSI_OP_UPLOAD", "nsiOP.txt", "op_code");
            Write_NSI("NSI_PAYMENT_TP_UPLOAD", "nsiPAYMENTTP.txt");
            Write_NSI("NSI_PENSION_TP_UPLOAD", "nsiPENSIONTP.txt");
            Write_NSI("NSI_PRICH_SN_UPLOAD", "nsiPRICHSN.txt", "psn_code");
            Write_NSI("NSI_PROF_TP_UPLOAD", "nsiPROFTP.txt");
            Write_NSI("NSI_PSB_UPLOAD", "nsiPSB.txt");
            Write_NSI("NSI_REASONDEATH_UPLOAD", "nsiREASONDEATH.txt", "rdt_code");
            Write_NSI("NSI_REASON_NOTPAY_UPLOAD", "nsiREASONNOTPAY.txt", "rnp_id");
            Write_NSI("NSI_RETENTION_INSTIT_UPLOAD", "nsiRETENTIONINSTIT.txt");
            Write_NSI("NSI_RETENTION_UPLOAD", "nsiRETENTION.txt");
            Write_NSI("NSI_SEN_LOCATION_UPLOAD", "nsiSENLOCATION.txt", "sl_id");
            Write_NSI("NSI_SETTLEMENT_UPLOAD", "nsiSETTLEMENT.txt", "stlm_id");
            //Write_NSI("", ".txt", "");
        }

        private void GetListRaj(ref List<string> listRn)
        {
            //DataTable dt = new System.Data.DataTable();// null;
            OracleDataReader d_reader = null;
            CommandType original_com_type = command.CommandType;
            command.CommandType = CommandType.Text;
            command.Parameters.Clear();
            command.CommandText = "select distinct(ls_raj) from LS_LS_DATA_UPLOAD order by ls_raj";
            string sss = string.Empty;
            StringBuilder sb = new StringBuilder();
            int allRecords = 0;
            try
            {
                this.OpenConnection();
                d_reader = command.ExecuteReader();
                while (d_reader.Read())
                {
                    sb.Clear();
                    for (int i = 0; i < d_reader.FieldCount; i++)
                    {
                        sb.Append(d_reader[i].ToString());
                    }
                    listRn.Add(sb.ToString());
                    Application.DoEvents();
                    allRecords++;
                }
                d_reader.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                //throw ex;
                //throw new DataAccessException(ex, string.Format("{0}.ExecSqlQuery()", this.GetType().FullName), sql_query_text, null);
            }
            finally
            {
                command.CommandType = original_com_type;
                if (d_reader != null && !d_reader.IsClosed) d_reader.Dispose();
                //if (!this.holdOpenConnection)
                this.CloseConnection();
            }
        }
        
        private void Write_NSI(string dbName, string fileNameToSave, string orderby = "")
        {
            label2.Text = "Таблиця: - " + dbName;
            this.Text = label1.Text + " ; " + label2.Text;
            string fullFileNameToSave = string.Empty;
            fullFileNameToSave = filepath + "\\" + fileNameToSave;
            if (orderby.Length > 0)
                selectSQL = string.Format("SELECT * FROM {0} ORDER BY {1}", dbName, orderby);
            else
                selectSQL = string.Format("SELECT * FROM {0} ", dbName);
            DataTable dtLs = ExecSqlQuery(selectSQL, fullFileNameToSave);
        }        
        
        private void Write_B_LS(string dbName, string fileNameToSave, string orderby = "")
        {
            label2.Text = "Таблиця: - " + dbName;
            this.Text = label1.Text + " ; " + label2.Text;
            if(progressBar1.Value < progressBar1.Maximum)
                progressBar1.Value += 1;
            string fullFileNameToSave = string.Empty;
            fullFileNameToSave = filepath + "\\" + fileNameToSave;
            if(orderby.Length > 0)
                selectSQL = string.Format("SELECT * FROM {0} WHERE LS_RAJ = {1} ORDER BY {2}", dbName, Raj, orderby);
            else
                selectSQL = string.Format("SELECT * FROM {0} WHERE LS_RAJ = {1}", dbName, Raj);            
            DataTable dtLs = ExecSqlQuery(selectSQL, fullFileNameToSave);
        }

        private void Write_B_KLUL(string dbName, string fileNameToSave, string orderby = "")
        {
            label2.Text = "Таблиця: - " + dbName;
            this.Text = label1.Text + " ; " + label2.Text;
            if (progressBar1.Value < progressBar1.Maximum)
                progressBar1.Value += 1;
            string fullFileNameToSave = string.Empty;
            fullFileNameToSave = filepath + "\\" + fileNameToSave;
            if (orderby.Length > 0)
                selectSQL = string.Format("SELECT * FROM {0} WHERE DISTRICT = {1} ORDER BY {2}", dbName, Raj, orderby);
            else
                selectSQL = string.Format("SELECT * FROM {0} WHERE DISTRICT = {1}", dbName, Raj);
            DataTable dtLs = ExecSqlQuery(selectSQL, fullFileNameToSave);
        }

        private void Write_B_LS_LS(string dbName, string fileNameToSave, string orderby = "")
        {
            label2.Text = "Таблиця: - " + dbName;
            this.Text = label1.Text + " ; " + label2.Text;
            progressBar1.Value += 1;
            string fullFileNameToSave = string.Empty;
            fullFileNameToSave = filepath + "\\" + fileNameToSave;
//            selectSQL = string.Format(@"select r.local_id, t.*
//from LS_LS_DATA_UPLOAD t left join reverse_migration_refbook r
//on ((floor(t.ls_nls / 10000000)) = r.code) and ((floor(t.ls_raj / 100))= r.region_code)
//where t.ls_raj = {0}
//order by t.ls_nls", Raj);
            selectSQL = string.Format(@"select r.local_id, t.*, u.ul_code
from LS_LS_DATA_UPLOAD t left join reverse_migration_refbook r
on ((floor(t.ls_nls / 10000000)) = r.code) and ((floor(t.ls_raj / 100))= r.region_code)
left join NSI_UL_UPLOAD u
on (t.ls_ul = u.ul_id)
where t.ls_raj = {0}
order by t.ls_nls", Raj);

            DataTable dtLs = ExecSqlQuery(selectSQL, fullFileNameToSave);
        }

        private static void SaveToCsv(string fullFileNameToSave, DataTable dtLs)
        {
            Encoding win1251 = Encoding.GetEncoding(1251);
            var lines = new List<string>();
            var valueLines = dtLs.AsEnumerable()
                               .Select(row => string.Join(";", row.ItemArray));
            lines.AddRange(valueLines);
            File.WriteAllLines(fullFileNameToSave, lines, win1251);
            dtLs.Clear();
        }

        /// <summary>
        /// функція для перевірки роботи модуля АСОПД
        /// </summary>
        private void CheckCppProcessed()
        {
            int allDelay = 0;
            while (true)
            {
                if (Directory.GetFiles(filepath).Length >= stPathIni.numberFilesWait) //є завдання для АСОПД, тоді чекаєм
                {
                    allDelay += stPathIni.delayMS;
                    if (allDelay > stPathIni.stopDelay)
                    {
                        //Application.Exit();
                    }
                    Application.DoEvents();
                    //Thread.Sleep(stPathIni.delayMS);
                    //Application.DoEvents();
                }
                else
                    break;
            }
        }
        protected void CloseConnection()
        {
            if (connection != null && connection.State != ConnectionState.Closed)
                connection.Close();
        }
        protected void OpenConnection()
        {
            if (connection != null && connection.State != ConnectionState.Open)
            {
                try
                {
                    connection.Open();
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }
        protected void CreateCommand()
        {
            command.CommandType = CommandType.Text;
            command.Connection = connection;
        }
        protected DataTable ExecSqlQuery(string sql_query_text, string fullNameToSave)
        {
            DataTable dt = new System.Data.DataTable();// null;
            OracleDataReader d_reader = null;
            CommandType original_com_type = command.CommandType;
            command.CommandType = CommandType.Text;
            command.Parameters.Clear();
            command.CommandText = sql_query_text;
            string sss = string.Empty;
            StringBuilder sb = new StringBuilder();
            var lines = new List<string>(1000);
            int allRecords = 0;
            try
            {
                this.OpenConnection();
                d_reader = command.ExecuteReader();
                while(d_reader.Read())
                {
                    sb.Clear();
                    for (int i = 0; i < d_reader.FieldCount; i++)
                    {
                        sb.Append(d_reader[i].ToString());
                        sb.Append(";");
                    }
                    lines.Add(sb.ToString());
                    if (lines.Count > 999)
                    {
                        AppendToCsv(lines, fullNameToSave);
                        lines.Clear();
                    }
                    Application.DoEvents();
                    allRecords++;
                }
                if (lines.Count > 0)
                {
                    AppendToCsv(lines, fullNameToSave);
                    lines.Clear();
                }
                d_reader.Close();
                //dt = new DataTable();
                //dt.Load(d_reader);
            }
            catch (Exception ex)
            {
                throw ex;
                //throw new DataAccessException(ex, string.Format("{0}.ExecSqlQuery()", this.GetType().FullName), sql_query_text, null);
            }
            finally
            {
                command.CommandType = original_com_type;
                if (d_reader != null && !d_reader.IsClosed) d_reader.Dispose();
                //if (!this.holdOpenConnection)
                this.CloseConnection();
            }
            return dt;
        }

        private void AppendToCsv(List<string> lines, string fullFileNameToSave)
        {
            Encoding win1251 = Encoding.GetEncoding(1251);
            //string fileNameToSave = string.Empty;
            //fileNameToSave = string.Format("nac.csv");
            //fullFileNameToSave = filepath + "\\" + fileNameToSave;
            File.AppendAllLines(fullFileNameToSave, lines, win1251);
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }
    }
}