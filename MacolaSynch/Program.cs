using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;

namespace MacolaSynch
{
    class Program
    {

        private static string MACOLA_CONN = "server=macola;database=DATA_01;user id=sa;password=trimac2k";
        private static string ACCESS_CONN = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\\MacolaSynch\\DATA FILE.mdb;Persist Security Info=False";
        private static double DECIMAL_VARIANCE = 2.00;

        private static TimeSpan WORKDAY_START_TIME = new TimeSpan(6, 0, 0);
        private static TimeSpan WORKDAY_END_TIME = new TimeSpan(16, 0, 0);

        private static DateTime SYNCH_FROM;
        private static List<string> alerts = new List<string>();

        private static List<AlertItem> m_lAlertItems = new List<AlertItem>();

        static void Main(string[] args)
        {
            Application app = new Application();
        }

        static void xMain(string[] args)
        {

            Console.WriteLine("Beginning synch...");

            SYNCH_FROM = new DateTime(1990, 1, 1);

            DoSynch();

            Console.WriteLine("Synch complete - here's a summary:\n\n");            

            using (StreamWriter sw = File.CreateText("c:\\MacolaSynch\\log.txt"))
            {
                Console.WriteLine("SKU\tDescription\tAlert Type\tAlert Severity\tAction Needed");
                sw.WriteLine("SKU\tDescription\tAlert Type\tAlert Severity\tAction Needed");

                foreach (AlertItem a in m_lAlertItems)
                {
                    Console.WriteLine(a.ItemNo + "\t" + a.Description + "\t" + Enum.GetName(typeof(AlertItem.AlertTypeEnum), a.Type)
                            + "\t" + Enum.GetName(typeof(AlertItem.AlertSeverityEnum), a.Severity) + "\t" + (a.ActionNeeded ? "Yes" : "No"));

                    sw.WriteLine(a.ItemNo + "\t" + a.Description + "\t" + Enum.GetName(typeof(AlertItem.AlertTypeEnum), a.Type)
                            + "\t" + Enum.GetName(typeof(AlertItem.AlertSeverityEnum), a.Severity) + "\t" + (a.ActionNeeded ? "Yes" : "No"));
                }
            }

            Console.WriteLine("\n\nEnd of summary, press any key to continue.");
            Console.ReadKey();
        }

        static void DoSynch()
        {

            string sql = "";
            double q1;
            double q2;
            double diff;

            using (OleDbConnection cnAcc = new OleDbConnection(ACCESS_CONN))
            {
                cnAcc.Open();

                // Query Macola for all SKUs at current location
                using (SqlConnection cnSql = new SqlConnection(MACOLA_CONN))
                {

                    string sSKU = "";

                    cnSql.Open();

                    sql = "select idx.item_no, idx.item_desc_1, idx.item_desc_2, idx.prod_cat, idx.uom, idx.item_weight_uom, idx.item_weight, idx.user_def_cd, idx.activity_cd, " +
                          "idx.cube_height_uom, idx.cube_width_uom, idx.cube_length_uom, idx.cube_height, idx.cube_width, idx.cube_length, loc.qty_on_hand, loc.qty_allocated, loc.qty_bkord " +
                          "from imitmidx_sql idx " +
                          "inner join iminvloc_sql loc on idx.item_no = loc.item_no and idx.loc = loc.loc " +
                          "where idx.loc = 'TMV'";

                    SqlCommand cmd = new SqlCommand(sql, cnSql);
                    SqlDataReader rdr = cmd.ExecuteReader();

                    OleDbCommand cmdAcc;
                    OleDbDataReader rdrAcc;

                    // Iterate all Macola items and process accordingly
                    while (rdr.Read())
                    {
                        sSKU = rdr["item_no"].ToString().Trim();

                        // Does item exist in access?
                        sql = "select count(*) as the_count from ItemINDEX where Item = '" + sSKU + "'";

                        cmdAcc = new OleDbCommand(sql, cnAcc);

                        rdrAcc = cmdAcc.ExecuteReader();
                        rdrAcc.Read();


                        // Item doesn't exist in access..
                        if (rdrAcc.GetInt32(0) == 0)
                        {
                            if (rdr["activity_cd"].ToString() == "A")
                            {

                                if (!IsIgnoredItem(sSKU))
                                {

                                    using (OleDbCommand cmdInsert = cnAcc.CreateCommand())
                                    {

                                        cmdInsert.CommandText = sql;


                                        cmdInsert.CommandText = "insert into [ItemINDEX] " +
                                                                "([Item], [Desc], [Cat], [UOM], [Wt_UOM], [Wt], [User_Defined_Code], [CH_UOM], [CW_UOM], [CL_UOM], [CH], [CW], [CL]) values " +
                                                                "(@ItemNo, @Description, @Category, @UOM, @WeightUOM, @Weight, @UserCode, @CHUOM, @CWUOM, @CLUOM, @CH, @CW, @CL)";

                                        cmdInsert.Parameters.AddRange(new OleDbParameter[]
                                        {
                                            new OleDbParameter("@ItemNo", sqlize(sSKU)),
                                            new OleDbParameter("@Description", sqlize(rdr["item_desc_1"].ToString().Trim())),
                                            new OleDbParameter("@Category", sqlize(rdr["prod_cat"].ToString().Trim())),
                                            new OleDbParameter("@UOM", sqlize(rdr["uom"].ToString())),
                                            new OleDbParameter("@WeightUOM", sqlize(rdr["item_weight_uom"].ToString().Trim())),
                                            new OleDbParameter("@Weight", Convert.ToDouble(rdr["item_weight"].ToString().Trim())),
                                            new OleDbParameter("@UserCode", sqlize(rdr["user_def_cd"].ToString().Trim())),
                                            new OleDbParameter("@CHUOM", sqlize(rdr["cube_height_uom"].ToString().Trim())),
                                            new OleDbParameter("@CWUOM", sqlize(rdr["cube_width_uom"].ToString().Trim())),
                                            new OleDbParameter("@CLUOM", sqlize(rdr["cube_length_uom"].ToString().Trim())),
                                            new OleDbParameter("@CH", SafeToDouble(rdr["cube_height"].ToString())),
                                            new OleDbParameter("@CW", SafeToDouble(rdr["cube_width"].ToString())),
                                            new OleDbParameter("@CL", SafeToDouble(rdr["cube_length"].ToString()))
                                        });

                                        cmdInsert.ExecuteNonQuery();

                                    }


                                    // Create  ItemQOH record
                                    using (OleDbCommand cmdInsert = cnAcc.CreateCommand())
                                    {
                                        cmdInsert.CommandText = "insert into [ItemQOH] " +
                                            "([Item], [QTY_IN]) values " +
                                            "(@ItemNo, @Quantity)";

                                        cmdInsert.Parameters.AddRange(new OleDbParameter[]
                                        {
                                            new OleDbParameter("@ItemNo", sqlize(sSKU)),
                                            new OleDbParameter("@Quantity", SafeToDouble(rdr["qty_on_hand"].ToString()))
                                        });

                                        cmdInsert.ExecuteNonQuery();

                                    }

                                    // Create initial ItemTRX record
                                    using (OleDbCommand cmdInsert = cnAcc.CreateCommand())
                                    {
                                        cmdInsert.CommandText = "insert into [ItemTRX] " +
                                            "([Date], [Item], [Qty_IN], [Notes]) values " +
                                            "(@Date, @ItemNo, @Quantity, @Notes)";

                                        cmdInsert.Parameters.AddWithValue("@Date", DateTime.Now.ToString("MM/dd/yyyy"));
                                        cmdInsert.Parameters.AddRange(new OleDbParameter[]
                                        {
                                            //new OleDbParameter("@Date", "#" + DateTime.Now.ToString("MM/dd/yyyy") + "#"),                                            
                                            new OleDbParameter("@ItemNo", sqlize(sSKU)),
                                            new OleDbParameter("@Quantity", SafeToDouble(rdr["qty_on_hand"].ToString())),
                                            new OleDbParameter("@Notes", "SYNCHED FROM MACOLA")
                                        });

                                        cmdInsert.ExecuteNonQuery();

                                    }

                                    AddAlert(sSKU, "Item added to Access.", AlertItem.AlertTypeEnum.Add, AlertItem.AlertSeverityEnum.Information, false, SafeToDouble(rdr["qty_on_hand"].ToString()), SafeToDouble(rdr["qty_on_hand"].ToString()));
                                }

                            }

                        }
                        else  // Item does exist in Access..
                        {
                            // Item is obsolete, remove from Access
                            if (rdr["activity_cd"].ToString() == "O")
                            {

                                //TODO:  Add check for existing qty on hand

                                // Delete item's transaction history
                                sql = "delete * from ItemTrx where Item = '" + sSKU + "'";
                                cmdAcc = new OleDbCommand(sql, cnAcc);
                                cmdAcc.ExecuteNonQuery();

                                // Delete ItemQOH
                                sql = "delete * from ItemQOH where Item = '" + sSKU + "'";
                                cmdAcc = new OleDbCommand(sql, cnAcc);
                                cmdAcc.ExecuteNonQuery();

                                // Delete ItemINDEX
                                sql = "delete * from ItemINDEX where Item = '" + sSKU + "'";
                                cmdAcc = new OleDbCommand(sql, cnAcc);
                                cmdAcc.ExecuteNonQuery();

                                AddAlert(sSKU, "Obsolete item deleted from Access.", AlertItem.AlertTypeEnum.Delete, AlertItem.AlertSeverityEnum.Information, false, null, null);
                            }
                            else
                            {

                                // Item exists and is still active - refresh SKU data
                                using (OleDbCommand cmdUpdate = cnAcc.CreateCommand())
                                {
                                    cmdUpdate.CommandText = "update [ItemINDEX] set " +
                                                            "[Desc] = @Description, " +
                                                            "[Cat] = @Category, " +
                                                            "[UOM] = @UOM, " +
                                                            "[Wt_UOM] = @WeightUOM, " +
                                                            "[Wt] = @Weight, " +
                                                            "[User_Defined_Code] = @UserCode, " +
                                                            "[CH_UOM] = @CHUOM, " +
                                                            "[CW_UOM] = @CWUOM, " +
                                                            "[CL_UOM] = @CLUOM, " +
                                                            "[CH] = @CH, " +
                                                            "[CW] = @CW, " +
                                                            "[CL] = @CL " +
                                                            "where [Item] = '" + sSKU + "'";

                                    cmdUpdate.Parameters.AddRange(new OleDbParameter[]
                                    {
                                        new OleDbParameter("@Description", sqlize(rdr["item_desc_1"].ToString())),
                                        new OleDbParameter("@Category", sqlize(rdr["prod_cat"].ToString())),
                                        new OleDbParameter("@UOM", sqlize(rdr["uom"].ToString())),
                                        new OleDbParameter("@WeightUOM", sqlize(rdr["item_weight_uom"].ToString())),
                                        new OleDbParameter("@Weight", Convert.ToDouble(rdr["item_weight"].ToString())),
                                        new OleDbParameter("@UserCode", sqlize(rdr["user_def_cd"].ToString())),
                                        new OleDbParameter("@CHUOM", sqlize(rdr["cube_height_uom"].ToString())),
                                        new OleDbParameter("@CWUOM", sqlize(rdr["cube_width_uom"].ToString())),
                                        new OleDbParameter("@CLUOM", sqlize(rdr["cube_length_uom"].ToString())),
                                        new OleDbParameter("@CH", SafeToDouble(rdr["cube_height"].ToString())),
                                        new OleDbParameter("@CW", SafeToDouble(rdr["cube_width"].ToString())),
                                        new OleDbParameter("@CL", SafeToDouble(rdr["cube_length"].ToString()))
                                    });

                                    cmdUpdate.ExecuteNonQuery();
                                }

                                //alerts.Add("UPDATE\t" + sSKU);

                                // Finally, let's do our qty comparisons and build alerts
                                using (OleDbCommand cmdQOH = cnAcc.CreateCommand())
                                {
                                    sql = "select QTY_IN - QTY_OUT as QOH from ItemQOH where Item = '" + sSKU + "'";

                                    cmdQOH.CommandText = sql;

                                    OleDbDataReader rdrQOH = cmdQOH.ExecuteReader();

                                    if (rdrQOH.Read())
                                    {

                                        q1 = SafeToDouble(rdr["qty_on_hand"].ToString());

                                        //q1 = SafeToDouble(rdr["qty_on_hand"].ToString()) + SafeToDouble(rdr["qty_allocated"].ToString());
                                        q2 = SafeToDouble(rdrQOH["QOH"].ToString());
                                        diff = Math.Abs(q1 - q2);

                                        // If Macola qty contains a decimal, we'll allow some variance before alerting
                                        if (q1 % 1 > 0)
                                        {
                                            if ((diff * 100) / q1 > DECIMAL_VARIANCE)
                                            {
                                                alerts.Add("DIFF " + diff.ToString() + " QOH\t" + sSKU + "\t" + q1.ToString() + " to " + q2.ToString());
                                                AddAlert(sSKU, "QOH Variance.", AlertItem.AlertTypeEnum.Add, AlertItem.AlertSeverityEnum.Information, false, q1, q2);
                                            }
                                            else
                                            {
                                                AddAlert(sSKU, "Non-critical QOH variance.", AlertItem.AlertTypeEnum.Variance, AlertItem.AlertSeverityEnum.Information, false, q1, q2);
                                            }
                                        }
                                        else
                                        {
                                            if (diff > 0)
                                            {

                                                // Does item have a 'recent' production - if so, we ignore for now
                                                if (HasRecentProduction(sSKU))
                                                {

                                                    AddAlert(sSKU, "QOH variance with recent production: Macola is " + q1.ToString() + ", Access is " + q2.ToString(), AlertItem.AlertTypeEnum.Variance, AlertItem.AlertSeverityEnum.Information, false, q1, q2);
                                                }
                                                else if (HasRecentSalesOrder(sSKU))
                                                {
                                                    AddAlert(sSKU, "QOH variance with recent sale(s): Macola is " + q1.ToString() + ", Access is " + q2.ToString(), AlertItem.AlertTypeEnum.Variance, AlertItem.AlertSeverityEnum.Information, false, q1, q2);
                                                }
                                                else
                                                {
                                                    AddAlert(sSKU, "Unexplained QOH variance: Macola is " + q1.ToString() + ", Access is " + q2.ToString(), AlertItem.AlertTypeEnum.Variance, AlertItem.AlertSeverityEnum.Severe, true, q1, q2);
                                                }
                                            }
                                        }
                                    }

                                    rdrQOH.Close();
                                }

                            }

                        }

                        rdrAcc.Close();

                    }

                    rdr.Close();

                    Console.WriteLine("Synch completed.");
                }

                cnAcc.Close();

            }
        }

        static void AddAlert(string ItemNo, string Description, AlertItem.AlertTypeEnum AlertType, AlertItem.AlertSeverityEnum Severity, bool ActionNeeded, Nullable<double> MacolaQOH, Nullable<double> AccessQOH)
        {
            m_lAlertItems.Add(new AlertItem(ItemNo, Description, AlertType, Severity, ActionNeeded, MacolaQOH, AccessQOH));
        }

        static DateTime GetPreviousWorkDay(DateTime date)
        {
            do
            {
                date = date.AddDays(-1);
            }
            while (IsHoliday(date) || IsWeekend(date));
        
            return date;
        }

        static bool IsHoliday(DateTime date)
        {
            //TODO - Flesh this out!
            return false;
        }

        static bool IsWeekend(DateTime date)
        {
            return date.DayOfWeek == DayOfWeek.Saturday ||
                   date.DayOfWeek == DayOfWeek.Sunday;
        }

        // Returns true if specified SKU has a recent production in Macola
        // 'Recent' will be ANY production since the start of the previous work day
        static bool HasRecentProduction(string ItemNo)
        {

            DateTime cutoff = GetPreviousWorkDay(DateTime.Today);
            cutoff.Add(WORKDAY_START_TIME);

            string sql = "select count(*) from iminvtrx_sql where source = 'P' and lev_no = 1 and item_no = @ItemNo and trx_dt + trx_tm >= @Cutoff";
            bool ret = false;

            using (SqlConnection cn = new SqlConnection(MACOLA_CONN))
            {
                SqlCommand cmd = new SqlCommand(sql, cn);

                cmd.Parameters.AddWithValue("@ItemNo", ItemNo);
                cmd.Parameters.AddWithValue("@Cutoff", cutoff);

                cn.Open();

                SqlDataReader rdr = cmd.ExecuteReader();

                if (rdr.Read())
                {
                    if (rdr.GetInt32(0) > 0)
                    {
                        ret = true;
                    }
                    else
                    {
                        ret = false;
                    }
                }
                //TODO -  err!

                rdr.Close();

            }

            return ret;
        }

        // Returns true if specified SKU has a recent sales order in Macola
        // 'Recent' will be ANY production since the start of the previous work day
        static bool HasRecentSalesOrder(string ItemNo)
        {

            DateTime cutoff = GetPreviousWorkDay(DateTime.Today);
            cutoff.Add(WORKDAY_START_TIME);

            //! - I'm using lev_no 2 at the line level here... not sure if that's right yet!
            string sql = "select count(*) from iminvtrx_sql where doc_source = 'O' and lev_no = 2 and item_no = @ItemNo and trx_dt + trx_tm >= @Cutoff";
            bool ret = false;

            using (SqlConnection cn = new SqlConnection(MACOLA_CONN))
            {
                SqlCommand cmd = new SqlCommand(sql, cn);

                cmd.Parameters.AddWithValue("@ItemNo", ItemNo);
                cmd.Parameters.AddWithValue("@Cutoff", cutoff);

                cn.Open();

                SqlDataReader rdr = cmd.ExecuteReader();

                if (rdr.Read())
                {
                    if (rdr.GetInt32(0) > 0)
                    {
                        ret = true;
                    }
                    else
                    {
                        ret = false;
                    }
                }
                //TODO -  err!

                rdr.Close();

            }

            return ret;
        }
  
        static string sqlize(string v)
        {
            return v.Replace("'", "''");
        }
        
        static double SafeToDouble(string v)
        {
            if (v == "")
            {
                return 0;
            }
            else
            {
                return Convert.ToDouble(v);
            }
        }

        static bool IsIgnoredItem(string ItemNo)
        {
            bool ret = false;

            
            if ((ItemNo == "QCPANEL") || (ItemNo == "TEST") || (ItemNo == "TEST1"))
            {
                ret = true;
            }

            return ret;
        }


    }
}
