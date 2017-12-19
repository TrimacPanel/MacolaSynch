using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;
using System.Net.Mail;
using System.Configuration;
using System.Net;

namespace MacolaSynch
{
    class Application
    {

        //TODO - Switch to trusted login!
        private string m_sAccessDb;
        private string m_sMacolaConn;
        private string m_sAccessConn;
        private string m_sLocationCode;

        private string m_sSmtpHost;
        private int m_iSmtpPort;
        private string m_sSmtpUsername;
        private string m_sSmtpPassword;
        private string m_sSmtpToAddress;
        private string m_sSmtpFromAddress;

        private double m_dblDecimalVariance = 2.00;

        private string m_sLogPath;

        private int m_iProdReportLag = 2;

        private List<string> alerts = new List<string>();

        private List<AlertItem> m_lAlertItems = new List<AlertItem>();

        public Application()
        {

            m_sLocationCode = Properties.Settings.Default.locationCode;
            m_sMacolaConn = Properties.Settings.Default.macolaConnection;
            m_sAccessDb = Properties.Settings.Default.accessFilePath;
            m_sLogPath = Properties.Settings.Default.logPath;

            m_sAccessConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + m_sAccessDb + ";Persist Security Info=False";

            m_sSmtpHost = Properties.Settings.Default.emailHostname;
            m_iSmtpPort = Properties.Settings.Default.emailPort;
            m_sSmtpUsername = Properties.Settings.Default.emailUsername;
            m_sSmtpPassword = Properties.Settings.Default.emailPassword;
            m_sSmtpToAddress = Properties.Settings.Default.emailSummaryTo;
            m_sSmtpFromAddress = Properties.Settings.Default.emailFrom;

            Console.WriteLine("Beginning synch...");

            DoSynch();

            SendEmailSummary();

            Console.WriteLine("Synch complete - here's a summary:\n\n");

            using (StreamWriter sw = File.CreateText(m_sLogPath))
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

        private void DoSynch()
        {

            string sql = "";
            double q1;
            double q2;
            double diff;
            bool bSafeToDelete;

            using (OleDbConnection cnAcc = new OleDbConnection(m_sAccessConn))
            {
                cnAcc.Open();

                // Query Macola for all SKUs at current location
                using (SqlConnection cnSql = new SqlConnection(m_sMacolaConn))
                {

                    string sSKU = "";

                    cnSql.Open();

                    sql = "select idx.item_no, idx.item_desc_1, idx.item_desc_2, idx.prod_cat, idx.uom, idx.item_weight_uom, idx.item_weight, idx.user_def_cd, idx.activity_cd, " +
                          "idx.cube_height_uom, idx.cube_width_uom, idx.cube_length_uom, idx.cube_height, idx.cube_width, idx.cube_length, loc.qty_on_hand, loc.qty_allocated, loc.qty_bkord " +
                          "from imitmidx_sql idx " +
                          "inner join iminvloc_sql loc on idx.item_no = loc.item_no and idx.loc = loc.loc " +
                          "where idx.loc = @LocationCode";

                    SqlCommand cmd = new SqlCommand(sql, cnSql);
                    cmd.Parameters.AddWithValue("@LocationCode", m_sLocationCode);

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
                                            "([Item], [Beg_Bal], [QTY_IN], [QTY_OUT]) values " +
                                            "(@ItemNo, 0, @Quantity, 0)";

                                        cmdInsert.Parameters.AddRange(new OleDbParameter[]
                                        {
                                            new OleDbParameter("@ItemNo", sqlize(sSKU)),
                                            new OleDbParameter("@Quantity", SafeToDouble(rdr["qty_on_hand"].ToString()))
                                        });

                                        try
                                        {
                                            cmdInsert.ExecuteNonQuery();
                                        }
                                        catch
                                        {
                                            // Just going to ignore for now - there's at least one instance of QOH records being
                                            // present without INDEX and TRX records
                                        }

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

                                        try
                                        {
                                            cmdInsert.ExecuteNonQuery();
                                        }
                                        catch
                                        {
                                            // We'll just ignore this for now, too
                                        }
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

                                // If item has an onhand in Access, it's a severe alert
                                // Otherwise, we'll just delete

                                bSafeToDelete = false;

                                using (OleDbCommand cmdQOH = cnAcc.CreateCommand())
                                {
                                    q2 = 0;
                                    cmdQOH.CommandText = "select [QTY_IN] - [QTY_OUT] as [QOH] from ItemQOH where [Item] = @ItemNo";

                                    cmdQOH.Parameters.AddRange(new OleDbParameter[]
                                    {
                                            new OleDbParameter("@ItemNo", sqlize(sSKU))
                                    });

                                    OleDbDataReader rdrQOH = cmdQOH.ExecuteReader();

                                    if (rdrQOH.Read())
                                    {
                                        q2 = rdrQOH.GetDouble(0);
                                        if (q2 == 0)
                                        {
                                            bSafeToDelete = true;
                                        }
                                    }
                                    else
                                    {
                                        // Safe if no QOH record...
                                        bSafeToDelete = true;
                                    }

                                    rdrQOH.Close();
                                }

                                if (bSafeToDelete)
                                {
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
                                    AddAlert(sSKU, "Obsolete item has QOH in Access.", AlertItem.AlertTypeEnum.Variance, AlertItem.AlertSeverityEnum.Severe, true, null, q2);
                                }
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
                                            if ((diff * 100) / q1 > m_dblDecimalVariance)
                                            {
                                                AddAlert(sSKU, "Unusual decimal QOH variance.", AlertItem.AlertTypeEnum.Variance, AlertItem.AlertSeverityEnum.Moderate, false, q1, q2);
                                            }
                                            else
                                            {
                                                AddAlert(sSKU, "Minor decimal QOH variance.", AlertItem.AlertTypeEnum.Variance, AlertItem.AlertSeverityEnum.Information, false, q1, q2);
                                            }
                                        }
                                        else
                                        {
                                            if (diff > 0)
                                            {

                                                // Does item have a pending production - if so, we ignore for now
                                                if (HasPendingProduction(sSKU))
                                                {
                                                    AddAlert(sSKU, "QOH variance with recent production: Macola is " + q1.ToString() + ", Access is " + q2.ToString(), AlertItem.AlertTypeEnum.Variance, AlertItem.AlertSeverityEnum.Information, false, q1, q2);
                                                }
                                                else if (HasRecentProduction(sSKU))
                                                {
                                                    AddAlert(sSKU, "Incomplete or unreported production: Macola is " + q1.ToString() + ", Access is " + q2.ToString(), AlertItem.AlertTypeEnum.Variance, AlertItem.AlertSeverityEnum.Moderate, false, q1, q2);
                                                }
                                                else if (HasRecentSalesOrder(sSKU))
                                                {
                                                    AddAlert(sSKU, "QOH variance with recent orders(s): Macola is " + q1.ToString() + ", Access is " + q2.ToString(), AlertItem.AlertTypeEnum.Variance, AlertItem.AlertSeverityEnum.Information, false, q1, q2);
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

        private void AddAlert(string ItemNo, string Description, AlertItem.AlertTypeEnum AlertType, AlertItem.AlertSeverityEnum Severity, bool ActionNeeded, Nullable<double> MacolaQOH = null, Nullable<double> AccessQOH = null)
        {
            m_lAlertItems.Add(new AlertItem(ItemNo, Description, AlertType, Severity, ActionNeeded, MacolaQOH, AccessQOH));
        }

        private void SendEmailSummary()
        {
            
            MailMessage msg = new MailMessage();


            string s;

            s = "";
            s = s + "<p>Here are the results of the Macola Synch for location " + m_sLocationCode + ":</p>\n";
            s = s + "<br/>";

            s = s + "<table style=\"padding:3px; border: 1px solid black; border-collapse:collapse\">";

            s = s + "<thead style=\"background-color: #15A1A2\">";
            s = s + "<th style=\"border: 1px solid black;\">SKU</th>";
            s = s + "<th style=\"border: 1px solid black;\">Description</th>";
            s = s + "<th style=\"border: 1px solid black;\">Type</th>";
            s = s + "<th style=\"border: 1px solid black;\">Severity</th>";
            s = s + "<th style=\"border: 1px solid black;\">Action Needed?</th>";
            s = s + "<th style=\"border: 1px solid black;\">M-Qty</th>";
            s = s + "<th style=\"border: 1px solid black;\">A-Qty</th>";
            s = s + "</thead>";

            IEnumerable<AlertItem> list;

            // Severe alerts section
            list = m_lAlertItems.Where(a => (a.Severity == AlertItem.AlertSeverityEnum.Severe));

            if (list.Count() > 0)
            {
                s = s + "<tr style=\"background-color: #FFAAAA\">";
                s = s + "<th colspan=\"7\" style=\"border: 1px solid black;\">Severe Alerts (Attention Required)</th>";
                s = s + "</tr>";

                foreach (var a in list)
                {
                    s = s + "<tr>";

                    s = s + "<td style=\"border: 1px solid black; text-align: left;\">" + a.ItemNo + "</td>";
                    s = s + "<td style=\"border: 1px solid black; text-align: left;\">" + a.Description + "</td>";
                    s = s + "<td style=\"border: 1px solid black; text-align: center;\">" + Enum.GetName(typeof(AlertItem.AlertTypeEnum), a.Type) + "</td>";
                    s = s + "<td style=\"border: 1px solid black; text-align: center;\">" + Enum.GetName(typeof(AlertItem.AlertSeverityEnum), a.Severity) + "</td>";
                    s = s + "<td style=\"border: 1px solid black; text-align: center;\">" + (a.ActionNeeded ? "Yes" : "No") + "</td>";
                    s = s + "<td style=\"border: 1px solid black; text-align: center;\">" + (a.MacolaQOH == null ? "-" : a.MacolaQOH.ToString()) + "</td>";
                    s = s + "<td style=\"border: 1px solid black; text-align: center;\">" + (a.AccessQOH == null ? "-" : a.AccessQOH.ToString()) + "</td>";

                    s = s + "</tr>";
                }
            }

            // Warning alerts section
            list = m_lAlertItems.Where(a => (a.Severity == AlertItem.AlertSeverityEnum.Moderate));

            if (list.Count() > 0)
            { 
                s = s + "<tr style=\"background-color: #FFA500\">";
                s = s + "<th colspan=\"7\" style=\"border: 1px solid black;\">Moderate Alerts (No Attention Required Yet)</th>";
                s = s + "</tr>";

                foreach (var a in list)
                {
                    s = s + "<tr>";

                    s = s + "<td style=\"border: 1px solid black; text-align: left;\">" + a.ItemNo + "</td>";
                    s = s + "<td style=\"border: 1px solid black; text-align: left;\">" + a.Description + "</td>";
                    s = s + "<td style=\"border: 1px solid black; text-align: center;\">" + Enum.GetName(typeof(AlertItem.AlertTypeEnum), a.Type) + "</td>";
                    s = s + "<td style=\"border: 1px solid black; text-align: center;\">" + Enum.GetName(typeof(AlertItem.AlertSeverityEnum), a.Severity) + "</td>";
                    s = s + "<td style=\"border: 1px solid black; text-align: center;\">" + (a.ActionNeeded ? "Yes" : "No") + "</td>";
                    s = s + "<td style=\"border: 1px solid black; text-align: center;\">" + (a.MacolaQOH == null ? "-" : a.MacolaQOH.ToString()) + "</td>";
                    s = s + "<td style=\"border: 1px solid black; text-align: center;\">" + (a.AccessQOH == null ? "-" : a.AccessQOH.ToString()) + "</td>";

                    s = s + "</tr>";
                }
            }

            // Informational alerts section
            list = m_lAlertItems.Where(a => (a.Severity == AlertItem.AlertSeverityEnum.Information));

            if (list.Count() > 0)
            { 
                s = s + "<tr style=\"background-color: #FFFF66\">";
                s = s + "<th colspan=\"7\" style=\"border: 1px solid black;\">Information Alerts (No Attention Required)</th>";
                s = s + "</tr>";

                IEnumerable<AlertItem> list2;

                // Adds
                list2 = m_lAlertItems.Where(a => (a.Type == AlertItem.AlertTypeEnum.Add));
                
                if (list2.Count() > 0)
                {
                    s = s + "<tr style=\"background-color: #DDDDDD\">";
                    s = s + "<th colspan=\"7\" style=\"border: 1px solid black;\">New Item(s) Added</th>";
                    s = s + "</tr>";

                    foreach (var a in list2)
                    {
                        s = s + "<tr>";

                        s = s + "<td style=\"border: 1px solid black; text-align: left;\">" + a.ItemNo + "</td>";
                        s = s + "<td style=\"border: 1px solid black; text-align: left;\">" + a.Description + "</td>";
                        s = s + "<td style=\"border: 1px solid black; text-align: center;\">" + Enum.GetName(typeof(AlertItem.AlertTypeEnum), a.Type) + "</td>";
                        s = s + "<td style=\"border: 1px solid black; text-align: center;\">" + Enum.GetName(typeof(AlertItem.AlertSeverityEnum), a.Severity) + "</td>";
                        s = s + "<td style=\"border: 1px solid black; text-align: center;\">" + (a.ActionNeeded ? "Yes" : "No") + "</td>";
                        s = s + "<td style=\"border: 1px solid black; text-align: center;\">" + (a.MacolaQOH == null ? "-" : a.MacolaQOH.ToString()) + "</td>";
                        s = s + "<td style=\"border: 1px solid black; text-align: center;\">" + (a.AccessQOH == null ? "-" : a.AccessQOH.ToString()) + "</td>";

                        s = s + "</tr>";
                    }
                }

                // Deletes
                list2 = m_lAlertItems.Where(a => (a.Type == AlertItem.AlertTypeEnum.Delete));

                if (list2.Count() > 0)
                {
                    s = s + "<tr style=\"background-color: #DDDDDD\">";
                    s = s + "<th colspan=\"7\" style=\"border: 1px solid black;\">Obsolete Item(s) Deleted</th>";
                    s = s + "</tr>";

                    foreach (var a in list2)
                    {
                        s = s + "<tr>";

                        s = s + "<td style=\"border: 1px solid black; text-align: left;\">" + a.ItemNo + "</td>";
                        s = s + "<td style=\"border: 1px solid black; text-align: left;\">" + a.Description + "</td>";
                        s = s + "<td style=\"border: 1px solid black; text-align: center;\">" + Enum.GetName(typeof(AlertItem.AlertTypeEnum), a.Type) + "</td>";
                        s = s + "<td style=\"border: 1px solid black; text-align: center;\">" + Enum.GetName(typeof(AlertItem.AlertSeverityEnum), a.Severity) + "</td>";
                        s = s + "<td style=\"border: 1px solid black; text-align: center;\">" + (a.ActionNeeded ? "Yes" : "No") + "</td>";
                        s = s + "<td style=\"border: 1px solid black; text-align: center;\">" + (a.MacolaQOH == null ? "-" : a.MacolaQOH.ToString()) + "</td>";
                        s = s + "<td style=\"border: 1px solid black; text-align: center;\">" + (a.AccessQOH == null ? "-" : a.AccessQOH.ToString()) + "</td>";

                        s = s + "</tr>";
                    }
                }

                // All others
                list2 = m_lAlertItems.Where(a => (a.Type == AlertItem.AlertTypeEnum.Variance));
                list2 = list2.Where(a => (a.Severity == AlertItem.AlertSeverityEnum.Information));

                if (list2.Count() > 0)
                {
                    s = s + "<tr style=\"background-color: #DDDDDD\">";
                    s = s + "<th colspan=\"7\" style=\"border: 1px solid black;\">Other</th>";
                    s = s + "</tr>";

                    foreach (var a in list2)
                    {
                        s = s + "<tr>";

                        s = s + "<td style=\"border: 1px solid black; text-align: left;\">" + a.ItemNo + "</td>";
                        s = s + "<td style=\"border: 1px solid black; text-align: left;\">" + a.Description + "</td>";
                        s = s + "<td style=\"border: 1px solid black; text-align: center;\">" + Enum.GetName(typeof(AlertItem.AlertTypeEnum), a.Type) + "</td>";
                        s = s + "<td style=\"border: 1px solid black; text-align: center;\">" + Enum.GetName(typeof(AlertItem.AlertSeverityEnum), a.Severity) + "</td>";
                        s = s + "<td style=\"border: 1px solid black; text-align: center;\">" + (a.ActionNeeded ? "Yes" : "No") + "</td>";
                        s = s + "<td style=\"border: 1px solid black; text-align: center;\">" + (a.MacolaQOH == null ? "-" : a.MacolaQOH.ToString()) + "</td>";
                        s = s + "<td style=\"border: 1px solid black; text-align: center;\">" + (a.AccessQOH == null ? "-" : a.AccessQOH.ToString()) + "</td>";

                        s = s + "</tr>";
                    }
                }
            }

            s = s + "</table>";

            msg.From = new MailAddress(m_sSmtpFromAddress);
            msg.To.Add(m_sSmtpToAddress);

            msg.Subject = "TMV Macola Sync Results";
            msg.IsBodyHtml = true;
            msg.Body = s;

            SmtpClient smtp = new SmtpClient();
            smtp.Host = m_sSmtpHost;
            smtp.Port = m_iSmtpPort;
            smtp.EnableSsl = true;
            smtp.Credentials = new NetworkCredential(m_sSmtpUsername, m_sSmtpPassword);

            smtp.Send(msg);

        }

        private DateTime GetPreviousWorkDay(DateTime date)
        {
            do
            {
                date = date.AddDays(-1);
            }
            while (IsHoliday(date) || IsWeekend(date));

            return date;
        }

        private bool IsHoliday(DateTime date)
        {
            //TODO - Flesh this out!
            return false;
        }

        private bool IsWeekend(DateTime date)
        {
            return date.DayOfWeek == DayOfWeek.Saturday ||
                   date.DayOfWeek == DayOfWeek.Sunday;
        }

        // Returns true if specified SKU has a recent production in Macola
        // 'Recent' will be ANY production since the start of the previous work day
        private bool HasPendingProduction(string ItemNo)
        {

            string sql = "select count(*) from iminvtrx_sql where source = 'P' and item_no = @ItemNo and promise_dt >= DATEADD(day, -1, GETDATE())";
            bool ret = false;

            DateTime cutoff = DateTime.Now;

            using (SqlConnection cn = new SqlConnection(m_sMacolaConn))
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

            /*
            DateTime cutoff = GetPreviousWorkDay(DateTime.Today);

            string sql = "select count(*) from iminvtrx_sql where source = 'P' and lev_no = 1 and item_no = @ItemNo and trx_dt + trx_tm >= @Cutoff";
            bool ret = false;

            using (SqlConnection cn = new SqlConnection(m_sMacolaConn))
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
            */
        }

        // Does SKU have a production within the past week?
        private bool HasRecentProduction(string ItemNo)
        {

            string sql = "select count(*) from iminvtrx_sql where source = 'P' and item_no = @ItemNo and trx_dt >= DATEADD(day, -7, GETDATE())";
            bool ret = false;

            DateTime cutoff = DateTime.Now;

            using (SqlConnection cn = new SqlConnection(m_sMacolaConn))
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
        private bool HasRecentSalesOrder(string ItemNo)
        {

            DateTime cutoff = GetPreviousWorkDay(DateTime.Today);

            //! - I'm using lev_no 2 at the line level here... not sure if that's right yet!
            string sql = "select count(*) from iminvtrx_sql where doc_source = 'O' and lev_no = 2 and item_no = @ItemNo and trx_dt + trx_tm >= @Cutoff";
            bool ret = false;

            using (SqlConnection cn = new SqlConnection(m_sMacolaConn))
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

        private string sqlize(string v)
        {
            return v.Replace("'", "''");
        }

        private double SafeToDouble(string v)
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

        private bool IsIgnoredItem(string ItemNo)
        {
            bool ret = false;


            if ((ItemNo == "QCPANEL") || (ItemNo == "TEST") || (ItemNo == "TEST1") || (ItemNo == "AFFILIATED NOTE"))
            {
                // Miscellaneous ignorable SKUs
                ret = true;
            }
            else if (ItemNo.Substring(ItemNo.Length - 2, 2) == "QC")
            {
                // Ignore all QC-suffixed items
                ret = true;
            }

            return ret;
        }


    }

}
