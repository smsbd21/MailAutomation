using System;
using System.Data;
using System.Net.Mail;
using System.Windows.Forms;
using System.Collections.Generic;
using oLook = Microsoft.Office.Interop.Outlook;

namespace HRMS
{
    public partial class ucMailAutomation : UserControl
    {
        ClsDbCon db = new ClsDbCon();
        static bool bSent = false;
        private static ApplicationContext _context;

        public ucMailAutomation()
        {
            InitializeComponent();
            ClsConfig.GetFxCtrl(base.Name, this);
            getMailName();
        }
        private bool getValid()
        {
            if (cmbMail.Text.Length == 0)
            {
                MessageBox.Show("Need mail name", "Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            else if (txtFrom.Text.Length == 0)
            {
                MessageBox.Show("Need From Address", "Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            else if (txtTo.Text.Length == 0)
            {
                MessageBox.Show("Need To Address", "Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            else return true;
        }
        private void getMailName()
        {
            DataSet ds = db.getDataSet("SELECT MAIL_ID,MAIL_NAME FROM TB_MAIL_AUTOMATION ORDER BY MAIL_NAME");
            if(ds.Tables[0].Rows.Count>0)
            {
                for(int i=0;i<ds.Tables[0].Rows.Count;i++)
                {
                    string key = ds.Tables[0].Rows[i]["MAIL_NAME"].ToString();
                    int value = Convert.ToInt32(ds.Tables[0].Rows[i]["MAIL_ID"]);
                    cmbMail.Items.Add(new KeyValuePair<string, int>(key, value));
                }
            }
            cmbMail.DisplayMember = "Key";
            cmbMail.ValueMember = "Value";
        }
        private void GetSendMail()
        {
            string strTo = txtTo.Text.Trim();
            string strCc = txtCc.Text.Trim();
            string strSub = txtSub.Text.Trim();
            string strBody = txtBody.Text.Trim();
            string strSql = txtQuery.Text.Trim();
            string strFrom = txtFrom.Text.Trim();

            List<string> sList = new List<string>();
            oLook.Application oApp = new oLook.Application();
            oLook.Accounts oAcc = oApp.Session.Accounts;
            foreach (oLook.Account oAc in oAcc)
            {
                sList.Add(oAc.SmtpAddress);
            }
            if (sList.Contains(strFrom.ToString()))
            {
                string sBd = "<Body><P>" + strBody + "</P><table>";
                DataSet ds = db.getDataSet(strSql);
                foreach (DataTable dt in ds.Tables)
                {
                    sBd += "<tr>\n";
                    foreach (DataColumn column in dt.Columns)
                    {
                        sBd += "<th>" + String.Format("{0:c}", column.ToString()) + "</th>\n";
                    }
                    sBd += "</tr>\n";
                }
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    if (chkType.Checked && dr["E_MAIL"].ToString() != "") this.SingleMail(oApp, dr); 
                    sBd += "<tr>";
                    for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                    {
                        sBd += "<td>" + dr[i] + "</td>\n";
                    } sBd += "</tr>\n";
                }
                oLook.MailItem mItem = (oLook.MailItem)oApp.CreateItem(oLook.OlItemType.olMailItem);
                mItem.Subject = strSub;
                sBd += "</table><br/><br/><P>Generate From NG Payroll<br/>Zaber and Zubair Fabrics Limited</P></Body>\n";
                mItem.HTMLBody = @"<HTML><head><style>table{font-size:10px;border-collapse:collapse;}
                    table,th,td{border: 1px solid black}
                    th{background-color:lightblue;}
                    td{text-align:center;max-width:100%;white-space:nowrap;}</style>
                    </head>" + sBd + "</HTML>"; 

                mItem.To = strTo; mItem.CC = strCc;
                mItem.Importance = oLook.OlImportance.olImportanceHigh;
                ((oLook._MailItem)mItem).Send();
                #region =============== [Dated On : 23-Aug-2021] ===============
                //((oLook.ItemEvents_10_Event)mItem).Close += MailItem_onClose;
                //((oLook.ItemEvents_10_Event)mItem).Send += MailItem_onSend;
                //mItem.Display(true);    // This call will make mailItem MODAL - 
                //// This way, you are not allowed to create another new mail, ob browse Outlook-Folders while mailItem is visible

                //// Using ApplicationContext will wait until your email is sent or closed without blocking other Outlook actions.
                //using (_context = new ApplicationContext())
                //{
                //    mItem.Display();
                //    Application.Run(_context);
                //}
                //if (mailWasSent) MessageBox.Show("Email was sent");
                //else MessageBox.Show("Email was NOT sent");
                #endregion
            }
        }
        private void SendAutoMail()
        {
            string strFrom = txtFrom.Text.Trim();
            List<string> sList = new List<string>();
            oLook.Application oApp = new oLook.Application();
            oLook.Accounts oAcc = oApp.Session.Accounts;
            foreach (oLook.Account oAc in oAcc)
            {
                sList.Add(oAc.SmtpAddress);
            }
            if (sList.Contains(strFrom.ToString()))
            {
                string strTo = txtTo.Text.Trim();
                string strCc = txtCc.Text.Trim();
                string strSub = txtSub.Text.Trim();
                string strSql = txtQuery.Text.Trim();
                string strBody = txtBody.Text.Trim();
                string sBd = "<Body><P>" + strBody + "</P><table>";
                MailMessage mMsg = new MailMessage();
                SmtpClient smtpClient = new SmtpClient("smtp.ionos.co.uk");//smtp server
                mMsg.From = new MailAddress("mdsumon@nomangroup.com"); //from mail address

                DataSet ds = db.getDataSet(strSql);
                foreach (DataTable dt in ds.Tables)
                {
                    sBd += "<tr>\n";
                    foreach (DataColumn column in dt.Columns)
                    {
                        sBd += "<th>" + String.Format("{0:c}", column.ToString()) + "</th>\n";
                    }
                    sBd += "</tr>\n";
                }
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    sBd += "<tr>";
                    for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                    {
                        sBd += "<td>" + dr[i] + "</td>\n";
                    } sBd += "</tr>\n";
                    if (chkType.Checked && dr["E_MAIL"].ToString() != "")
                    {
                        //this.SingleMail(oApp, dr);
                        mMsg.To.Add(dr["E_MAIL"].ToString()); mMsg.CC.Add(strTo + "," + strCc);
                        mMsg.Subject = strSub + " on : " + DateTime.Now.Date.ToString("dd-MMM-yyyy"); mMsg.IsBodyHtml = true;
                        mMsg.Body = "Dear Concern,\n\n" + txtSingBody.Text.Trim() + "\n\n\nGenerated From NG Payroll\nZaber & Zubair Fabrics Limited";
                        smtpClient.Port = 587;
                        smtpClient.Credentials = new System.Net.NetworkCredential("mdsumon@nomangroup.com", "GhgFq61sMXs1@6n=Q9P8q9V5q");//user name,password
                        smtpClient.EnableSsl = true; smtpClient.Send(mMsg);
                    }
                }
                #region =============== [Dated On : 27-Aug-2021] ===============
                //MailMessage mMsg = new MailMessage();
                //SmtpClient smtpClient = new SmtpClient("smtp.ionos.co.uk");//smtp server
                //mMsg.From = new MailAddress("mdsumon@nomangroup.com"); //from mail address

                mMsg.To.Add(strTo); // tomail address
                //mMsg.To.Add("mdsumon@nomangroup.com,smsbd9@gmail.com");// for multiple to use ,

                mMsg.CC.Add(strCc); mMsg.IsBodyHtml = true;
                mMsg.Subject = strSub + " on : " + DateTime.Now.Date.ToString("dd-MMM-yyyy");
                sBd += "</table><br/><br/><P>Generated From NG Payroll<br/>Zaber and Zubair Fabrics Limited</P></Body>\n";
                mMsg.Body = @"<HTML><head><style>table{font-size:10px;border-collapse:collapse;}
                    table,th,td{border: 1px solid black}
                    th{background-color:lightblue;}
                    td{text-align:center;max-width:100%;white-space:nowrap;}</style>
                    </head>" + sBd + "</HTML>";

                //mMsg.Attachments.Add(new Attachment(FileName));
                smtpClient.Port = 587;
                smtpClient.Credentials = new System.Net.NetworkCredential("mdsumon@nomangroup.com", "GhgFq61sMXs1@6n=Q9P8q9V5q");//user name,password
                smtpClient.EnableSsl = true; smtpClient.Send(mMsg);
                #endregion
            }
        }
        private static void MailItem_onSend(ref bool Cancel)
        {
            bSent = true;
        }
        private static void MailItem_onClose(ref bool Cancel)
        {
            _context.ExitThread();
        }
        private void btnSend_Click(object sender, EventArgs e)
        {
            this.SendAutoMail();
            //this.GetSendMail();
        }
        private void btnExit_Click(object sender, EventArgs e)
        {
            ClsConfig.GetExitUc(this);
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            string strType = chkType.Checked ? "Y" : "N";
            string strAuto = chkAuto.Checked ? "Y" : "N";
            string strQuery, strSql = txtQuery.Text.Trim().ToString();
            strSql = strSql.Replace("'", "''");
            if (getValid())
            {
                if (cmbMail.Tag == null)
                {
                    DataSet ds = db.getDataSet("SELECT NVL(MAX(MAIL_ID),0) + 1 AS MAIL_ID FROM TB_MAIL_AUTOMATION");
                    if (ds.Tables[0].Rows.Count > 0) cmbMail.Tag = ds.Tables[0].Rows[0]["MAIL_ID"].ToString();
                    strQuery = @"INSERT INTO TB_MAIL_AUTOMATION(MAIL_ID,MAIL_NAME,MAIL_TO,MAIL_FROM,MAIL_CC,MAIL_SUBJECT,MAIL_BODY,MAIL_SQL,IS_SINGLE,SINGLE_BODY,IS_AUTOMATIC,CREATED_BY) VALUES 
                            ('" + cmbMail.Tag + "','" + cmbMail.Text.Trim() + "','" + txtTo.Text.Trim() + "','" + txtFrom.Text.Trim() + "','" + txtCc.Text.Trim() + "','" + txtSub.Text.Trim() + "','" + txtBody.Text.Trim() + "','" + strSql + "','" + strType + "','" + txtSingBody.Text.Trim() + "','" + strAuto + "','" + ClsConfig.iUser + "')";
                }
                else strQuery = "UPDATE TB_MAIL_AUTOMATION SET IS_AUTOMATIC='" + strAuto + "',MAIL_TO='" + txtTo.Text.Trim().ToString() + "',MAIL_FROM='" + txtFrom.Text.Trim().ToString() + "',MAIL_CC='" + txtCc.Text.Trim().ToString() + "',MAIL_SUBJECT='" + txtSub.Text.Trim().ToString() + "',MAIL_BODY='" + txtBody.Text.Trim() + "',MAIL_SQL='" + strSql + "',IS_SINGLE='" + strType + "',SINGLE_BODY='" + txtSingBody.Text.Trim() + "',UPDATED_BY='" + ClsConfig.iUser + "',UPDATED_DATE=CURRENT_TIMESTAMP WHERE MAIL_ID='" + cmbMail.Tag + "'";
                if (db.runDmlQuery(strQuery))
                {
                    MessageBox.Show("Date Saved Successfully", "Save", MessageBoxButtons.OK);
                    ClsConfig.GetClrData(this);
                    cmbMail.Items.Clear();
                    getMailName();
                }
                else
                {
                    MessageBox.Show("Failed To Inserted");
                    ClsConfig.GetClrData(this);
                }
            }
        }
        private void btnCancel_Click(object sender, EventArgs e)
        {
            ClsConfig.GetClrData(this);
        }
        private void SingleMail(oLook.Application oApp, DataRow dr)
        {
            //oLook.Application oApp = new oLook.Application();
            oLook.MailItem mItem = (oLook.MailItem)oApp.CreateItem(oLook.OlItemType.olMailItem);
            mItem.To = dr["E_MAIL"].ToString();
            mItem.Body = "Dear Concern,\n\n" + txtSingBody.Text.Trim() + "\n\n\nGenerated From NG Payroll\nZaber & Zubair Fabrics Limited";
            mItem.Importance = oLook.OlImportance.olImportanceHigh;
            ((oLook._MailItem)mItem).Send();
            #region =============== [Dated On : 23-Aug-2021] ===============
            //((oLook.ItemEvents_10_Event)mItem).Close += MailItem_onClose;
            //((oLook.ItemEvents_10_Event)mItem).Send += MailItem_onSend;
            //// mItem.Display(true);    // This call will make mailItem MODAL - 
            //// This way, you are not allowed to create another new mail, ob browse Outlook-Folders while mailItem is visible

            /////((oLook._MailItem)mItem).Send();
            //// Using ApplicationContext will wait until your email is sent or closed without blocking other Outlook actions.
            //using (_context = new ApplicationContext())
            //{
            //    mItem.Display();
            //    Application.Run(_context);
            //}
            //if (mailWasSent) MessageBox.Show("Email was sent");
            //else MessageBox.Show("Email was NOT sent");
            #endregion
        }
        private void chkType_CheckedChanged(object sender, EventArgs e)
        {
            lbMb.Visible = false; 
            txtSingBody.Visible = false;
            if (chkType.Checked)
            {
                lbMb.Visible = true;
                txtSingBody.Visible = true;
            }
        }
        private void cmbMail_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbMail.Text.Trim().Length < 1)
            {
                cmbMail.Tag = null;
                cmbMail.SelectedIndex = -1;
            }
            else
            {
                KeyValuePair<string, int> sVal = (KeyValuePair<string, int>)cmbMail.SelectedItem;
                cmbMail.Tag = sVal.Value; cmbMail.Text = sVal.Key;
                string strSql = @"SELECT MAIL_ID,MAIL_NAME,MAIL_TO,MAIL_FROM,MAIL_CC,MAIL_SUBJECT,MAIL_BODY,
                    MAIL_SQL,IS_SINGLE,SINGLE_BODY,DAYS_DIFF,SENDING_TIME,IS_AUTOMATIC,MAIL_ATTRIBUTE,MAIL_REMARKS 
                FROM TB_MAIL_AUTOMATION WHERE MAIL_ID='" + cmbMail.Tag + "'";

                DataSet ds = db.getDataSet(strSql);
                txtTo.Text = ds.Tables[0].Rows[0]["MAIL_TO"].ToString();
                txtCc.Text = ds.Tables[0].Rows[0]["MAIL_CC"].ToString();
                txtBody.Tag = ds.Tables[0].Rows[0]["DAYS_DIFF"].ToString();
                txtFrom.Text = ds.Tables[0].Rows[0]["MAIL_FROM"].ToString();
                txtBody.Text = ds.Tables[0].Rows[0]["MAIL_BODY"].ToString();
                txtQuery.Text = ds.Tables[0].Rows[0]["MAIL_SQL"].ToString();
                txtSub.Text = ds.Tables[0].Rows[0]["MAIL_SUBJECT"].ToString();
                txtQuery.Tag = ds.Tables[0].Rows[0]["MAIL_ATTRIBUTE"].ToString();
                txtSingBody.Tag = ds.Tables[0].Rows[0]["SENDING_TIME"].ToString();
                txtSingBody.Text = ds.Tables[0].Rows[0]["SINGLE_BODY"].ToString();
                chkType.Checked = ds.Tables[0].Rows[0]["IS_SINGLE"].ToString() == "Y" ? true : false;
                chkAuto.Checked = ds.Tables[0].Rows[0]["IS_AUTOMATIC"].ToString() == "Y" ? true : false;
            }
        }
    }       
}
    

