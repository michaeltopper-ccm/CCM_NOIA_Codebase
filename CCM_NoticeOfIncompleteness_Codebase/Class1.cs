
//CCM NOIA Codebase which helps with the functionality behind the
//Notice of Incompleteness input form in Encompass
//Last Revision 12/14/2023 by Michael Topper
using EllieMae.Encompass.Automation;
using EllieMae.Encompass.BusinessObjects;
using EllieMae.Encompass.BusinessObjects.Users;
using EllieMae.Encompass.Forms;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using Microsoft.Office.Interop.Word;
using EllieMae.Encompass.BusinessObjects.Loans;
using System.IO;
using EllieMae.Encompass.BusinessObjects.Loans.Logging;
using System.Linq;
using System.Configuration;

namespace CCM_NoticeOfIncompleteness_Codebase
{
    public class CCM_NoticeOfIncompleteness_Codebase : Form
    {
        private List<ControlPair> _ControlList;
        private static Panel _MainPanel;
        //public static Button _BtnPrint;
        //public static DropdownBox DDBorrowerFocusSelection;
        public static string LastBorrower;
        public static string MonitoredFields;
        public static bool WriteNOIAHistory;
        public static bool InitializeLoanHandlers;
        public override void CreateControls()
        {
            //Disable the form if loan is read only
            if (!EncompassApplication.CurrentLoan.Locked)
            {
                this.Form.Enabled = false;
            }

            //Find the Main Panel where all our checkboxes and dropdowns exist            
            _MainPanel = (Panel)FindControl("pnlForm");

            //Setup for the Print Button - subscribes to its click event
            //_BtnPrint = (Button)FindControl("btnPrint");
            //_BtnPrint.Click += BtnPrint_Click;

            //setup for the clear button
            Button btnClear = (Button)FindControl("btnClear");
            btnClear.Click += BtnClear_Click;


            //Disable print button if no options selected on form load
            TestPrintAccess();
            BuildMonitoredFieldlist();
            RefreshNOIAList();

            EncompassApplication.CurrentLoan.LogEntryAdded += NOIAScreen_LogEntryAdded;
            EncompassApplication.CurrentLoan.LogEntryChange += NOIAScreen_LogEntryChange;
            EncompassApplication.CurrentLoan.FieldChange += NOIAScreen_FieldChangeHandler;
            if (!InitializeLoanHandlers)
            {
                EncompassApplication.CurrentLoan.BeforeCommit += NOIAScreen_BeforeCommit;
                EncompassApplication.LoanClosing += NOIAScreen_LoanClosing;
                InitializeLoanHandlers = true;
            }
            this.Form.Unload += NOIAScreen_FormUnload;
        }

        private void NOIAScreen_LoanClosing(object sender, EventArgs e)
        {
            if (InitializeLoanHandlers)
            {
                EncompassApplication.CurrentLoan.BeforeCommit -= NOIAScreen_BeforeCommit;
                EncompassApplication.LoanClosing -= NOIAScreen_LoanClosing;
            }
        }

        private void NOIAScreen_BeforeCommit(object source, CancelableEventArgs e)
        {
            if (WriteNOIAHistory)
            {
                WriteHistory();
            }
        }

        private void NOIAScreen_LogEntryAdded(object source, LogEntryEventArgs e)
        {
            if (e.LogEntry.EntryType.Equals(LogEntryType.TrackedDocument))
            {
                TrackedDocument doc = (TrackedDocument)e.LogEntry;

                if (doc.Title.Equals("Notice of Incomplete Application"))
                {
                    if (Convert.ToDateTime(doc.DateAdded) >= Convert.ToDateTime(EncompassApplication.CurrentLoan.LastModified))
                    {
                        WriteNOIAHistory = true;
                    }
                }
            }
        }
        private void NOIAScreen_LogEntryChange(object source, LogEntryEventArgs e)
        {
            if (e.LogEntry.EntryType.Equals(LogEntryType.TrackedDocument))
            {
                TrackedDocument doc = (TrackedDocument)e.LogEntry;

                if (doc.Title.Equals("Notice of Incomplete Application"))
                {
                    if (Convert.ToDateTime(doc.DateAdded) >= Convert.ToDateTime(EncompassApplication.CurrentLoan.LastModified))
                    {
                        WriteNOIAHistory = true;
                    }
                }
            }
        }

        private void BuildMonitoredFieldlist()
        {
            MonitoredFields = "CX.NOIA.ASSETS.1, CX.NOIA.ASSETS.2, CX.NOIA.ASSETS.3, CX.NOIA.ASSETS.4, CX.NOIA.ASSETS.5, CX.NOIA.ASSETS.OTHER.1, CX.NOIA.ASSETS.OTHER.2";
            MonitoredFields = MonitoredFields + ", CX.NOIA.CREDIT.1, CX.NOIA.CREDIT.2, CX.NOIA.CREDIT.3, CX.NOIA.CREDIT.4, CX.NOIA.CREDIT.5, CX.NOIA.CREDIT.6, CX.NOIA.CREDIT.7" +
                ", CX.NOIA.CREDIT.8, CX.NOIA.CREDIT.9, CX.NOIA.CREDIT.10, CX.NOIA.CREDIT.11, CX.NOIA.CREDIT.12, CX.NOIA.CREDIT.OTHER.1, CX.NOIA.CREDIT.OTHER.2";
            MonitoredFields = MonitoredFields + ", CX.NOIA.INCOME.1, CX.NOIA.INCOME.2, CX.NOIA.INCOME.3, CX.NOIA.INCOME.4, CX.NOIA.INCOME.5, CX.NOIA.INCOME.6, CX.NOIA.INCOME.7" +
                ", CX.NOIA.INCOME.8, CX.NOIA.INCOME.9, CX.NOIA.INCOME.10, CX.NOIA.INCOME.COMMENTS.1, CX.NOIA.INCOME.COMMENTS.2";
            MonitoredFields = MonitoredFields + ", CX.NOIA.MISC.1, CX.NOIA.MISC.2, CX.NOIA.MISC.3, CX.NOIA.MISC.4, CX.NOIA.MISC.5, CX.NOIA.MISC.6, CX.NOIA.MISC.COMMENTS.1, CX.NOIA.MISC.COMMENTS.2";


        }

        private void NOIAScreen_FormUnload(object sender, EventArgs e)
        {
            EncompassApplication.CurrentLoan.LogEntryAdded -= NOIAScreen_LogEntryAdded;
            EncompassApplication.CurrentLoan.LogEntryChange -= NOIAScreen_LogEntryChange;
            EncompassApplication.CurrentLoan.FieldChange -= NOIAScreen_FieldChangeHandler;
            this.Form.Unload -= NOIAScreen_FormUnload;
        }

        private void NOIAScreen_FieldChangeHandler(object source, FieldChangeEventArgs e)
        {
            if (MonitoredFields.Contains(e.FieldID))
            {
                RefreshNOIAList();
            }
        }

        private void ClearUI()
        {
            foreach (Control c in _MainPanel.Controls)
            {
                if (c.ControlID.Equals("ddBorrowerFocusSelection"))
                {
                    continue;
                }
                else if (c.ControlID.StartsWith("cb"))
                {
                    EllieMae.Encompass.Forms.CheckBox thisCheckBox = (EllieMae.Encompass.Forms.CheckBox)c;
                    thisCheckBox.Checked = false;
                }
                else if (c.ControlID.StartsWith("txt") && !c.ControlID.Equals("txtHistory"))
                {
                    EllieMae.Encompass.Forms.TextBox thisTextBox = (EllieMae.Encompass.Forms.TextBox)c;
                    thisTextBox.Text = "";
                }
                else if (c.ControlID.StartsWith("dd"))
                {
                    EllieMae.Encompass.Forms.DropdownBox thisDropdownBox = (EllieMae.Encompass.Forms.DropdownBox)c;
                    thisDropdownBox.SelectedIndex = 0;
                }
            }

            //_BtnPrint.Enabled = false;
        }

        private void BtnClear_Click(object sender, EventArgs e)
        {
            ClearUI();
        }

        private void TestPrintAccess()
        {
            if (!Macro.GetField("CX.NOIA.CONTENT").Equals(""))
            {
                //_BtnPrint.Enabled = true;
            }
        }

        private void BtnPrint_Click(object sender, EventArgs e)
        {
            //hand the NOIA data off to the doc printer methods
            RefreshNOIAList();

            FetchNOIADoc();

            WriteHistory();
            
        }

        private void RefreshNOIAList()
        {
            string entries = string.Empty;

            if (Macro.GetField("CX.NOIA.ASSETS.1").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Bank Statement(s) (all pages required)" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Bank Statement(s) (all pages required)" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.ASSETS.2").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Business Bank Statement(s) (all pages required)" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Business Bank Statement(s) (all pages required)" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.ASSETS.3").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Earnest Money Check or Wire Verification" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Earnest Money Check or Wire Verification" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.ASSETS.4").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Proof of Gift Funds" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Proof of Gift Funds" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.ASSETS.5").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Retirement/Investment Statements" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Retirement/Investment Statements" + Environment.NewLine;
                }
            }
            if (!Macro.GetField("CX.NOIA.ASSETS.OTHER.1").Equals(""))
            {
                if (entries.Equals(""))
                {
                    entries = "- "+Macro.GetField("CX.NOIA.ASSETS.OTHER.1") + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- "+Macro.GetField("CX.NOIA.ASSETS.OTHER.1") + Environment.NewLine;
                }
            }
            if (!Macro.GetField("CX.NOIA.ASSETS.OTHER.2").Equals(""))
            {
                if (entries.Equals(""))
                {
                    entries = "- "+Macro.GetField("CX.NOIA.ASSETS.OTHER.2") + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- "+Macro.GetField("CX.NOIA.ASSETS.OTHER.2") + Environment.NewLine;
                }
            }

            if (Macro.GetField("CX.NOIA.CREDIT.1").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Bankruptcy Paperwork" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Bankruptcy Paperwork" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.CREDIT.2").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Child Support Order" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Child Support Order" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.CREDIT.3").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Divorce Decree and/or Separation Agreement" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Divorce Decree and/or Separation Agreement" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.CREDIT.4").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Hazard Insurance Policy" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Hazard Insurance Policy" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.CREDIT.5").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Homeowners Association Dues" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Homeowners Association Dues" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.CREDIT.6").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Lease Agreement(s)" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Lease Agreement(s)" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.CREDIT.7").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Mortgage Note" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Mortgage Note" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.CREDIT.8").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Mortgage Statement(s)" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Mortgage Statement(s)" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.CREDIT.9").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Property Tax Bill" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Property Tax Bill" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.CREDIT.10").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Student Loan Statement(s)" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Student Loan Statement(s)" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.CREDIT.11").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Trust Agreement" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Trust Agreement" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.CREDIT.12").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- VA Certificate of Eligibility" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- VA Certificate of Eligibility" + Environment.NewLine;
                }
            }
            if (!Macro.GetField("CX.NOIA.CREDIT.OTHER.1").Equals(""))
            {
                if (entries.Equals(""))
                {
                    entries = "- "+Macro.GetField("CX.NOIA.CREDIT.OTHER.1") + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- "+Macro.GetField("CX.NOIA.CREDIT.OTHER.1") + Environment.NewLine;
                }
            }
            if (!Macro.GetField("CX.NOIA.CREDIT.OTHER.2").Equals(""))
            {
                if (entries.Equals(""))
                {
                    entries = "- "+Macro.GetField("CX.NOIA.CREDIT.OTHER.2") + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- "+Macro.GetField("CX.NOIA.CREDIT.OTHER.2") + Environment.NewLine;
                }
            }

            if (Macro.GetField("CX.NOIA.INCOME.1").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Balance Sheet - Signed & Dated" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Balance Sheet - Signed & Dated" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.INCOME.2").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Business Tax Return(s) - Signed & Dated" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Business Tax Return(s) - Signed & Dated" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.INCOME.3").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Disability Income Award Letter: Most recent" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Disability Income Award Letter: Most recent" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.INCOME.4").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- K-1(s) for all Businesses" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- K-1(s) for all Businesses" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.INCOME.5").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Paystub(s)" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Paystub(s)" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.INCOME.6").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Pension/Retirement Statement" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Pension/Retirement Statement" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.INCOME.7").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Personal Tax Return(s) - Signed & Dated" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Personal Tax Return(s) - Signed & Dated" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.INCOME.8").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Profit & Loss - Signed & Dated" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Profit & Loss - Signed & Dated" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.INCOME.9").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Social Security Awards Letter" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Social Security Awards Letter" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.INCOME.10").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- W-2(s)" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- W-2(s)" + Environment.NewLine;
                }
            }
            if (!Macro.GetField("CX.NOIA.INCOME.OTHER.1").Equals(""))
            {
                if (entries.Equals(""))
                {
                    entries = "- "+Macro.GetField("CX.NOIA.INCOME.OTHER.1") + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- "+Macro.GetField("CX.NOIA.INCOME.OTHER.1") + Environment.NewLine;
                }
            }
            if (!Macro.GetField("CX.NOIA.INCOME.OTHER.2").Equals(""))
            {
                if (entries.Equals(""))
                {
                    entries = "- "+Macro.GetField("CX.NOIA.INCOME.OTHER.2") + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- "+Macro.GetField("CX.NOIA.INCOME.OTHER.2") + Environment.NewLine;
                }
            }

            if (Macro.GetField("CX.NOIA.MISC.1").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Appraisal Requirements Outstanding" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Appraisal Requirements Outstanding" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.MISC.2").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Construction Plans & Specifications" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Construction Plans & Specifications" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.MISC.3").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Itemized Budget & Draw Schedule" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Itemized Budget & Draw Schedule" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.MISC.4").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Initial Disclosures: Fully executed by all applicant(s)" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Initial Disclosures: Fully executed by all applicant(s)" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.MISC.5").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Purchase Contract: Fully executed with all applicable addendum(s)" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Purchase Contract: Fully executed with all applicable addendum(s)" + Environment.NewLine;
                }
            }
            if (Macro.GetField("CX.NOIA.MISC.6").Equals("X"))
            {
                if (entries.Equals(""))
                {
                    entries = "- Statement of Repairs & Bids" + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- Statement of Repairs & Bids" + Environment.NewLine;
                }
            }
            if (!Macro.GetField("CX.NOIA.MISC.OTHER.1").Equals(""))
            {
                if (entries.Equals(""))
                {
                    entries = "- "+Macro.GetField("CX.NOIA.MISC.OTHER.1") + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- "+Macro.GetField("CX.NOIA.MISC.OTHER.1") + Environment.NewLine;
                }
            }
            if (!Macro.GetField("CX.NOIA.MISC.OTHER.2").Equals(""))
            {
                if (entries.Equals(""))
                {
                    entries = "- "+Macro.GetField("CX.NOIA.MISC.OTHER.2") + Environment.NewLine;
                }
                else
                {
                    entries = entries + "- "+Macro.GetField("CX.NOIA.MISC.OTHER.2") + Environment.NewLine;
                }
            }

            Macro.SetFieldNoRules("CX.NOIA.CONTENT", entries);
            TestPrintAccess();
        }

        private void WriteHistory()
        {

            //string NOIA_EntryHistory = Macro.GetField("CX.NOIA.HISTORY");

            //string newHistory = "NOIA Added On " + System.DateTime.Now + " by " + EncompassApplication.CurrentUser.ID + Environment.NewLine + Macro.GetField("CX.NOIA.CONTENT") + Environment.NewLine + Environment.NewLine; ;

            //newHistory = newHistory +  Environment.NewLine + NOIA_EntryHistory;

            //Macro.SetFieldNoRules("CX.NOIA.HISTORY", newHistory);
        }

        private void FetchNOIADoc()
        {
            
            //getting the AutoDocs Global CDO 
            List<AutoDoc> docList = new List<AutoDoc>();
            var docs = EncompassApplication.CurrentUser.Session.DataExchange.GetCustomDataObject("AutoDocs");
            if (docs != null)
            {
                //using Newtonsoft to convert my Ellie Mae data object back into a custom class
                string json;
                json = Encoding.UTF8.GetString(docs.Data);
                try
                {
                    //transforming json string into List of AutoDoc
                    docList = JsonConvert.DeserializeObject<List<AutoDoc>>(json);
                    //go through each Auto-doc and find the NOIA doc
                    foreach (AutoDoc doc in docList)
                    {
                        if (doc.TriggerCode == "NOIA")
                        {
                            //convert it from a byte array back into a word doc and store it in the SmartClient Cache
                            string fullPath = @"C:\SmartClientCache\" + doc.OriginalFileName;
                            byte[] byteArray = doc.Content;
                            File.WriteAllBytes(fullPath, byteArray);

                            //Merges data from the loan into the Word doc
                            MergeDataIntoTempDoc(fullPath);
                            string fullPathAsPDF = fullPath.Replace(".docx", ".pdf");

                            //checking to see if this docs eFolder bucket already exists
                            bool docExists = false;
                            foreach (TrackedDocument d in EncompassApplication.CurrentLoan.Log.TrackedDocuments)
                            {
                                if (d.Title == doc.EfolderName)
                                {
                                    docExists = true;
                                    break;
                                }
                            }
                            //if the eFolder bucket was missing we'll make a placeholder
                            if (!docExists)
                            {
                                if (EncompassApplication.CurrentLoan.Log.MilestoneEvents.LastCompletedEvent.MilestoneName.Equals("Completion"))
                                {
                                    EncompassApplication.CurrentLoan.Log.TrackedDocuments.Add(doc.EfolderName, "Completion");
                                }
                                else
                                {
                                    EncompassApplication.CurrentLoan.Log.TrackedDocuments.Add(doc.EfolderName, EncompassApplication.CurrentLoan.Log.MilestoneEvents.NextEvent.MilestoneName);
                                }
                                
                            }
                            //find the new placeholder and attach the new Word document to eFolder bucket
                            bool attached = false;
                            foreach (TrackedDocument d in EncompassApplication.CurrentLoan.Log.TrackedDocuments)
                            {
                                if (d.Title == doc.EfolderName)
                                {
                                    Attachment att = EncompassApplication.CurrentLoan.Attachments.Add(fullPathAsPDF);
                                    d.Attach(att);
                                    attached = true;
                                    break;
                                }
                            }
                            foreach(TrackedDocument d in EncompassApplication.CurrentLoan.Log.TrackedDocuments)
                            {
                                if (d.Title == "Notice of Incomplete Application")
                                {
                                    foreach(Attachment att in d.GetAttachments())
                                    {
                                        if (att.Title == "Untitled")
                                        {
                                            string borrower = Macro.GetField("4000") + " " + Macro.GetField("4002");
                                            if (!Macro.GetField("4004").Equals(""))
                                            {
                                                borrower = borrower + " and " + Macro.GetField("4004") + " " + Macro.GetField("4006");
                                            }
                                            att.Title = "Notice of Incomplete Application - " + borrower;
                                        }
                                    }
                                }
                            }
                            //delete the temp word doc in the SmartClientCache
                            System.IO.File.Delete(fullPathAsPDF);
                        }
                    }
                }
                catch (Exception)
                {

                }
            }
        }

        private void MergeDataIntoTempDoc(string fullPath)
        {
            //This method depends on the MS.Office.Interop.Word Reference
            //instantiate new Word Application
            Application app = new Application();

            //instantiate and bind temp Word doc into my MS Document object
            Document doc = app.Documents.Open(fullPath);

            //runs it all in the background
            app.Visible = false;

            foreach (Microsoft.Office.Interop.Word.Field field in doc.Fields)
            {
                if (field.Code.Text.Contains("MERGEFIELD"))
                {
                    int index_OpeningBracket = field.Code.Text.IndexOf("MERGEFIELD");

                    int index_Underscore = field.Code.Text.IndexOf("_");

                    string fieldID = field.Code.Text;
                    //fieldID = fieldID.Replace(fieldID.Substring(index_OpeningBracket, 10),"");

                    fieldID = fieldID.Replace(fieldID.Substring(index_Underscore, 1), "");
                    fieldID = fieldID.Replace("MERGEFIELD", "");
                    fieldID = fieldID.Replace(" ", "");
                    int index_M = fieldID.IndexOf("M");
                    fieldID = fieldID.Replace(fieldID.Substring(index_M, 1), "");
                    fieldID = fieldID.Replace("dot", ".");
                    try
                    {
                        string fieldValue = Macro.GetField(fieldID);
                        if (fieldValue == null || fieldValue.Equals(""))
                        {
                            field.Select();
                            doc.Application.Selection.TypeText(" ");
                        }
                        else
                        {
                            field.Select();
                            doc.Application.Selection.TypeText(fieldValue);
                        }


                    }
                    catch (Exception)
                    {
                        field.Select();
                        doc.Application.Selection.TypeText(" ");
                        

                    }

                }
            }


            //save the temp doc as a pdf with the merged data, close the doc, quit the app
            string fullPathAsPDF = fullPath.Replace(".docx", ".pdf");
            doc.ExportAsFixedFormat(fullPathAsPDF, WdExportFormat.wdExportFormatPDF);

            doc.Close();
            app.Quit();
        }

    }

    internal class AutoDoc
    {
        public string TriggerCode { get; set; }
        public byte[] Content { get; set; }
        public string EfolderName { get; set; }
        public string OriginalFileName { get; set; }
    }

    internal class ControlPair
    {
        public EllieMae.Encompass.Forms.CheckBox checkBoxControl { get; set; }
        public EllieMae.Encompass.Forms.TextBox textBoxControl { get; set; }
        public EllieMae.Encompass.Forms.DropdownBox dropdownBoxControl { get; set; }
    }
}
