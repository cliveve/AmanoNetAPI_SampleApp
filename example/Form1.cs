using ADODB;
using MSDASC;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//declare using AmanoNetSDK
using AmanoNetSDK;
using System.Text.RegularExpressions;

namespace AmanoNetAPI_SampleApp
{
    public partial class Form1 : Form
    {
        //Reference an instance of the API class
        public AmanoNetAPI amanoNetApi;

        int SiteDropdownDirty = 0;
        int DeptDropdownDirty = 0;

        public Form1()
        {
            InitializeComponent();
            this.MasterSiteComboBox.DropDown += MasterSiteComboBox_DropDown;
            this.MasterSiteComboBox.SelectedIndexChanged += new System.EventHandler(this.MasterSiteComboBox_SelectedIndexChanged);
            this.DepartmentComboBox.DropDown += DepartmentComboBox_DropDown;
            this.DepartmentComboBox.SelectedIndexChanged += DepartmentComboBox_SelectedIndexChanged;
            this.MSTSQ_TextBox.LostFocus += MSTSQ_TextBox_LostFocus;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //on load create new instance of AmanoNetAPI passing your app name and database string of AmanoNet database
            amanoNetApi = new AmanoNetAPI("My App", AmanoNetDbTextBox.Text);

            MasterTypeComboBox.Text = "Employee";
        }

        #region API_Usage

        //Here is where we do the update/insert for MASTER data
        private void InsertMasterDataButton_Click(object sender, EventArgs e)
        {
            string returnValue;

            //values need to be of the correct type, so from your drop downs you need to convert values to database types
            //from our example:
            //MasterCurrentComboBox has values 'True' or 'False', but you need to send an integer value 1 or 0. so...
            Int16 MST_Current;
            if (MasterCurrentComboBox.Text == "True")
                MST_Current = 1;
            else
                MST_Current = 0;

            //GenderComboBox has values 'Male' or 'Female', but you need to send a string value of M or F. so...
            string MST_Gender;
            if (GenderComboBox.Text == "Male")
                MST_Gender = "M";
            else
                MST_Gender = "F";

            //MasterTypeComboBox has values 'Employee' or 'Contractor/Visitor', but you need to send an integer value 0 or 1. so...
            Int16 MST_Type;
            if (MasterTypeComboBox.Text == "Employee")
                MST_Type = 0;
            else
                MST_Type = 1;

            //ClockingTypeComboBox has values 'Access Only' or 'Access and Time', but you need to send an integer value 1 or 2. so...
            int MT_NO;
            if (ClockingTypeComboBox.Text == "Access Only")
                MT_NO = 1;
            else
                MT_NO = 2;

            //Build up MASTER table values, assign values to the amanoNetApi.BuildMaster method from the API.
            //We do this to make sure our format is correct before doing Insert/Update...
            var Master_Record = amanoNetApi.BuildMaster(int.Parse(MSTSQ_TextBox.Text), TitleComboBox.Text, FirstNameTextBox.Text, MiddleNameTextBox.Text, LastNameTextBox.Text, "", IDNumberTextBox.Text, MST_Gender,
                                                        null, //MST_PIN is a nullable field, but you can populate this with an integer if you have the requirement
                                                        MST_Type, MST_Current, MasterSiteComboBox.Text,
                                                        null, //USRPRF_NUM is a nullable field, but you can populate this with an integer if you have the requirement
                                                        "", //because MST_CDATE is a string, it cannot be nullable from code, but you can pass a "" string if not using.
                                                        MT_NO //MT_NO can be a nullable field, but it is recommended to have the clocking type
                                                        );

            //if Master Type (MST_Type) = 'Employee' (0) use amanoNetApi.UpdateInsertMasterEmployee
            if (MST_Type == 0)
            {
                //Build up EMPLOYEE table values using amanoNetApi.BuildEmployee
                var Employee_Record = amanoNetApi.BuildEmployee(int.Parse(MSTSQ_TextBox.Text), EmployeeNumberTextBox.Text, EmployeeEmployerTextBox.Text, EmployeePositionTextBox.Text, Convert.ToInt32(DepartmentComboBox.Text), MasterSiteComboBox.Text);

                //Now do Insert/Update by calling amanoNetApi.UpdateInsertMasterEmployee and using variables 'Master_Record' and 'Employee_Record'
                returnValue = amanoNetApi.UpdateInsertMasterEmployee(Master_Record, Employee_Record);
            }
            //else if Master Type (MST_Type) = 'Contractor/Visitor' (1) use amanoNetApi.UpdateInsertMasterVisitor
            else
            {
                //Build up for VISITOR table values using amanoNetApi.BuildEmployee
                var Visitor_Record = amanoNetApi.BuildVisitor(int.Parse(MSTSQ_TextBox.Text), VisitorCompanyTextBox.Text, 0, Convert.ToInt32(DepartmentComboBox.Text),  MasterSiteComboBox.Text,
                                                                null //HOSTED_BY is a nullable field, you would need to know the MST_SQ of the host first to populate this.
                                                                );

                //Now do Insert/Update by calling amanoNetApi.UpdateInsertMasterVisitor and using variables 'Master_Record' and 'Visitor_Record'
                returnValue = amanoNetApi.UpdateInsertMasterVisitor(Master_Record, Visitor_Record);                                                                
            }

            //returnValue for Employee should be "MASTER='value > 0';EMPLOYEE='value > 0 and equal to MASTER value'".
            //if either values are 0 it means the Insert/Update failed for either of those tables. 
            //returnValue for Visitor should be "MASTER='value > 0';VISITOR='value > 0 and equal to MASTER value'". 
            //if either values are 0 it means the Insert/Update failed for either of those tables. 
            InsertUpdateMasterDataResponseTextBox.Text = returnValue;

        }

        //Here is an example of doing a search by Firstname/Lastname or Middlename
        private void GetSpecificMasterRecordButton_Click(object sender, EventArgs e)
        {
            GetSpecificMasterRecordResponseTextBox.Text = "";

            String GetSpecificMasterRecord = amanoNetApi.SearchForSpecificMasterRecord(SearchCriteriaTextBox.Text);

            if (GetSpecificMasterRecord.Length > 0)
            {
                for (int i = 0; i < GetSpecificMasterRecord.ParseJSON<MASTER>().Count(); i++)
                {
                    GetSpecificMasterRecordResponseTextBox.AppendText(
                        "Master Record " + i.ToString() + ":" + Environment.NewLine +
                        "       MST_SQ = " + GetSpecificMasterRecord.ParseJSON<MASTER>()[i].MST_SQ.ToString() + Environment.NewLine +
                        "       Title = " + GetSpecificMasterRecord.ParseJSON<MASTER>()[i].MST_Title.ToString() + Environment.NewLine +
                        "       First Name = " + GetSpecificMasterRecord.ParseJSON<MASTER>()[i].MST_FirstName.ToString() + Environment.NewLine +
                        "       Middle Name = " + GetSpecificMasterRecord.ParseJSON<MASTER>()[i].MST_MiddleName.ToString() + Environment.NewLine +
                        "       Last Name = " + GetSpecificMasterRecord.ParseJSON<MASTER>()[i].MST_LastName.ToString() + Environment.NewLine +
                        "       Suffix = " + GetSpecificMasterRecord.ParseJSON<MASTER>()[i].MST_Suffix.ToString() + Environment.NewLine +
                        "       ID Number = " + GetSpecificMasterRecord.ParseJSON<MASTER>()[i].MST_ID.ToString() + Environment.NewLine +
                        "       Gender = " + GetSpecificMasterRecord.ParseJSON<MASTER>()[i].MST_Gender.ToString() + Environment.NewLine +
                        "       Pin = " + GetSpecificMasterRecord.ParseJSON<MASTER>()[i].MST_PIN.ToString() + Environment.NewLine +
                        "       Master Type = " + GetSpecificMasterRecord.ParseJSON<MASTER>()[i].MST_Type.ToString() + Environment.NewLine +
                        "       Current = " + GetSpecificMasterRecord.ParseJSON<MASTER>()[i].MST_Current.ToString() + Environment.NewLine +
                        "       Site SLA No = " + GetSpecificMasterRecord.ParseJSON<MASTER>()[i].SITE_SLA.ToString() + Environment.NewLine +
                        "       Site Name = " + GetSpecificMasterRecord.ParseJSON<SITE>()[i].SITE_Name.ToString() + Environment.NewLine +
                        "       User Profile No = " + GetSpecificMasterRecord.ParseJSON<MASTER>()[i].USRPRF_NUM.ToString() + Environment.NewLine
                        );
                }
            }
        }

        //Here is an example of calling MASTER data associated with the database primary key (MST_SQ). 
        //Type a number in 'Database MST_SQ' field, click anywhere else and the data will populate
        private void MSTSQ_TextBox_LostFocus(object sender, EventArgs e)
        {
            String GetMasterRecord = amanoNetApi.GetMasterRecordUsingMSTSQ(MSTSQ_TextBox.Text);

            if (GetMasterRecord.Length > 0)
            {
                for (int i = 0; i < GetMasterRecord.ParseJSON<MASTER>().Count(); i++)
                {
                    TitleComboBox.Text = GetMasterRecord.ParseJSON<MASTER>()[i].MST_Title.ToString();
                    FirstNameTextBox.Text = GetMasterRecord.ParseJSON<MASTER>()[i].MST_FirstName.ToString();
                    MiddleNameTextBox.Text = GetMasterRecord.ParseJSON<MASTER>()[i].MST_MiddleName.ToString();
                    LastNameTextBox.Text = GetMasterRecord.ParseJSON<MASTER>()[i].MST_LastName.ToString();
                    IDNumberTextBox.Text = GetMasterRecord.ParseJSON<MASTER>()[i].MST_ID.ToString();
                    if (GetMasterRecord.ParseJSON<MASTER>()[i].MST_Gender.ToString() == "M")
                        GenderComboBox.SelectedIndex = 0;
                    else
                        GenderComboBox.SelectedIndex = 1;
                    if (GetMasterRecord.ParseJSON<MASTER>()[i].MST_Type.ToString() == "0")
                    {
                        MasterTypeComboBox.SelectedIndex = 0;
                        EmployeeNumberTextBox.Text = GetMasterRecord.ParseJSON<EMPLOYEE>()[i].EMP_EmployeeNo.ToString();
                        EmployeeEmployerTextBox.Text = GetMasterRecord.ParseJSON<EMPLOYEE>()[i].EMP_Employer.ToString();
                        EmployeePositionTextBox.Text = GetMasterRecord.ParseJSON<EMPLOYEE>()[i].EMP_Position.ToString();
                        DepartmentComboBox.Text = GetMasterRecord.ParseJSON<EMPLOYEE>()[i].DEPT_No.ToString();
                        DepartmentNameLabel.Text = GetMasterRecord.ParseJSON<DEPARTMENT>()[i].DEPT_Name.ToString();
                    }
                    else
                    {
                        MasterTypeComboBox.SelectedIndex = 1;
                        DepartmentComboBox.Text = GetMasterRecord.ParseJSON<VISITOR>()[i].DEPT_No.ToString();
                        DepartmentNameLabel.Text = GetMasterRecord.ParseJSON<DEPARTMENT>()[i].DEPT_Name.ToString();
                    }
                    if (GetMasterRecord.ParseJSON<MASTER>()[i].MST_Current.ToString() == "1")
                        MasterCurrentComboBox.SelectedIndex = 0;
                    else
                        MasterCurrentComboBox.SelectedIndex = 1;
                    MasterSiteComboBox.Text = GetMasterRecord.ParseJSON<MASTER>()[i].SITE_SLA.ToString();
                    SiteNameLabel.Text = GetMasterRecord.ParseJSON<SITE>()[i].SITE_Name.ToString();
                    if (GetMasterRecord.ParseJSON<MASTER>()[i].MT_NO.ToString() == "1")
                        ClockingTypeComboBox.SelectedIndex = 0;
                    else
                        ClockingTypeComboBox.SelectedIndex = 1;
                   
                }
            }
        }

        #endregion

        #region FormMethods

        //helper to set up database connection for testing
        private void PortalDBbutton_Click(object sender, EventArgs e)
        {
            try
            {
                // DataLink Properties Dialog
                DataLinksClass objDataLink = new DataLinksClass();
                objDataLink.hWnd = this.Handle.ToInt32();
                // Prompt For the Dialog
                Connection objConn = (Connection)objDataLink.PromptNew();
                // Open the connection
                objConn.Open(objConn.ConnectionString, null, null, (int)ConnectModeEnum.adModeUnknown);

                StringBuilder builder = new StringBuilder();
                string[] words = objConn.ConnectionString.Split(';');
                foreach (string word in words)
                {
                    if (word.Contains("Password"))
                    {
                        builder.Append(word).Append(";");
                    }
                    if (word.Contains("User ID"))
                    {
                        builder.Append(word).Append(";");
                    }
                    if (word.Contains("Initial Catalog"))
                    {
                        builder.Append(word).Append(";");
                    }
                    if (word.Contains("Data Source"))
                    {
                        builder.Append(word).Append(";");
                    }
                    if (word.Contains("Integrated Security"))
                    {
                        builder.Append(word).Append(";");
                    }
                }

                if (builder.Length > 1)
                    builder.Remove(builder.Length - 1, 1);

                AmanoNetDbTextBox.Text = builder + ";Asynchronous Processing=true;";
                objDataLink = null;
                objConn = null;

            }
            catch (Exception exObj)
            {
                MessageBox.Show(exObj.Message);
            }
        }

        //helper to parse values for dropdowns datasource
        private DataTable ConvertJSONToDataTable(string jsonString)
        {
            DataTable dt = new DataTable();
            //strip out bad characters
            string[] jsonParts = Regex.Split(jsonString.Replace("[", "").Replace("]", ""), "},{");

            //hold column names
            List<string> dtColumns = new List<string>();

            //get columns
            foreach (string jp in jsonParts)
            {
                //only loop thru once to get column names
                string[] propData = Regex.Split(jp.Replace("{", "").Replace("}", ""), ",");
                foreach (string rowData in propData)
                {
                    try
                    {
                        int idx = rowData.IndexOf(":");
                        string n = rowData.Substring(0, idx - 1);
                        string v = rowData.Substring(idx + 1);
                        if (!dtColumns.Contains(n))
                        {
                            dtColumns.Add(n.Replace("\"", ""));
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception(string.Format("Error Parsing Column Name : {0}", rowData));
                    }

                }
                break; // TODO: might not be correct. Was : Exit For
            }

            //build dt
            foreach (string c in dtColumns)
            {
                dt.Columns.Add(c);
            }
            //get table data
            foreach (string jp in jsonParts)
            {
                string[] propData = Regex.Split(jp.Replace("{", "").Replace("}", ""), ",");
                DataRow nr = dt.NewRow();
                foreach (string rowData in propData)
                {
                    try
                    {
                        int idx = rowData.IndexOf(":");
                        string n = rowData.Substring(0, idx - 1).Replace("\"", "");
                        string v = rowData.Substring(idx + 1).Replace("\"", "");
                        nr[n] = v;
                    }
                    catch (Exception ex)
                    {
                        continue;
                    }

                }
                dt.Rows.Add(nr);
            }
            return dt;
        }

        private void MasterSiteComboBox_DropDown(object sender, System.EventArgs e)
        {
            string Sites = MasterSiteComboBox.Text;

            if (SiteDropdownDirty == 0)
            {
                string SiteList = amanoNetApi.GetSites();
                DataTable dt = ConvertJSONToDataTable(SiteList);
                MasterSiteComboBox.DataSource = dt;
                MasterSiteComboBox.DisplayMember = "SITE_SLA";
                MasterSiteComboBox.ValueMember = "SITE_Name"; //<--need this to populate the labels

                //Set dropdown as dirty (=1)
                SiteDropdownDirty = 1;
                DeptDropdownDirty = 0;
                DepartmentComboBox.Text = "";
                DepartmentComboBox.DataSource = null;
            }
        }

        private void MasterSiteComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            SiteNameLabel.Text = MasterSiteComboBox.SelectedValue.ToString();
            DeptDropdownDirty = 0;
            DepartmentComboBox.Text = "";
            DepartmentComboBox.DataSource = null;
        }

        private void DepartmentComboBox_DropDown(object sender, System.EventArgs e)
        {
            string Dept = DepartmentComboBox.Text;

            if (DeptDropdownDirty == 0)
            {
                string DeptList = amanoNetApi.GetDepartments(MasterSiteComboBox.Text);
                DataTable dt = ConvertJSONToDataTable(DeptList);
                DepartmentComboBox.DataSource = dt;
                DepartmentComboBox.DisplayMember = "DEPT_No";
                DepartmentComboBox.ValueMember = "DEPT_Name";
                
                //Set dropdown as dirty (=1)
                DeptDropdownDirty = 1;
            }
        }

        private void DepartmentComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            DepartmentNameLabel.Text = DepartmentComboBox.SelectedValue.ToString();
        }

        private void MasterTypeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(MasterTypeComboBox.Text == "Employee")
            {
                VisitorGroupBox.Enabled = false;
                EmployeeGroupBox.Enabled = true;
            }
            else
            {
                VisitorGroupBox.Enabled = true;
                EmployeeGroupBox.Enabled = false;
            }
        }

        #endregion
    }
}
