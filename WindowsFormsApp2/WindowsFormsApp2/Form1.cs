using MetroFramework.Controls;
using MetroFramework.Forms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using WindowsFormsApp2;

namespace WindowsFormsApp2
{
    public partial class form1 : MetroForm
    {
        internal List<string> advisers = new List<string>(); //use List<string> to be able to add data receieved from DB. Arrays cannot dynmically increase in size due to contraints of arrays.
        internal List<int> emp_ID = new List<int>();
        internal List<string> storeName = new List<string>();
        internal List<string> regionName = new List<string>();
        internal List<int> storeID = new List<int>();
        internal List<double> phaseArray = new List<double>();
        internal List<double> deficitArray = new List<double>();
        internal List<string> deficitFullList = new List<string>();

        internal OleDbConnection conn = new OleDbConnection(); //initialise the db connection here as it is used serveral times later on. Do need to remember to close each time it's opened

        //All the buttons, tabs, labels, pictureboxes, message boxes and progress bars are declared here
        internal MetroMessageBox noDataBackBox = new MetroMessageBox();
        internal MetroMessageBox noStoreBox = new MetroMessageBox();
        internal MetroMessageBox noDataForwardBox = new MetroMessageBox();
        internal MetroMessageBox noRegionBox = new MetroMessageBox();
        internal MetroTabPage tab;
        internal MetroTabPage storeTabName;
        internal MetroTabPage phasesTab;
        internal MetroTabPage regionTabName;
        internal MetroButton noStoreokButton = new MetroButton();
        internal MetroButton noDataBackOkButton = new MetroButton();
        internal MetroButton noDataForwardOkButton = new MetroButton();
        internal MetroButton noRegionOkButton = new MetroButton();
        internal MetroLabel newPhaseLabel = new MetroLabel();
        internal MetroLabel newPhaseLabelTotal = new MetroLabel();
        internal MetroLabel paygPhaseLabel = new MetroLabel();
        internal MetroLabel paygPhaseLabelTotal = new MetroLabel();
        internal MetroLabel broadPhaseLabel = new MetroLabel();
        internal MetroLabel broadPhaseLabelTotal = new MetroLabel();
        internal MetroLabel revPhaseLabel = new MetroLabel();
        internal MetroLabel revPhaseLabelTotal = new MetroLabel();
        internal MetroLabel bigAlertsLabel = new MetroLabel();

        internal CircularProgressBar.CircularProgressBar newSalesProgressBar = new CircularProgressBar.CircularProgressBar();
        internal CircularProgressBar.CircularProgressBar paygSalesProgressBar = new CircularProgressBar.CircularProgressBar();
        internal CircularProgressBar.CircularProgressBar broadbandProgressBar = new CircularProgressBar.CircularProgressBar();
        internal CircularProgressBar.CircularProgressBar revenueProgressBar = new CircularProgressBar.CircularProgressBar();
        internal CircularProgressBar.CircularProgressBar accsProgressBar = new CircularProgressBar.CircularProgressBar();
        internal CircularProgressBar.CircularProgressBar newMTDBar = new CircularProgressBar.CircularProgressBar();
        internal CircularProgressBar.CircularProgressBar paygMTDBar = new CircularProgressBar.CircularProgressBar();
        internal CircularProgressBar.CircularProgressBar broadMTDBar = new CircularProgressBar.CircularProgressBar();
        internal CircularProgressBar.CircularProgressBar revMTDBar = new CircularProgressBar.CircularProgressBar();
        internal CircularProgressBar.CircularProgressBar accMTDBar = new CircularProgressBar.CircularProgressBar();

        internal MetroLabel currentMonthLabel = new MetroLabel();
        internal PictureBox backMonth = new PictureBox();
        internal PictureBox forwardMonth = new PictureBox();
        internal MetroButton acceptButton = new MetroButton();

        internal String storeCode;
        internal int currentAdviserIndex; //Store the current index of the advisor to match up to the tabcontrol index to find which tabpage is open
        internal String regionCode;
        internal int currentStoreIndex; //Store the current index of the store to match up to the tabcontrol index to find which tabpage is open

        //Declaring the sales and target variables here for code readability
        internal int newSales;
        internal int paygSales;
        internal int broadbandSales;
        internal int revenueSales;
        internal int accsSales;

        internal int newSalesTarget;
        internal int paygSalesTarget;
        internal int broadbandSalesTarget;
        internal int revenueSalesTarget;
        internal int accsSalesTarget;

        internal double newMTDTarget;
        internal double paygMTDTarget;
        internal double broadMTDTarget;
        internal double revMTDTarget;
        internal double accMTDTarget;

        internal double newDeficit;
        internal double paygDeficit;
        internal double broadDeficit;
        internal double revDeficit;
        internal double accDeficit;

        internal int deficitOne = 0;
        internal int deficitTwo = 0;
        internal int deficitThree = 0;

        //Used to detect whether the store or region view is open
        internal Boolean storeSelected = false;

        internal DateTime now = DateTime.Now;
        internal DateTime startDate_datetime;
        internal DateTime endDate_datetime;
        internal DateTime implementationDate = new DateTime(2017, 12, 1, 0, 0, 0); //Using an implementation date to stop the user going too far back as db calls are more expensive
        internal int dateCompare;
        internal double MTDCalc;
        internal double daysInMonth;

        public form1()
        {
            InitializeComponent();

            TabControl2.SelectedIndexChanged += TabControl2_SelectedIndexChanged;
            this.backMonth.Click += new System.EventHandler(this.backMonth_Click);
            this.forwardMonth.Click += new System.EventHandler(this.forwardMonth_Click);
            this.noStoreokButton.Click += new System.EventHandler(this.noStoreokButton_Click);
            this.noDataBackOkButton.Click += new System.EventHandler(this.noDataBackOkButton_Click);
            this.noDataForwardOkButton.Click += new System.EventHandler(this.noDataForwardOkButton_Click);
        }

        public void MTDGetter()
        {
            if (now == startDate_datetime && now.Month != DateTime.Today.Month) //As going back and forward through months makes now equal to the startdate, the mtd calc equals 0, just gets the targets for the full month
            {
                MTDCalc = Convert.ToInt32((DateTime.DaysInMonth(now.Year, now.Month)));
            }

            else if (now.Month == DateTime.Today.Month)
            {
                MTDCalc = Convert.ToInt32(DateTime.Today.Day);
            }

            else
            {
                MTDCalc = Convert.ToInt32(Math.Ceiling((now - startDate_datetime).TotalDays)); //Needs to be rounded up as there have been 19 days by 19/02/2018, but only 18 days between 01/02/2018 and 19/02/2018
            }            

            daysInMonth = DateTime.DaysInMonth(now.Year, now.Month); //getting the days in the currently selected month

            newMTDTarget = Convert.ToDouble(Math.Ceiling((newSalesTarget / daysInMonth) * MTDCalc)); //Rounding upwards for targets at user's request
            paygMTDTarget = Convert.ToDouble(Math.Ceiling((paygSalesTarget / daysInMonth) * MTDCalc));
            broadMTDTarget = Convert.ToDouble(Math.Ceiling((broadbandSalesTarget / daysInMonth) * MTDCalc));
            revMTDTarget = Convert.ToDouble(Math.Ceiling((revenueSalesTarget / daysInMonth) * MTDCalc));
            accMTDTarget = Convert.ToDouble(Math.Ceiling((accsSalesTarget / daysInMonth) * MTDCalc));
        }

        public void getData()
        {
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\dashboard.accdb";
            OleDbCommand getEmpName = new OleDbCommand("SELECT Emp_Name FROM Employees WHERE Emp_Store = @storeCode", conn); //using parameterized queries to prevent SQL injection attacks
            OleDbCommand getEmpNum = new OleDbCommand("SELECT EmployeeID FROM Employees WHERE Emp_Store = @storeCode", conn);
            OleDbCommand getStoreName = new OleDbCommand("SELECT store_Name FROM Stores WHERE StoreNumber = @storeCode", conn);

            OleDbDataReader getEmpNameReader;
            OleDbDataReader getEmpNumReader;
            OleDbDataReader getStoreNameReader;

            getEmpName.Parameters.AddWithValue("@storecode", storeCode); 
            getEmpNum.Parameters.AddWithValue("@storeCode", storeCode);
            getStoreName.Parameters.AddWithValue("@storeCode", storeCode);

            try
            {
                conn.Open(); //open the db connection

                getEmpNameReader = getEmpName.ExecuteReader(CommandBehavior.CloseConnection); //closing the reader connections here after use
                getEmpNumReader = getEmpNum.ExecuteReader(CommandBehavior.CloseConnection);
                getStoreNameReader = getStoreName.ExecuteReader(CommandBehavior.CloseConnection);

                if (getStoreNameReader.Read()) //Checks that the store name exists in the db
                {
                    while (getEmpNameReader.Read() && getEmpNumReader.Read()) //while statement to allow every item to be added into the lists
                    {
                        advisers.Add(Convert.ToString(getEmpNameReader[0])); //adds items to lists from the reader at index 0 as as only 1 thing is stored by the readers at index 0
                        emp_ID.Add(Convert.ToInt32(getEmpNumReader[0])); //Conversion to the correct data type is required as objects are returned by the readers
                        storeName.Add(Convert.ToString(getStoreNameReader[0]));

                        //makes everything invisible apart from the tabcontrol
                        storeCodeInput.Visible = false;
                        storePromptLabel.Visible = false;
                        regionCodeBox.Visible = false;
                        regionLoginButton.Visible = false;
                        regionLabel.Visible = false;
                        loginButton.Visible = false;
                        TabControl2.Visible = true;
                    }

                    conn.Close(); //close the connection before moving on as the connection cannot be used again of it is open
                    addStoreTab();
                }

                else
                {
                    conn.Close(); //Close the connection otherwise the connection remains open and the user cannot reattempt to login
                    noStorecodeBox();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public void addStoreTab()
        {
            storeTabName = new MetroTabPage(); //Initialises the tabpage

            storeTabName.Text = storeName[0]; //Grabs the name from the list at index 0. Is stored in a list for the regional view and code re-use
            storeTabName.Theme = MetroFramework.MetroThemeStyle.Dark; //Default style is light
            storeTabName.AutoScroll = true;

            TabControl2.Controls.Add(storeTabName); //Adds the tabpage to the tabcontrol
            generateTabs();
        }

        //Called after the store tab generation as the store will be the first tab followed by the advisers
        public void generateTabs()
        {
            for (int i = 0; i < advisers.Count; i++) //Only adds tabs for the length of the advisor list
            {
                tab = new MetroTabPage(); //Initialises the tabpage
                tab.Text = advisers[i]; //Names the tab after the adivsor at index[i] where i is always increasing
                tab.Theme = MetroFramework.MetroThemeStyle.Dark;

                TabControl2.Controls.Add(tab);
            }
            genStorePhaseTab();
        }

        //Phases tab is generated last as it will be after all the advisors
        public void genStorePhaseTab()
        {
            phasesTab = new MetroTabPage(); //Initialises the phasesTab tab

            phasesTab.Text = "Phases"; //This tab will alway be called phases, so doesn't need a list
            phasesTab.Theme = MetroFramework.MetroThemeStyle.Dark;

            TabControl2.Controls.Add(phasesTab);

            //Checks if the user is in the store or regional view
            if (TabControl2.SelectedTab == storeTabName)
            {
                getStoreTabData();
            }

            else if (TabControl2.SelectedTab == regionTabName)
            {
                getRegionTabData();
            }

            else //Just in case users are able to move faster than the PCs they are on, which aren't particularly fast, calls the advisor methods
            {
                getEmployeeIndex();
                getTabData();
            }
        }

        public void getEmployeeIndex()
        {
            //Checks if the user is in store or regional view
            if (storeSelected == true) 
            {
                if (TabControl2.SelectedIndex == 0) //Store is always at 0, so I know to call to 0 in this case
                {
                    getStoreTabData();
                }

                //Takes 1 off the tabcontrol index to account for the store tab
                else
                {
                    currentAdviserIndex = TabControl2.SelectedIndex - 1; //Advisor index is used to check the employee id to be sent to the db to get data and relies on the tabcontrol index
                }
            }

            else if (storeSelected == false)
            {
                if (TabControl2.SelectedIndex == 0)
                {
                    getRegionTabData();
                }
                //Takes 1 off the tabcontrol index to account for the store tab
                else
                {
                    currentStoreIndex = TabControl2.SelectedIndex; //Store index is used to check the employee id to be sent to the db to get data and relies on the tabcontrol index
                }
            }
        }

        public void getTabData()
        {
            getEmployeeIndex(); //Gets the employee index to use in the paramaterized queries

            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\dashboard.accdb"; //The connection needs the connection type defined as well as the path which is not an absolute path
            //Each query needs an individual command
            OleDbCommand getnewSales = new OleDbCommand("SELECT COUNT (sale_Type) FROM Sales WHERE Emp_ID = @Emp_ID AND sale_Type = 1 AND date_Sold BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getNewTarget = new OleDbCommand("SELECT new_Targ FROM Targets WHERE Emp_ID = @Emp_ID AND targ_Month BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getPaygSales = new OleDbCommand("SELECT COUNT (sale_Type) FROM Sales WHERE Emp_ID = @Emp_ID AND sale_Type = 2 AND date_Sold BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getPaygTarget = new OleDbCommand("SELECT payg_Targ from Targets where Emp_ID = @Emp_ID AND targ_Month BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getRevSales = new OleDbCommand("SELECT sale_Rev FROM Sales WHERE Emp_ID = @Emp_ID AND date_Sold BETWEEN @startDate AND endDate", conn);
            OleDbCommand getRevSalesTarget = new OleDbCommand("SELECT rev_Targ FROM Targets WHERE Emp_ID = @Emp_ID AND targ_Month BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getBroadSales = new OleDbCommand("SELECT COUNT (sale_Type) FROM Sales WHERE Emp_ID = @Emp_ID AND sale_Type = 3 AND date_Sold BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getBroadTarget = new OleDbCommand("SELECT broadband_Targ FROM Targets WHERE Emp_ID = @Emp_ID AND targ_Month BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getAccSales = new OleDbCommand("SELECT SUM (sale_Rev) FROM Sales WHERE Emp_ID = @Emp_ID AND sale_Type = 4 AND date_Sold BETWEEN @start_Date AND @end_Date", conn);
            OleDbCommand getAccTarget = new OleDbCommand("SELECT Acc_Targ FROM Targets WHERE Emp_ID = @Emp_ID AND targ_Month BETWEEN @start_Date AND @end_Date", conn);

            //Each reader is defined within the scope of the method as there are not used in any other method
            OleDbDataReader getnewSalesReader;
            OleDbDataReader getnewTargetsReader;
            OleDbDataReader getPaygSalesReader;
            OleDbDataReader getPaygTargetsReader;
            OleDbDataReader getRevSalesReader;
            OleDbDataReader getRevSalesTargetReader;
            OleDbDataReader getBroadSalesReader;
            OleDbDataReader getBroadTargetReader;
            OleDbDataReader getAccSalesReader;
            OleDbDataReader getAccTargetsReader;

            getnewSales.Parameters.AddWithValue("@Emp_ID", emp_ID[currentAdviserIndex]);
            getnewSales.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getnewSales.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getNewTarget.Parameters.AddWithValue("@Emp_ID", emp_ID[currentAdviserIndex]);
            getNewTarget.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getNewTarget.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getPaygSales.Parameters.AddWithValue("@Emp_ID", emp_ID[currentAdviserIndex]);
            getPaygSales.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getPaygSales.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getPaygTarget.Parameters.AddWithValue("@Emp_ID", emp_ID[currentAdviserIndex]);
            getPaygTarget.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getPaygTarget.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getRevSales.Parameters.AddWithValue("@Emp_ID", emp_ID[currentAdviserIndex]);
            getRevSales.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getRevSales.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getRevSalesTarget.Parameters.AddWithValue("@Emp_ID", emp_ID[currentAdviserIndex]);
            getRevSalesTarget.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getRevSalesTarget.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getBroadSales.Parameters.AddWithValue("@Emp_ID", emp_ID[currentAdviserIndex]);
            getBroadSales.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getBroadSales.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getBroadTarget.Parameters.AddWithValue("@Emp_ID", emp_ID[currentAdviserIndex]);
            getBroadTarget.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getBroadTarget.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getAccSales.Parameters.AddWithValue("@Emp_ID", emp_ID[currentAdviserIndex]);
            getAccSales.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getAccSales.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getAccTarget.Parameters.AddWithValue("@Emp_ID", emp_ID[currentAdviserIndex]);
            getAccTarget.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getAccTarget.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            try
            {
                conn.Open();

                getnewSalesReader = getnewSales.ExecuteReader(CommandBehavior.CloseConnection);                
                getnewTargetsReader = getNewTarget.ExecuteReader(CommandBehavior.CloseConnection);                
                getPaygSalesReader = getPaygSales.ExecuteReader(CommandBehavior.CloseConnection);
                getPaygTargetsReader = getPaygTarget.ExecuteReader(CommandBehavior.CloseConnection);
                getRevSalesReader = getRevSales.ExecuteReader(CommandBehavior.CloseConnection);
                getRevSalesTargetReader = getRevSalesTarget.ExecuteReader(CommandBehavior.CloseConnection);
                getBroadSalesReader = getBroadSales.ExecuteReader(CommandBehavior.CloseConnection);
                getBroadTargetReader = getBroadTarget.ExecuteReader(CommandBehavior.CloseConnection);
                getAccSalesReader = getAccSales.ExecuteReader(CommandBehavior.CloseConnection);
                getAccTargetsReader = getAccTarget.ExecuteReader(CommandBehavior.CloseConnection);

                //Makes every reader fetch the information as well as looping until the readers are finished
                while (getnewSalesReader.Read() && getnewTargetsReader.Read() && getPaygSalesReader.Read() && getPaygTargetsReader.Read() && getRevSalesReader.Read() && getRevSalesTargetReader.Read() && getBroadSalesReader.Read() && getBroadTargetReader.Read() && getAccSalesReader.Read() && getAccTargetsReader.Read())
                {
                    if(getAccSalesReader.GetValue(0) == DBNull.Value)
                    {
                        accsSales = 0;
                    }
                    else
                    {
                        accsSales = Convert.ToInt32(getAccSalesReader.GetValue(0));
                    }

                    if (getnewSalesReader.GetValue(0) == DBNull.Value)
                    {
                        newSales = 0;
                    }
                    else
                    {
                        newSales = Convert.ToInt32(getnewSalesReader.GetValue(0));
                    }

                    if (getPaygSalesReader.GetValue(0) == DBNull.Value)
                    {
                        paygSales = 0;
                    }
                    else
                    {
                        paygSales = Convert.ToInt32(getPaygSalesReader.GetValue(0));
                    }

                    if (getRevSalesReader.GetValue(0) == DBNull.Value)
                    {
                        revenueSales = 0;
                    }
                    else
                    {
                        revenueSales = Convert.ToInt32(getRevSalesReader.GetValue(0));
                    }

                    if (getBroadSalesReader.GetValue(0) == DBNull.Value)
                    {
                        broadbandSales = 0;
                    }
                    else
                    {
                        broadbandSales = Convert.ToInt32(getBroadSalesReader.GetValue(0));
                    }

                    //Conversion to int32 has to be done as readers return objects
                    newSalesTarget = Convert.ToInt32(getnewTargetsReader.GetValue(0));
                    paygSalesTarget = Convert.ToInt32(getPaygTargetsReader.GetValue(0));
                    revenueSalesTarget = Convert.ToInt32(getRevSalesTargetReader.GetValue(0));
                    broadbandSalesTarget = getBroadTargetReader.GetInt32(0);
                    accsSalesTarget = Convert.ToInt32(getAccTargetsReader.GetValue(0));

                    MTDGetter(); //Must be called after the targets are retrieved from db so MTD targets can be calculated

                    //Sets the text inside the progress bars
                    newSalesProgressBar.Text = newSales + "/" + newSalesTarget + "\r\nnew \r\nconnections";
                    paygSalesProgressBar.Text = paygSales + "/" + paygSalesTarget + "\r\nPAYG \r\nconnections";
                    revenueProgressBar.Text = "£" + revenueSales + "/£" + revenueSalesTarget + "\r\nNet Rev";
                    broadbandProgressBar.Text = broadbandSales + "/" + broadbandSalesTarget + "\r\n Broadband";
                    accsProgressBar.Text = accsSales + "/" + accsSalesTarget + "\r\nAccessory \r\nRev";

                    newMTDBar.Text = newSales + "/" + newMTDTarget + "\r\nnew \r\nconnections \r\nMTD";
                    paygMTDBar.Text = paygSales + "/" + paygMTDTarget + "\r\nPAYG \r\nconnections \r\nMTD";
                    revMTDBar.Text = "£" + revenueSales + "/£" + revMTDTarget + "\r\nNet Rev \r\nMTD";
                    broadMTDBar.Text = broadbandSales + "/" + broadMTDTarget + "\r\n Broadband \r\nMTD";
                    accMTDBar.Text = accsSales + "/" + accMTDTarget + "\r\n Accessory \r\nMTD";

                    currentMonthLabel.Text = now.ToString("MMMM yyyy");

                    //using selectedtab as only tabpages can be added to tabcontrols, so have to add the items to the tab, which means adding them to the selected tab as advisor tabs are fetched from the db
                    TabControl2.SelectedTab.Controls.Add(newSalesProgressBar);
                    TabControl2.SelectedTab.Controls.Add(paygSalesProgressBar);
                    TabControl2.SelectedTab.Controls.Add(currentMonthLabel);
                    TabControl2.SelectedTab.Controls.Add(backMonth);
                    TabControl2.SelectedTab.Controls.Add(forwardMonth);
                    TabControl2.SelectedTab.Controls.Add(revenueProgressBar);
                    TabControl2.SelectedTab.Controls.Add(broadbandProgressBar);
                    TabControl2.SelectedTab.Controls.Add(accsProgressBar);
                    TabControl2.SelectedTab.Controls.Add(newMTDBar);
                    TabControl2.SelectedTab.Controls.Add(paygMTDBar);
                    TabControl2.SelectedTab.Controls.Add(revMTDBar);
                    TabControl2.SelectedTab.Controls.Add(broadMTDBar);
                    TabControl2.SelectedTab.Controls.Add(accMTDBar);
                }
                    progressBarsIncrement();
                    itemPlace();
                }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                conn.Close(); //Connection needs to be closed to that it can be reused at another point
            }
        }

        public void getStoreTabData()
        {
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=|DataDirectory|\dashboard.accdb";
            OleDbCommand getStoreNewSales = new OleDbCommand("SELECT COUNT (sale_Type) FROM Sales WHERE store_ID = @store_ID AND sale_Type = 1 AND date_Sold BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getStoreNewTarget = new OleDbCommand("SELECT SUM (new_Targ) FROM Targets WHERE store_Number = @store_ID AND targ_Month BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getStorePaygSales = new OleDbCommand("SELECT COUNT (sale_Type) FROM Sales WHERE store_ID = @store_ID AND sale_Type = 2 AND date_Sold BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getStorePaygTarget = new OleDbCommand("SELECT SUM (payg_Targ) FROM Targets WHERE store_Number = @store_ID AND targ_Month BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getStoreRevSales = new OleDbCommand("SELECT SUM (sale_Rev) FROM Sales WHERE store_ID = @store_ID AND date_Sold BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getStoreRevSalesTarget = new OleDbCommand("SELECT SUM (rev_Targ) FROM Targets WHERE store_Number = @store_ID AND targ_Month BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getStoreBroadSales = new OleDbCommand("SELECT COUNT (sale_Type) FROM Sales WHERE store_ID = @store_ID AND sale_Type = 3 AND date_Sold BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getStoreBroadSalesTarget = new OleDbCommand("SELECT SUM (broadband_Targ) FROM Targets WHERE store_Number = @store_ID AND targ_Month BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getStoreAccSales = new OleDbCommand("SELECT SUM (sale_Rev) FROM Sales WHERE store_ID = @store_ID AND sale_Type = 4 AND date_Sold BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getStoreAccTarget = new OleDbCommand("SELECT SUM (Acc_Targ) FROM Targets WHERE store_Number = @store_ID AND targ_Month BETWEEN @startDate AND @endDate", conn);

            OleDbDataReader getStoreNewSalesReader;
            OleDbDataReader getStoreNewTargetsReader;
            OleDbDataReader getStorePaygSalesReader;
            OleDbDataReader getStorePaygTargetsReader;
            OleDbDataReader getStoreRevSalesReader;
            OleDbDataReader getStoreRevTargetReader;
            OleDbDataReader getStoreBroadSalesReader;
            OleDbDataReader getStoreBroadTargetsReader;
            OleDbDataReader getStoreAccSalesReader;
            OleDbDataReader getStoreAccTargetReader;

            getStoreNewSales.Parameters.AddWithValue("@store_ID", storeCode);
            getStoreNewSales.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getStoreNewSales.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getStoreNewTarget.Parameters.AddWithValue("@store_ID", storeCode);
            getStoreNewTarget.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getStoreNewTarget.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getStorePaygSales.Parameters.AddWithValue("@store_ID", storeCode);
            getStorePaygSales.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getStorePaygSales.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getStorePaygTarget.Parameters.AddWithValue("@store_ID", storeCode);
            getStorePaygTarget.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getStorePaygTarget.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getStoreRevSales.Parameters.AddWithValue("@store_ID", storeCode);
            getStoreRevSales.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getStoreRevSales.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getStoreRevSalesTarget.Parameters.AddWithValue("@store_ID", storeCode);
            getStoreRevSalesTarget.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getStoreRevSalesTarget.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getStoreBroadSales.Parameters.AddWithValue("@store_ID", storeCode);
            getStoreBroadSales.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getStoreBroadSales.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getStoreBroadSalesTarget.Parameters.AddWithValue("@store_ID", storeCode);
            getStoreBroadSalesTarget.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getStoreBroadSalesTarget.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getStoreAccSales.Parameters.AddWithValue("@store_ID", storeCode);
            getStoreAccSales.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getStoreAccSales.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getStoreAccTarget.Parameters.AddWithValue("@store_ID", storeCode);
            getStoreAccTarget.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getStoreAccTarget.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            try
            {
                conn.Open();

                getStoreNewSalesReader = getStoreNewSales.ExecuteReader(CommandBehavior.CloseConnection);
                getStoreNewTargetsReader = getStoreNewTarget.ExecuteReader(CommandBehavior.CloseConnection);
                getStorePaygSalesReader = getStorePaygSales.ExecuteReader(CommandBehavior.CloseConnection);
                getStorePaygTargetsReader = getStorePaygTarget.ExecuteReader(CommandBehavior.CloseConnection);
                getStoreRevSalesReader = getStoreRevSales.ExecuteReader(CommandBehavior.CloseConnection);
                getStoreRevTargetReader = getStoreRevSalesTarget.ExecuteReader(CommandBehavior.CloseConnection);
                getStoreBroadSalesReader = getStoreBroadSales.ExecuteReader(CommandBehavior.CloseConnection);
                getStoreBroadTargetsReader = getStoreBroadSalesTarget.ExecuteReader(CommandBehavior.CloseConnection);
                getStoreAccSalesReader = getStoreAccSales.ExecuteReader(CommandBehavior.CloseConnection);
                getStoreAccTargetReader = getStoreAccTarget.ExecuteReader(CommandBehavior.CloseConnection);

                while (getStoreNewSalesReader.Read() && getStoreNewTargetsReader.Read() && getStorePaygSalesReader.Read() && getStorePaygTargetsReader.Read() && getStoreRevSalesReader.Read() && getStoreRevTargetReader.Read() && getStoreBroadSalesReader.Read() && getStoreBroadTargetsReader.Read() && getStoreAccSalesReader.Read() && getStoreAccTargetReader.Read())
                {
                    //Conversion to int32 has to be done as readers return objects
                    newSalesTarget = Convert.ToInt32(getStoreNewTargetsReader.GetValue(0));
                    paygSalesTarget = Convert.ToInt32(getStorePaygTargetsReader.GetValue(0));
                    revenueSalesTarget = Convert.ToInt32(getStoreRevTargetReader.GetValue(0));
                    broadbandSalesTarget = Convert.ToInt32(getStoreBroadTargetsReader.GetValue(0));
                    accsSalesTarget = Convert.ToInt32(getStoreAccTargetReader.GetValue(0));

                    if (getStoreAccSalesReader.GetValue(0) == DBNull.Value)
                    {
                        accsSales = 0;
                        accDeficit = 100;
                    }
                    else
                    {
                        accsSales = Convert.ToInt32(getStoreAccSalesReader.GetValue(0));
                        accDeficit = (accsSales / Convert.ToDouble(accsSalesTarget)) * 100;
                    }

                    if (getStoreNewSalesReader.GetValue(0) == DBNull.Value)
                    {
                        newSales = 0;
                        newDeficit = 100;
                    }
                    else
                    {
                        newSales = Convert.ToInt32(getStoreNewSalesReader.GetValue(0));
                        newDeficit = (newSales / Convert.ToDouble(newSalesTarget)) * 100; //Gets the deficit as a percentage and stores it in a variable
                    }

                    if (getStorePaygSalesReader.GetValue(0) == DBNull.Value)
                    {
                        paygSales = 0;
                        paygDeficit = 100;
                    }
                    else
                    {
                        paygSales = Convert.ToInt32(getStorePaygSalesReader.GetValue(0));
                        paygDeficit = (paygSales / Convert.ToDouble(paygSalesTarget)) * 100;
                    }

                    if (getStoreRevSalesReader.GetValue(0) == DBNull.Value)
                    {
                        revenueSales = 0;
                        revDeficit = 100;
                    }
                    else
                    {
                        revenueSales = Convert.ToInt32(getStoreRevSalesReader.GetValue(0));
                        revDeficit = (revenueSales / Convert.ToDouble(revenueSalesTarget)) * 100;
                    }

                    if (getStoreBroadSalesReader.GetValue(0) == DBNull.Value)
                    {
                        broadbandSales = 0;
                        broadDeficit = 100;
                    }
                    else
                    {
                        broadbandSales = Convert.ToInt32(getStoreBroadSalesReader.GetValue(0));
                        broadDeficit = (broadbandSales / Convert.ToDouble(broadbandSalesTarget)) * 100;
                    }

                    MTDGetter();

                    for (int i = 0; i < 4; i++) //Adds the phases into an list that is used in the phases method
                    {
                        phaseArray.Add(Math.Ceiling(Convert.ToDouble(newSalesTarget) / Convert.ToDouble(DateTime.DaysInMonth(now.Year, now.Month)))); //Phase is rounded up at user request, to make getting phase easier
                        phaseArray.Add(Math.Ceiling(Convert.ToDouble(paygSalesTarget) / Convert.ToDouble(DateTime.DaysInMonth(now.Year, now.Month))));
                        phaseArray.Add(Math.Ceiling(Convert.ToDouble(broadbandSalesTarget) / Convert.ToDouble(DateTime.DaysInMonth(now.Year, now.Month))));
                        phaseArray.Add(Math.Ceiling(Convert.ToDouble(revenueSalesTarget) / Convert.ToDouble(DateTime.DaysInMonth(now.Year, now.Month))));
                        phaseArray.Add(Math.Ceiling(Convert.ToDouble(accsSalesTarget) / Convert.ToDouble(DateTime.DaysInMonth(now.Year, now.Month))));
                    }

                    if (TabControl2.SelectedTab == storeTabName) //checks that the selected tab is the store tab as the phases method calls it and don't want all the progress bars moving around
                    {
                        currentMonthLabel.Text = now.ToString("MMMM yyyy");

                        newSalesProgressBar.Text = newSales + "/" + newSalesTarget + "\r\nnew \r\nconnections";
                        paygSalesProgressBar.Text = paygSales + "/" + paygSalesTarget + "\r\nPAYG \r\nconnections";
                        revenueProgressBar.Text = revenueSales + "/" + revenueSalesTarget + "\r\nNet Rev";
                        broadbandProgressBar.Text = broadbandSales + "/" + broadbandSalesTarget + "\r\n Broadband";
                        accsProgressBar.Text = accsSales + "/" + accsSalesTarget + "\r\nAccessory \r\nRev";

                        newMTDBar.Text = newSales + "/" + newMTDTarget + "\r\nnew \r\nconnections \r\nMTD";
                        paygMTDBar.Text = paygSales + "/" + paygMTDTarget + "\r\nPAYG \r\nconnections \r\nMTD";
                        revMTDBar.Text = revenueSales + "/" + revMTDTarget + "\r\nNet Rev \r\nMTD";
                        broadMTDBar.Text = broadbandSales + "/" + broadMTDTarget + "\r\n Broadband \r\nMTD";
                        accMTDBar.Text = accsSales + "/" + accMTDTarget + "\r\n Accessory \r\nMTD";

                        TabControl2.SelectedTab.Controls.Add(currentMonthLabel);
                        TabControl2.SelectedTab.Controls.Add(backMonth);
                        TabControl2.SelectedTab.Controls.Add(forwardMonth);
                        TabControl2.SelectedTab.Controls.Add(newSalesProgressBar);
                        TabControl2.SelectedTab.Controls.Add(paygSalesProgressBar);
                        TabControl2.SelectedTab.Controls.Add(revenueProgressBar);
                        TabControl2.SelectedTab.Controls.Add(broadbandProgressBar);
                        TabControl2.SelectedTab.Controls.Add(accsProgressBar);
                        TabControl2.SelectedTab.Controls.Add(newMTDBar);
                        TabControl2.SelectedTab.Controls.Add(paygMTDBar);
                        TabControl2.SelectedTab.Controls.Add(revMTDBar);
                        TabControl2.SelectedTab.Controls.Add(broadMTDBar);
                        TabControl2.SelectedTab.Controls.Add(accMTDBar);

                        progressBarsIncrement();
                        itemPlace();
                    }

                }
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }

        public void getDeficitsData()
        {
            getPhaseData(); //Calls the informstionabout the phase so that is ready and not reliant on the deficit data being pulled in
            deficitArray.Clear(); //Clears the list to stop 
            getStoreTabData(); //Gets the data to be added into the array

            deficitArray.Add(newDeficit);
            deficitArray.Add(paygDeficit);
            deficitArray.Add(broadDeficit);
            deficitArray.Add(revDeficit);
            deficitArray.Add(accDeficit);

            deficitArray.Sort(); //Sorts the list in ascending order, worst deficit at the top

            currentMonthLabel.Text = now.ToString("MMMM yyyy");

            //Tests if each deficit is the first, second or thrid worst by looking at its position in the list
            if (deficitArray[0] == newDeficit)
            {
                newSalesProgressBar.Text = newSales + "/" + newSalesTarget + "\r\nnew \r\nconnections";
                TabControl2.SelectedTab.Controls.Add(currentMonthLabel);
                TabControl2.SelectedTab.Controls.Add(backMonth);
                TabControl2.SelectedTab.Controls.Add(forwardMonth);
                TabControl2.SelectedTab.Controls.Add(newSalesProgressBar);

                newSalesProgressBar.Location = new Point(150, 600);

                progressBarStyling();
                progressBarsIncrement();
            }

            else if (deficitArray[1] == newDeficit)
            {
                newSalesProgressBar.Text = newSales + "/" + newSalesTarget + "\r\nnew \r\nconnections";
                TabControl2.SelectedTab.Controls.Add(currentMonthLabel);
                TabControl2.SelectedTab.Controls.Add(backMonth);
                TabControl2.SelectedTab.Controls.Add(forwardMonth);
                TabControl2.SelectedTab.Controls.Add(newSalesProgressBar);

                newSalesProgressBar.Location = new Point(this.Width / 2 - 170, 600);

                progressBarStyling();
                progressBarsIncrement();
            }

            else if (deficitArray[2] == newDeficit)
            {
                newSalesProgressBar.Text = newSales + "/" + newSalesTarget + "\r\nnew \r\nconnections";
                TabControl2.SelectedTab.Controls.Add(currentMonthLabel);
                TabControl2.SelectedTab.Controls.Add(backMonth);
                TabControl2.SelectedTab.Controls.Add(forwardMonth);
                TabControl2.SelectedTab.Controls.Add(newSalesProgressBar);

                newSalesProgressBar.Location = new Point(1430, 600);

                progressBarStyling();
                progressBarsIncrement();
            }

            if (deficitArray[0] == paygDeficit)
            {
                paygSalesProgressBar.Text = paygSales + "/" + paygSalesTarget + "\r\nPAYG \r\nconnections";
                TabControl2.SelectedTab.Controls.Add(currentMonthLabel);
                TabControl2.SelectedTab.Controls.Add(backMonth);
                TabControl2.SelectedTab.Controls.Add(forwardMonth);
                TabControl2.SelectedTab.Controls.Add(paygSalesProgressBar);

                paygSalesProgressBar.Location = new Point(150, 600);

                progressBarStyling();
                progressBarsIncrement();
            }

            else if (deficitArray[1] == paygDeficit)
            {
                paygSalesProgressBar.Text = paygSales + "/" + paygSalesTarget + "\r\nPAYG \r\nconnections";
                TabControl2.SelectedTab.Controls.Add(currentMonthLabel);
                TabControl2.SelectedTab.Controls.Add(backMonth);
                TabControl2.SelectedTab.Controls.Add(forwardMonth);
                TabControl2.SelectedTab.Controls.Add(paygSalesProgressBar);

                paygSalesProgressBar.Location = new Point(this.Width / 2 - 170, 600);

                progressBarStyling();
                progressBarsIncrement();
            }

            else if (deficitArray[2] == paygDeficit)
            {
                paygSalesProgressBar.Text = paygSales + "/" + paygSalesTarget + "\r\nPAYG \r\nconnections";
                TabControl2.SelectedTab.Controls.Add(currentMonthLabel);
                TabControl2.SelectedTab.Controls.Add(backMonth);
                TabControl2.SelectedTab.Controls.Add(forwardMonth);
                TabControl2.SelectedTab.Controls.Add(paygSalesProgressBar);

                paygSalesProgressBar.Location = new Point(1430, 600);

                progressBarStyling();
                progressBarsIncrement();
            }

            if (deficitArray[0] == broadDeficit)
            {
                broadbandProgressBar.Text = broadbandSales + "/" + broadbandSalesTarget + "\r\n Broadband";
                TabControl2.SelectedTab.Controls.Add(currentMonthLabel);
                TabControl2.SelectedTab.Controls.Add(backMonth);
                TabControl2.SelectedTab.Controls.Add(forwardMonth);
                TabControl2.SelectedTab.Controls.Add(broadbandProgressBar);

                broadbandProgressBar.Location = new Point(150, 600);

                progressBarStyling();
                progressBarsIncrement();
            }

            else if (deficitArray[1] == broadDeficit)
            {
                broadbandProgressBar.Text = broadbandSales + "/" + broadbandSalesTarget + "\r\n Broadband";
                TabControl2.SelectedTab.Controls.Add(currentMonthLabel);
                TabControl2.SelectedTab.Controls.Add(backMonth);
                TabControl2.SelectedTab.Controls.Add(forwardMonth);
                TabControl2.SelectedTab.Controls.Add(broadbandProgressBar);

                broadbandProgressBar.Location = new Point(this.Width / 2 - 170, 600);

                progressBarStyling();
                progressBarsIncrement();
            }

            else if (deficitArray[2] == broadDeficit)
            {
                broadbandProgressBar.Text = broadbandSales + "/" + broadbandSalesTarget + "\r\n Broadband";
                TabControl2.SelectedTab.Controls.Add(currentMonthLabel);
                TabControl2.SelectedTab.Controls.Add(backMonth);
                TabControl2.SelectedTab.Controls.Add(forwardMonth);
                TabControl2.SelectedTab.Controls.Add(broadbandProgressBar);

                broadbandProgressBar.Location = new Point(1430, 600);

                progressBarStyling();
                progressBarsIncrement();
            }

            if (deficitArray[0] == revDeficit)
            {
                revenueProgressBar.Text = revenueSales + "/" + revenueSalesTarget + "\r\nNet Rev";
                TabControl2.SelectedTab.Controls.Add(currentMonthLabel);
                TabControl2.SelectedTab.Controls.Add(backMonth);
                TabControl2.SelectedTab.Controls.Add(forwardMonth);
                TabControl2.SelectedTab.Controls.Add(revenueProgressBar);

                revenueProgressBar.Location = new Point(150, 600);

                progressBarStyling();
                progressBarsIncrement();
            }

            else if (deficitArray[1] == revDeficit)
            {
                revenueProgressBar.Text = revenueSales + "/" + revenueSalesTarget + "\r\nNet Rev";
                TabControl2.SelectedTab.Controls.Add(currentMonthLabel);
                TabControl2.SelectedTab.Controls.Add(backMonth);
                TabControl2.SelectedTab.Controls.Add(forwardMonth);
                TabControl2.SelectedTab.Controls.Add(revenueProgressBar);

                revenueProgressBar.Location = new Point(this.Width / 2 - 170, 600);

                progressBarStyling();
                progressBarsIncrement();
            }

            else if (deficitArray[2] == revDeficit)
            {
                revenueProgressBar.Text = revenueSales + "/" + revenueSalesTarget + "\r\nNet Rev";
                TabControl2.SelectedTab.Controls.Add(currentMonthLabel);
                TabControl2.SelectedTab.Controls.Add(backMonth);
                TabControl2.SelectedTab.Controls.Add(forwardMonth);
                TabControl2.SelectedTab.Controls.Add(revenueProgressBar);

                revenueProgressBar.Location = new Point(1430, 600);

                progressBarStyling();
                progressBarsIncrement();
            }

            if (deficitArray[0] == accDeficit)
            {
                accsProgressBar.Text = accsSales + "/" + accsSalesTarget + "\r\n Accessory Rev";
                TabControl2.SelectedTab.Controls.Add(currentMonthLabel);
                TabControl2.SelectedTab.Controls.Add(backMonth);
                TabControl2.SelectedTab.Controls.Add(forwardMonth);
                TabControl2.SelectedTab.Controls.Add(accsProgressBar);

                accsProgressBar.Location = new Point(150, 600);

                progressBarStyling();
                progressBarsIncrement();
            }

            else if (deficitArray[1] == accDeficit)
            {
                accsProgressBar.Text = accsSales + "/" + accsSalesTarget + "\r\n Accessory Rev";
                TabControl2.SelectedTab.Controls.Add(currentMonthLabel);
                TabControl2.SelectedTab.Controls.Add(backMonth);
                TabControl2.SelectedTab.Controls.Add(forwardMonth);
                TabControl2.SelectedTab.Controls.Add(accsProgressBar);

                accsProgressBar.Location = new Point(150, 600);

                progressBarStyling();
                progressBarsIncrement();
            }

            else if (deficitArray[2] == accDeficit)
            {
                accsProgressBar.Text = accsSales + "/" + accsSalesTarget + "\r\n Accessory Rev";
                TabControl2.SelectedTab.Controls.Add(currentMonthLabel);
                TabControl2.SelectedTab.Controls.Add(backMonth);
                TabControl2.SelectedTab.Controls.Add(forwardMonth);
                TabControl2.SelectedTab.Controls.Add(accsProgressBar);

                accsProgressBar.Location = new Point(150, 600);

                progressBarStyling();
                progressBarsIncrement();
            }
        }

        public void getPhaseData()
        {
            //Sets the text for the labels
            newPhaseLabel.Text = "New Phase";
            paygPhaseLabel.Text = "PAYG Phase";
            broadPhaseLabel.Text = "Broadband Phase";
            revPhaseLabel.Text = "Revenue Phase";

            //Gets the phases for the data from the array, hence why the storetabdata is called in the deicits data method, it adds items into the phase list
            newPhaseLabelTotal.Text = phaseArray[0].ToString();
            paygPhaseLabelTotal.Text = phaseArray[1].ToString();
            broadPhaseLabelTotal.Text = phaseArray[2].ToString();
            revPhaseLabelTotal.Text = "£ " + phaseArray[3].ToString();

            bigAlertsLabel.Text = "Biggest Deficits";

            TabControl2.SelectedTab.Controls.Add(newPhaseLabel);
            TabControl2.SelectedTab.Controls.Add(paygPhaseLabel);
            TabControl2.SelectedTab.Controls.Add(broadPhaseLabel);
            TabControl2.SelectedTab.Controls.Add(revPhaseLabel);
            TabControl2.SelectedTab.Controls.Add(newPhaseLabelTotal);
            TabControl2.SelectedTab.Controls.Add(paygPhaseLabelTotal);
            TabControl2.SelectedTab.Controls.Add(broadPhaseLabelTotal);
            TabControl2.SelectedTab.Controls.Add(revPhaseLabelTotal);
            TabControl2.SelectedTab.Controls.Add(bigAlertsLabel);

            //Sets the location on the form
            newPhaseLabel.Location = new Point(200, 100);
            paygPhaseLabel.Location = new Point(600, 100);
            broadPhaseLabel.Location = new Point(1000, 100);
            revPhaseLabel.Location = new Point(1400, 100);
            newPhaseLabelTotal.Location = new Point(235, 200);
            paygPhaseLabelTotal.Location = new Point(635, 200);
            broadPhaseLabelTotal.Location = new Point(1060, 200);
            revPhaseLabelTotal.Location = new Point(1435, 200);
            bigAlertsLabel.Location = new Point(this.Width / 2 - 80, 400); //Labels center from the left hand side of the label, so if it is centered only the left hand side is centered, -80 brings the label to the center

            newPhaseLabel.Size = new Size(150, 50);
            paygPhaseLabel.Size = new Size(150, 50);
            broadPhaseLabel.Size = new Size(250, 50);
            revPhaseLabel.Size = new Size(150, 50);
            newPhaseLabelTotal.Size = new Size(150, 50);
            paygPhaseLabelTotal.Size = new Size(150, 50);
            broadPhaseLabelTotal.Size = new Size(150, 50);
            revPhaseLabelTotal.Size = new Size(150, 50);
            bigAlertsLabel.Size = new Size(150, 50);

            //As default theme is light, need to change these all to dark programmatically
            newPhaseLabel.Theme = MetroFramework.MetroThemeStyle.Dark;
            paygPhaseLabel.Theme = MetroFramework.MetroThemeStyle.Dark;
            broadPhaseLabel.Theme = MetroFramework.MetroThemeStyle.Dark;
            revPhaseLabel.Theme = MetroFramework.MetroThemeStyle.Dark;
            newPhaseLabelTotal.Theme = MetroFramework.MetroThemeStyle.Dark;
            paygPhaseLabelTotal.Theme = MetroFramework.MetroThemeStyle.Dark;
            broadPhaseLabelTotal.Theme = MetroFramework.MetroThemeStyle.Dark;
            revPhaseLabelTotal.Theme = MetroFramework.MetroThemeStyle.Dark;
            bigAlertsLabel.Theme = MetroFramework.MetroThemeStyle.Dark;

            //As user's ahve requested, label text is bolded and large
            newPhaseLabel.FontSize = MetroFramework.MetroLabelSize.Tall;
            newPhaseLabel.FontWeight = MetroFramework.MetroLabelWeight.Bold;
            paygPhaseLabel.FontSize = MetroFramework.MetroLabelSize.Tall;
            paygPhaseLabel.FontWeight = MetroFramework.MetroLabelWeight.Bold;
            broadPhaseLabel.FontSize = MetroFramework.MetroLabelSize.Tall;
            broadPhaseLabel.FontWeight = MetroFramework.MetroLabelWeight.Bold;
            revPhaseLabel.FontSize = MetroFramework.MetroLabelSize.Tall;
            revPhaseLabel.FontWeight = MetroFramework.MetroLabelWeight.Bold;
            newPhaseLabelTotal.FontSize = MetroFramework.MetroLabelSize.Tall;
            newPhaseLabelTotal.FontWeight = MetroFramework.MetroLabelWeight.Bold;
            paygPhaseLabelTotal.FontSize = MetroFramework.MetroLabelSize.Tall;
            paygPhaseLabelTotal.FontWeight = MetroFramework.MetroLabelWeight.Bold;
            broadPhaseLabelTotal.FontSize = MetroFramework.MetroLabelSize.Tall;
            broadPhaseLabelTotal.FontWeight = MetroFramework.MetroLabelWeight.Bold;
            revPhaseLabelTotal.FontSize = MetroFramework.MetroLabelSize.Tall;
            revPhaseLabelTotal.FontWeight = MetroFramework.MetroLabelWeight.Bold;
            bigAlertsLabel.FontSize = MetroFramework.MetroLabelSize.Tall;
            bigAlertsLabel.FontWeight = MetroFramework.MetroLabelWeight.Bold;
        }

        public void noStorecodeBox()
        {
            //If the storecode that the user entered was wrong, this sets up a message box to let the user know
            noStoreBox.Text = "Storecode not found";
            noStoreBox.Size = new Size(400, 150);
            noStoreBox.Theme = MetroFramework.MetroThemeStyle.Dark;
            noStoreBox.Style = MetroFramework.MetroColorStyle.Red; //As the dashboard is being designed for Vodafone at first, a red style is used to match their brand theme

            noStoreBox.Controls.Add(noStoreokButton); //Adding and OK button as if the user closes the box with through windows closing controls, the messagebox can't be used again and throws an error
            noStoreokButton.Text = "Ok";
            noStoreokButton.Location = new Point(150, 75);

            noStoreBox.Show(); //Shows the messagebox after its properties have been decided
        }

        public void noBackData()
        {
            //Sets up a messagebox if there is no data as far back as the user is requesting
            noDataBackBox.Text = "No data that far back.";
            noDataBackBox.Size = new Size(400, 150);
            noDataBackBox.Theme = MetroFramework.MetroThemeStyle.Dark;
            noDataBackBox.Style = MetroFramework.MetroColorStyle.Red;

            noDataBackBox.Controls.Add(noDataBackOkButton);
            noDataBackOkButton.Text = "Ok";
            noDataBackOkButton.Location = new Point(150, 75);

            noDataBackBox.Show(); //Makes the messagebox visible

            if (storeSelected == true)
            {
                if (TabControl2.SelectedIndex == 0) //Tests where the user, is. If they're in the store tab index[0], the get the store data, or in the phases tab, get the phases data, else they must be in an employee tab
                {
                    getStoreTabData();
                }

                else if (TabControl2.SelectedTab == phasesTab)
                {
                    getStoreTabData();
                    getDeficitsData();
                }

                else
                {
                    getTabData();
                }
            } //Tests if the user is in the store view or regional view

            else if (storeSelected == false)
            {
                getRegionTabData();
                getRegionStoreData();
            }
        }

        public void noForwardDate()
        {
            //Shows a messagebox if the user tries to get data from the future
            noDataForwardBox.Text = "No data for the future.";
            noDataForwardBox.Size = new Size(400, 150);
            noDataForwardBox.Theme = MetroFramework.MetroThemeStyle.Dark;
            noDataForwardBox.Style = MetroFramework.MetroColorStyle.Red;

            noDataForwardBox.Controls.Add(noDataForwardOkButton);
            noDataForwardOkButton.Text = "Ok";
            noDataForwardOkButton.Location = new Point(150, 75);

            noDataForwardBox.Show();

            if (storeSelected == true)
            {
                if (TabControl2.SelectedIndex == 0)
                {
                    getStoreTabData();
                }

                else if (TabControl2.SelectedTab == phasesTab)
                {
                    getStoreTabData();
                    getDeficitsData();
                }

                else
                {
                    getTabData();
                }
            }

            else if (storeSelected == false)
            {
                getRegionTabData();
                getRegionStoreData();
            }
        }

        public void dateForward()
        {
            //If the user can move the date forward, then this tests if they'rein the store or regional view before testing the tab they're in and calling the relevant data
            if (storeSelected == true)
            {
                if (TabControl2.SelectedIndex == 0)
                {
                    getStoreTabData();
                }

                else if (TabControl2.SelectedTab == phasesTab)
                {
                    getStoreTabData();
                    getDeficitsData();
                }

                else
                {
                    getTabData();
                }
            }

            else if (storeSelected == false)
            {
                getRegionTabData();
                getRegionStoreData();
            }
        }

        public void dateBack()
        {
            //If the user can move the date back, then this tests if they'rein the store or regional view before testing the tab they're in and calling the relevant data
            if (storeSelected == true)
            {
                if (TabControl2.SelectedIndex == 0)
                {
                    getStoreTabData();
                }

                else if (TabControl2.SelectedTab == phasesTab)
                {
                    getStoreTabData();
                    getDeficitsData();
                }

                else
                {
                    getTabData();
                }
            }

            else if (storeSelected == false)
            {
                getRegionTabData();
                getRegionStoreData();
            }
        }

        public void progressBarsIncrement()
        {
            //Sets the maximum size that the progressbars can get too, otherwise they have no way of increasing towards it, therefor not moving tha bar any further round
            newSalesProgressBar.Maximum = newSalesTarget;
            paygSalesProgressBar.Maximum = paygSalesTarget;
            revenueProgressBar.Maximum = revenueSalesTarget;
            broadbandProgressBar.Maximum = broadbandSalesTarget;
            accsProgressBar.Maximum = accsSalesTarget;
            //MTD Targets are stored as doubles to allow for rounding, but need to be converted to ints here
            newMTDBar.Maximum = Convert.ToInt32(newMTDTarget);
            paygMTDBar.Maximum = Convert.ToInt32(paygMTDTarget);
            revMTDBar.Maximum = Convert.ToInt32(revMTDTarget);
            broadMTDBar.Maximum = Convert.ToInt32(broadMTDTarget);
            accMTDBar.Maximum = Convert.ToInt32(accMTDTarget);

            //If the user has sold more than their MTD target then increase the max of the bar as the bar cannot go over its maximum otherwise IE 5/3 new connections is not allowed by default
            if (newSales > newSalesProgressBar.Maximum)
            {
                newSalesProgressBar.Maximum = newSales;
            }

            if (newSales > newMTDBar.Maximum)
            {
                newMTDBar.Maximum = newSales;
            }

            if (paygSales > paygSalesProgressBar.Maximum)
            {
                paygSalesProgressBar.Maximum = paygSales;
            }

            if (paygSales > paygMTDBar.Maximum)
            {
                paygMTDBar.Maximum = paygSales;
            }

            if (revenueSales > revenueProgressBar.Maximum)
            {
                revenueProgressBar.Maximum = revenueSales;
            }

            if (revenueSales > revMTDBar.Maximum)
            {
                revMTDBar.Maximum = revenueSales;
            }

            if (broadbandSales > broadbandProgressBar.Maximum)
            {
                broadbandProgressBar.Maximum = broadbandSales;
            }

            if (broadbandSales > broadMTDBar.Maximum)
            {
                broadMTDBar.Maximum = broadbandSales;
            }

            if (accsSales > accsProgressBar.Maximum)
            {
                accsProgressBar.Maximum = accsSales;
            }

            if (accsSales > accMTDBar.Maximum)
            {
                accMTDBar.Maximum = accsSales;
            }

            //Sets the value of the bar, this changes how far around the bar is
            newSalesProgressBar.Value = newSales;
            paygSalesProgressBar.Value = paygSales;
            revenueProgressBar.Value = revenueSales;
            broadbandProgressBar.Value = broadbandSales;
            accsProgressBar.Value = accsSales;
            newMTDBar.Value = newSales;
            paygMTDBar.Value = paygSales;
            revMTDBar.Value = revenueSales;
            broadMTDBar.Value = broadbandSales;
            accMTDBar.Value = accsSales;
        }

        public void itemPlace()
        {
            progressBarStyling(); //Calls the method to make sure the progress bars are coloured correctly before placing them on the tabcontrol

            //These are manually placed, so the dont' need any anchoring or docking
            newSalesProgressBar.Anchor = AnchorStyles.None;
            paygSalesProgressBar.Anchor = AnchorStyles.None;
            revenueProgressBar.Anchor = AnchorStyles.None;
            broadbandProgressBar.Anchor = AnchorStyles.None;
            accsProgressBar.Anchor = AnchorStyles.None;
            newMTDBar.Anchor = AnchorStyles.None;
            paygMTDBar.Anchor = AnchorStyles.None;
            broadMTDBar.Anchor = AnchorStyles.None;
            revMTDBar.Anchor = AnchorStyles.None;
            accMTDBar.Anchor = AnchorStyles.None;
            backMonth.Anchor = AnchorStyles.None;
            forwardMonth.Anchor = AnchorStyles.None;
            currentMonthLabel.Anchor = AnchorStyles.None;

            newSalesProgressBar.Dock = DockStyle.None;
            paygSalesProgressBar.Dock = DockStyle.None;
            revenueProgressBar.Dock = DockStyle.None;
            broadbandProgressBar.Dock = DockStyle.None;
            accsProgressBar.Dock = DockStyle.None;
            newMTDBar.Dock = DockStyle.None;
            paygMTDBar.Dock = DockStyle.None;
            revMTDBar.Dock = DockStyle.None;
            broadMTDBar.Dock = DockStyle.None;
            accMTDBar.Dock = DockStyle.None;
            backMonth.Dock = DockStyle.None;
            forwardMonth.Dock = DockStyle.None;

            //This sets where the progress bars are on the tabcontrol
            newSalesProgressBar.Location = new Point(55, 100);
            paygSalesProgressBar.Location = new Point(455, 100);
            revenueProgressBar.Location = new Point(855, 100);
            broadbandProgressBar.Location = new Point(1255, 100);
            accsProgressBar.Location = new Point(1605, 100);
            newMTDBar.Location = new Point(80, 500);
            paygMTDBar.Location = new Point(480, 500);
            revMTDBar.Location = new Point(880, 500);
            broadMTDBar.Location = new Point(1280, 500);
            accMTDBar.Location = new Point(1635, 500);

            newSalesProgressBar.AutoSize = true;
            paygSalesProgressBar.AutoSize = true;
            revenueProgressBar.AutoSize = true;

            //Sets the size of the label as autosize
            currentMonthLabel.Size = new Size(150, 50);

            currentMonthLabel.Location = new Point(927, 0); //Zero on the Y axis as tabpages work with (0,0) as top left
            currentMonthLabel.Theme = MetroFramework.MetroThemeStyle.Dark;
            currentMonthLabel.FontSize = MetroFramework.MetroLabelSize.Tall;
            currentMonthLabel.FontWeight = MetroFramework.MetroLabelWeight.Bold;

            backMonth.Image = Properties.Resources.backMonth; //Sets the picture in the picture box to an image stored in the project resources
            forwardMonth.Image = Properties.Resources.ForwardMonth;

            backMonth.BackColor = ColorTranslator.FromHtml("#111111"); //Changes the background of the picture box to the same colour as the application so only the arros are visible
            backMonth.Location = new Point(841, 0);

            forwardMonth.BackColor = ColorTranslator.FromHtml("111111");
            forwardMonth.Location = new Point(1100, 0);

            noDataBackBox.Theme = MetroFramework.MetroThemeStyle.Dark;
            noDataBackBox.Style = MetroFramework.MetroColorStyle.Red;
            noDataBackBox.Size = new Size(400, 100);
        }

        //Does everything that the store view does, but up a level, so stores = employees and the top level controlling the stores in the region
        public void getRegionTabs()
        {
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\dashboard.accdb";
            OleDbCommand getRegionName = new OleDbCommand("SELECT region_Name FROM Stores WHERE region_Code = @region_Code", conn);
            OleDbCommand getStoreID = new OleDbCommand("SELECT storeNumber FROM Stores WHERE region_Code = @region_Code", conn);
            OleDbCommand getStoreName = new OleDbCommand("SELECT store_Name FROM Stores WHERE region_Code = @region_Code", conn);

            OleDbDataReader getRegionNameReader;
            OleDbDataReader getRegionStoreIDReader;
            OleDbDataReader getRegionStoreNameReader;

            getRegionName.Parameters.AddWithValue("@region_Code", regionCode);
            getStoreID.Parameters.AddWithValue("@region_Code", regionCode);
            getStoreName.Parameters.AddWithValue("@region_Code", regionCode);

            try
            {
                conn.Open();

                getRegionNameReader = getRegionName.ExecuteReader(CommandBehavior.CloseConnection);
                getRegionStoreIDReader = getStoreID.ExecuteReader(CommandBehavior.CloseConnection);
                getRegionStoreNameReader = getStoreName.ExecuteReader(CommandBehavior.CloseConnection);

                if (getRegionNameReader.Read())
                {
                    while (getRegionStoreNameReader.Read() && getRegionStoreIDReader.Read())
                    {
                        regionName.Add(Convert.ToString(getRegionNameReader[0]));
                        storeID.Add(Convert.ToInt32(getRegionStoreIDReader[0]));
                        storeName.Add(Convert.ToString(getRegionStoreNameReader[0]));

                        storeCodeInput.Visible = false;
                        storePromptLabel.Visible = false;
                        regionCodeBox.Visible = false;
                        regionLoginButton.Visible = false;
                        regionLabel.Visible = false;
                        loginButton.Visible = false;
                        TabControl2.Visible = true;
                    }

                    conn.Close();
                    addRegionTab();
                }

                else
                {
                    conn.Close();
                    noRegionCodeBox();
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            finally
            {
                conn.Close();
            }
        }

        public void noRegionCodeBox()
        {
            noRegionBox.Text = "Region code not found";
            noRegionBox.Size = new Size(400, 150);
            noRegionBox.Theme = MetroFramework.MetroThemeStyle.Dark;
            noRegionBox.Style = MetroFramework.MetroColorStyle.Red;

            noRegionBox.Controls.Add(noRegionOkButton);
            noRegionOkButton.Text = "Ok";
            noRegionOkButton.Location = new Point(150, 75);

            noRegionBox.Show();
        }

        public void addRegionTab()
        {
            regionTabName = new MetroTabPage();

            regionTabName.Text = regionName[0];
            regionTabName.Theme = MetroFramework.MetroThemeStyle.Dark;

            TabControl2.Controls.Add(regionTabName);
            generateStoreRegionTabs();
        }

        public void generateStoreRegionTabs()
        {
            for (int i = 0; i < storeName.Count; i++)
            {
                storeTabName = new MetroTabPage();

                storeTabName.Text = storeName[i];
                storeTabName.Theme = MetroFramework.MetroThemeStyle.Dark;

                TabControl2.Controls.Add(storeTabName);
            }
            genStorePhaseTab();
        }

        public void getStoreIndex()
        {
            if (storeSelected == true)
            {
                if (TabControl2.SelectedIndex == 0)
                {
                    getStoreTabData();
                }

                else
                {
                    currentAdviserIndex = TabControl2.SelectedIndex - 1;
                }
            }

            else if (storeSelected == false)
            {
                if (TabControl2.SelectedIndex == 0)
                {
                    getRegionTabData();
                }

                else
                {
                    currentStoreIndex = TabControl2.SelectedIndex - 1;
                }
            }
        }

        public void getRegionTabData()
        {
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=|DataDirectory|\dashboard.accdb";
            OleDbCommand getRegionNewSales = new OleDbCommand("SELECT COUNT (sale_Type) FROM Sales WHERE region_Code = @region_Code AND sale_Type = 1 AND date_Sold BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getRegionNewTargets = new OleDbCommand("SELECT SUM (new_Targ) FROM Targets WHERE region_Code = @region_Code AND targ_Month BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getPaygSales = new OleDbCommand("SELECT COUNT (sale_Type) FROM Sales WHERE region_Code = @region_Code AND sale_Type = 2 AND date_Sold BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getPaygTarget = new OleDbCommand("SELECT SUM (payg_Targ) from Targets where region_Code = @region_Code AND targ_Month BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getRevSales = new OleDbCommand("SELECT SUM (sale_Rev) FROM Sales WHERE region_Code = @region_Code AND date_Sold BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getRevTarget = new OleDbCommand("SELECT SUM (rev_Targ) from Targets where region_Code = @region_Code AND targ_Month BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getBroadSales = new OleDbCommand("SELECT COUNT (sale_Type) FROM Sales WHERE region_Code = @region_Code AND sale_Type = 3 AND date_Sold BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getBroadTarget = new OleDbCommand("SELECT SUM (broadband_Targ) FROM Targets WHERE region_Code = @region_Code AND targ_Month BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getAccSales = new OleDbCommand("SELECT SUM (sale_Rev) FROM Sales WHERE region_Code = @region_Code AND sale_Type = 4 AND date_Sold BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getAccTarget = new OleDbCommand("SELECT SUM (Acc_Targ) FROM Targets WHERE region_Code = @region_Code AND targ_Month BETWEEN @startDate AND @endDate", conn);

            OleDbDataReader getRegionNewSalesReader;
            OleDbDataReader getRegionNewTargetsReader;
            OleDbDataReader getRegionPaygSalesReader;
            OleDbDataReader getRegionPaygTargetsReader;
            OleDbDataReader getRegionRevSalesReader;
            OleDbDataReader getRegionRevTargetReader;
            OleDbDataReader getRegionBroadSalesReader;
            OleDbDataReader getRegionBroadTargReader;
            OleDbDataReader getAccSalesReader;
            OleDbDataReader getAccTargetReader;

            getRegionNewSales.Parameters.AddWithValue("@region_Code", regionCode);
            getRegionNewSales.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getRegionNewSales.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getRegionNewTargets.Parameters.AddWithValue("@region_Code", regionCode);
            getRegionNewTargets.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getRegionNewTargets.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getPaygSales.Parameters.AddWithValue("@region_Code", regionCode);
            getPaygSales.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getPaygSales.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getPaygTarget.Parameters.AddWithValue("@region_Code", regionCode);
            getPaygTarget.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getPaygTarget.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getRevSales.Parameters.AddWithValue("@region_Code", regionCode);
            getRevSales.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getRevSales.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getRevTarget.Parameters.AddWithValue("@region_Code", regionCode);
            getRevTarget.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getRevTarget.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getBroadSales.Parameters.AddWithValue("@region_Code", regionCode);
            getBroadSales.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getBroadSales.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getBroadTarget.Parameters.AddWithValue("@region_Code", regionCode);
            getBroadTarget.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getBroadTarget.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getAccSales.Parameters.AddWithValue("@store_ID", storeCode);
            getAccSales.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getAccSales.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getAccTarget.Parameters.AddWithValue("@store_ID", storeCode);
            getAccTarget.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getAccTarget.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            try
            {
                conn.Open();

                getRegionNewSalesReader = getRegionNewSales.ExecuteReader(CommandBehavior.CloseConnection);
                getRegionNewTargetsReader = getRegionNewTargets.ExecuteReader(CommandBehavior.CloseConnection);
                getRegionPaygSalesReader = getPaygSales.ExecuteReader(CommandBehavior.CloseConnection);
                getRegionPaygTargetsReader = getPaygTarget.ExecuteReader(CommandBehavior.CloseConnection);
                getRegionRevSalesReader = getRevSales.ExecuteReader(CommandBehavior.CloseConnection);
                getRegionRevTargetReader = getRevTarget.ExecuteReader(CommandBehavior.CloseConnection);
                getRegionBroadSalesReader = getBroadSales.ExecuteReader(CommandBehavior.CloseConnection);
                getRegionBroadTargReader = getBroadTarget.ExecuteReader(CommandBehavior.CloseConnection);
                getAccSalesReader = getAccSales.ExecuteReader(CommandBehavior.CloseConnection);
                getAccTargetReader = getAccTarget.ExecuteReader(CommandBehavior.CloseConnection);

                while (getRegionNewSalesReader.Read() && getRegionNewTargetsReader.Read() && getRegionPaygSalesReader.Read() && getRegionPaygTargetsReader.Read() && getRegionRevSalesReader.Read() && getRegionRevTargetReader.Read() && getRegionBroadSalesReader.Read() && getRegionBroadTargReader.Read() && getAccSalesReader.Read() && getAccTargetReader.Read())
                {
                    if (getAccSalesReader.GetValue(0) == DBNull.Value)
                    {
                        accsSales = 0;
                    }
                    else
                    {
                        accsSales = Convert.ToInt32(getAccSalesReader.GetValue(0));
                    }

                    if (getRegionNewSalesReader.GetValue(0) == DBNull.Value)
                    {
                        newSales = 0;
                    }
                    else
                    {
                        newSales = Convert.ToInt32(getRegionNewSalesReader.GetValue(0));
                    }

                    if (getRegionPaygSalesReader.GetValue(0) == DBNull.Value)
                    {
                        paygSales = 0;
                    }
                    else
                    {
                        paygSales = Convert.ToInt32(getRegionPaygSalesReader.GetValue(0));
                    }

                    if (getRegionRevSalesReader.GetValue(0) == DBNull.Value)
                    {
                        revenueSales = 0;
                    }
                    else
                    {
                        revenueSales = Convert.ToInt32(getRegionRevSalesReader.GetValue(0));
                    }

                    if (getRegionBroadSalesReader.GetValue(0) == DBNull.Value)
                    {
                        broadbandSales = 0;
                    }
                    else
                    {
                        broadbandSales = Convert.ToInt32(getRegionBroadSalesReader.GetValue(0));
                    }



                    //Conversion to int32 has to be done as readers return objects
                    newSalesTarget = Convert.ToInt32(getRegionNewTargetsReader.GetValue(0));
                    paygSalesTarget = Convert.ToInt32(getRegionPaygTargetsReader.GetValue(0));
                    revenueSalesTarget = Convert.ToInt32(getRegionRevTargetReader.GetValue(0));
                    broadbandSalesTarget = getRegionBroadTargReader.GetInt32(0);
                    accsSalesTarget = Convert.ToInt32(getAccTargetReader.GetValue(0));

                    MTDGetter();

                    for (int i = 0; i < 4; i++)
                    {
                        phaseArray.Add(Math.Ceiling(Convert.ToDouble(newSalesTarget) / Convert.ToDouble(DateTime.DaysInMonth(now.Year, now.Month))));
                        phaseArray.Add(Math.Ceiling(Convert.ToDouble(paygSalesTarget) / Convert.ToDouble(DateTime.DaysInMonth(now.Year, now.Month))));
                        phaseArray.Add(Math.Ceiling(Convert.ToDouble(broadbandSalesTarget) / Convert.ToDouble(DateTime.DaysInMonth(now.Year, now.Month))));
                        phaseArray.Add(Math.Ceiling(Convert.ToDouble(revenueSalesTarget) / Convert.ToDouble(DateTime.DaysInMonth(now.Year, now.Month))));
                        phaseArray.Add(Math.Ceiling(Convert.ToDouble(accsSalesTarget) / Convert.ToDouble(DateTime.DaysInMonth(now.Year, now.Month))));
                    }

                    newDeficit = (newSales / Convert.ToDouble(newSalesTarget)) * 100;
                    paygDeficit = (paygSales / Convert.ToDouble(paygSalesTarget)) * 100;
                    broadDeficit = (broadbandSales / Convert.ToDouble(broadbandSalesTarget)) * 100;
                    revDeficit = (revenueSales / Convert.ToDouble(revenueSalesTarget)) * 100;
                    accDeficit = (accsSales / Convert.ToInt32(accsSalesTarget)) * 100;

                    if (TabControl2.SelectedTab == regionTabName)
                    {
                        currentMonthLabel.Text = now.ToString("MMMM yyyy");

                        newSalesProgressBar.Text = newSales + "/" + newSalesTarget + "\r\nnew \r\nconnections";
                        paygSalesProgressBar.Text = paygSales + "/" + paygSalesTarget + "\r\nPAYG \r\nconnections";
                        revenueProgressBar.Text = revenueSales + "/" + revenueSalesTarget + "\r\nNet Rev";
                        broadbandProgressBar.Text = broadbandSales + "/" + broadbandSalesTarget + "\r\n Broadband";
                        accsProgressBar.Text = accsSales + "/" + accsSalesTarget + "\r\nAccessory \r\nRev";

                        newMTDBar.Text = newSales + "/" + newMTDTarget + "\r\nnew \r\nconnections \r\nMTD";
                        paygMTDBar.Text = paygSales + "/" + paygMTDTarget + "\r\nPAYG \r\nconnections \r\nMTD";
                        revMTDBar.Text = revenueSales + "/" + revMTDTarget + "\r\nNet Rev \r\nMTD";
                        broadMTDBar.Text = broadbandSales + "/" + broadMTDTarget + "\r\n Broadband \r\nMTD";
                        accMTDBar.Text = accsSales + "/" + accMTDTarget + "\r\n Accessory \r\nMTD";

                        TabControl2.SelectedTab.Controls.Add(newSalesProgressBar);
                        TabControl2.SelectedTab.Controls.Add(paygSalesProgressBar);
                        TabControl2.SelectedTab.Controls.Add(revenueProgressBar);
                        TabControl2.SelectedTab.Controls.Add(broadbandProgressBar);
                        TabControl2.SelectedTab.Controls.Add(accsProgressBar);
                        TabControl2.SelectedTab.Controls.Add(currentMonthLabel);
                        TabControl2.SelectedTab.Controls.Add(backMonth);
                        TabControl2.SelectedTab.Controls.Add(forwardMonth);
                        TabControl2.SelectedTab.Controls.Add(newMTDBar);
                        TabControl2.SelectedTab.Controls.Add(paygMTDBar);
                        TabControl2.SelectedTab.Controls.Add(revMTDBar);
                        TabControl2.SelectedTab.Controls.Add(broadMTDBar);
                        TabControl2.SelectedTab.Controls.Add(accMTDBar);

                        progressBarsIncrement();
                        itemPlace();
                    }
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            finally
            {
                conn.Close();
            }
        }

        public void getRegionStoreData()
        {
            getStoreIndex();

            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=|DataDirectory|\dashboard.accdb";
            OleDbCommand getStoreNewSales = new OleDbCommand("SELECT COUNT (sale_Type) FROM Sales WHERE store_ID = @store_ID AND sale_Type = 1 AND date_Sold BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getStoreNewTarget = new OleDbCommand("SELECT SUM (new_Targ) FROM Targets WHERE store_Number = @store_ID AND targ_Month BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getStorePaygSales = new OleDbCommand("SELECT COUNT (sale_Type) FROM Sales WHERE store_ID = @store_ID AND sale_Type = 2 AND date_Sold BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getStorePaygTarget = new OleDbCommand("SELECT SUM (payg_Targ) FROM Targets WHERE store_Number = @store_ID AND targ_Month BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getStoreRevSales = new OleDbCommand("SELECT SUM (sale_Rev) FROM Sales WHERE store_ID = @store_ID AND date_Sold BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getStoreRevSalesTarget = new OleDbCommand("SELECT SUM (rev_Targ) FROM Targets WHERE store_Number = @store_ID AND targ_Month BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getStoreBroadSales = new OleDbCommand("SELECT COUNT (sale_Type) FROM Sales WHERE store_ID = @store_ID AND sale_Type = 3 AND date_Sold BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getStoreBroadSalesTarget = new OleDbCommand("SELECT SUM (broadband_Targ) FROM Targets WHERE store_Number = @store_ID AND targ_Month BETWEEN @startDate AND @endDate", conn);
            OleDbCommand getStoreAccSales = new OleDbCommand("SELECT SUM (sale_Rev) FROM Sales WHERE store_Number = @store_ID AND sale_Type = 4 AND date_Sold BETWEEN @start_Date AND @end_Date", conn);
            OleDbCommand getStoreAccTarget = new OleDbCommand("SELECT Acc_Targ FROM Targets WHERE store_Number = @store_ID AND targ_Month BETWEEN @start_Date AND @end_Date", conn);

            OleDbDataReader getStoreNewSalesReader;
            OleDbDataReader getStoreNewTargetsReader;
            OleDbDataReader getStorePaygSalesReader;
            OleDbDataReader getStorePaygTargetsReader;
            OleDbDataReader getStoreRevSalesReader;
            OleDbDataReader getStoreRevTargetReader;
            OleDbDataReader getStoreBroadSalesReader;
            OleDbDataReader getStoreBroadTargetsReader;
            OleDbDataReader getStoreAccSalesReader;
            OleDbDataReader getStoreAccTargetsReader;

            getStoreNewSales.Parameters.AddWithValue("@store_ID", storeID[currentStoreIndex]);
            getStoreNewSales.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getStoreNewSales.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getStoreNewTarget.Parameters.AddWithValue("@store_ID", storeID[currentStoreIndex]);
            getStoreNewTarget.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getStoreNewTarget.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getStorePaygSales.Parameters.AddWithValue("@store_ID", storeID[currentStoreIndex]);
            getStorePaygSales.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getStorePaygSales.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getStorePaygTarget.Parameters.AddWithValue("@store_ID", storeID[currentStoreIndex]);
            getStorePaygTarget.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getStorePaygTarget.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getStoreRevSales.Parameters.AddWithValue("@store_ID", storeID[currentStoreIndex]);
            getStoreRevSales.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getStoreRevSales.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getStoreRevSalesTarget.Parameters.AddWithValue("@store_ID", storeID[currentStoreIndex]);
            getStoreRevSalesTarget.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getStoreRevSalesTarget.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getStoreBroadSales.Parameters.AddWithValue("@store_ID", storeID[currentStoreIndex]);
            getStoreBroadSales.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getStoreBroadSales.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getStoreBroadSalesTarget.Parameters.AddWithValue("@store_ID", storeID[currentStoreIndex]);
            getStoreBroadSalesTarget.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getStoreBroadSalesTarget.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getStoreAccSales.Parameters.AddWithValue("@store_ID", storeID[currentStoreIndex]);
            getStoreAccSales.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getStoreAccSales.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            getStoreAccTarget.Parameters.AddWithValue("@store_ID", storeID[currentStoreIndex]);
            getStoreAccTarget.Parameters.AddWithValue("@startDate", startDate_datetime.ToShortDateString());
            getStoreAccTarget.Parameters.AddWithValue("@endDate", endDate_datetime.ToShortDateString());

            try
            {
                conn.Open();

                getStoreNewSalesReader = getStoreNewSales.ExecuteReader(CommandBehavior.CloseConnection);
                getStoreNewTargetsReader = getStoreNewTarget.ExecuteReader(CommandBehavior.CloseConnection);
                getStorePaygSalesReader = getStorePaygSales.ExecuteReader(CommandBehavior.CloseConnection);
                getStorePaygTargetsReader = getStorePaygTarget.ExecuteReader(CommandBehavior.CloseConnection);
                getStoreRevSalesReader = getStoreRevSales.ExecuteReader(CommandBehavior.CloseConnection);
                getStoreRevTargetReader = getStoreRevSalesTarget.ExecuteReader(CommandBehavior.CloseConnection);
                getStoreBroadSalesReader = getStoreBroadSales.ExecuteReader(CommandBehavior.CloseConnection);
                getStoreBroadTargetsReader = getStoreBroadSalesTarget.ExecuteReader(CommandBehavior.CloseConnection);
                getStoreAccSalesReader = getStoreAccSales.ExecuteReader(CommandBehavior.CloseConnection);
                getStoreAccTargetsReader = getStoreAccTarget.ExecuteReader(CommandBehavior.CloseConnection);

                while (getStoreNewSalesReader.Read() && getStoreNewTargetsReader.Read() && getStorePaygSalesReader.Read() && getStorePaygTargetsReader.Read() && getStoreRevSalesReader.Read() && getStoreRevTargetReader.Read() && getStoreBroadSalesReader.Read() && getStoreBroadTargetsReader.Read() && getStoreAccSalesReader.Read() && getStoreAccTargetsReader.Read())
                {
                    if (getStoreAccSalesReader.GetValue(0) == DBNull.Value)
                    {
                        accsSales = 0;
                    }
                    else
                    {
                        accsSales = Convert.ToInt32(getStoreAccSalesReader.GetValue(0));
                    }

                    if (getStoreNewSalesReader.GetValue(0) == DBNull.Value)
                    {
                        newSales = 0;
                    }
                    else
                    {
                        newSales = Convert.ToInt32(getStoreNewSalesReader.GetValue(0));
                    }

                    if (getStorePaygSalesReader.GetValue(0) == DBNull.Value)
                    {
                        paygSales = 0;
                    }
                    else
                    {
                        paygSales = Convert.ToInt32(getStorePaygSalesReader.GetValue(0));
                    }

                    if (getStoreRevSalesReader.GetValue(0) == DBNull.Value)
                    {
                        revenueSales = 0;
                    }
                    else
                    {
                        revenueSales = Convert.ToInt32(getStoreRevSalesReader.GetValue(0));
                    }

                    if (getStoreBroadSalesReader.GetValue(0) == DBNull.Value)
                    {
                        broadbandSales = 0;
                    }
                    else
                    {
                        broadbandSales = Convert.ToInt32(getStoreBroadSalesReader.GetValue(0));
                    }



                    //Conversion to int32 has to be done as readers return objects
                    newSalesTarget = Convert.ToInt32(getStoreNewTargetsReader.GetValue(0));
                    paygSalesTarget = Convert.ToInt32(getStorePaygTargetsReader.GetValue(0));
                    revenueSalesTarget = Convert.ToInt32(getStoreRevTargetReader.GetValue(0));
                    broadbandSalesTarget = getStoreBroadTargetsReader.GetInt32(0);
                    accsSalesTarget = Convert.ToInt32(getStoreAccTargetsReader.GetValue(0));

                    MTDGetter();

                    currentMonthLabel.Text = now.ToString("MMMM yyyy");

                    newSalesProgressBar.Text = newSales + "/" + newSalesTarget + "\r\nnew \r\nconnections";
                    paygSalesProgressBar.Text = paygSales + "/" + paygSalesTarget + "\r\nPAYG \r\nconnections";
                    revenueProgressBar.Text = revenueSales + "/" + revenueSalesTarget + "\r\nNet Rev";
                    broadbandProgressBar.Text = broadbandSales + "/" + broadbandSalesTarget + "\r\n Broadband";
                    accsProgressBar.Text = accsSales + "/" + accsSalesTarget + "\r\nAccessory \r\nRev";

                    newMTDBar.Text = newSales + "/" + newMTDTarget + "\r\nnew \r\nconnections \r\nMTD";
                    paygMTDBar.Text = paygSales + "/" + paygMTDTarget + "\r\nPAYG \r\nconnections \r\nMTD";
                    revMTDBar.Text = revenueSales + "/" + revMTDTarget + "\r\nNet Rev \r\nMTD";
                    broadMTDBar.Text = broadbandSales + "/" + broadMTDTarget + "\r\n Broadband \r\nMTD";
                    accMTDBar.Text = accsSales + "/" + accMTDTarget + "\r\n Accessory \r\nMTD";

                    TabControl2.SelectedTab.Controls.Add(currentMonthLabel);
                    TabControl2.SelectedTab.Controls.Add(backMonth);
                    TabControl2.SelectedTab.Controls.Add(forwardMonth);
                    TabControl2.SelectedTab.Controls.Add(newSalesProgressBar);
                    TabControl2.SelectedTab.Controls.Add(paygSalesProgressBar);
                    TabControl2.SelectedTab.Controls.Add(revenueProgressBar);
                    TabControl2.SelectedTab.Controls.Add(broadbandProgressBar);
                    TabControl2.SelectedTab.Controls.Add(accsProgressBar);
                    TabControl2.SelectedTab.Controls.Add(newMTDBar);
                    TabControl2.SelectedTab.Controls.Add(paygMTDBar);
                    TabControl2.SelectedTab.Controls.Add(revMTDBar);
                    TabControl2.SelectedTab.Controls.Add(broadMTDBar);
                    TabControl2.SelectedTab.Controls.Add(accMTDBar);
                }

                progressBarsIncrement();
                itemPlace();
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            finally
            {
                conn.Close();
            }
        }

        private void progressBarStyling()
        {
            //Sets the sizes of all the progress bars
            newSalesProgressBar.Size = new Size(300, 300);
            newSalesProgressBar.SubscriptText = null;
            newSalesProgressBar.SuperscriptText = null;
            newSalesProgressBar.InnerColor = Color.White;
            newSalesProgressBar.Font = new Font("Arial", 25, FontStyle.Bold); //Changes the font to readable font in the progfress bars for ease of use

            paygSalesProgressBar.Size = new Size(300, 300);
            paygSalesProgressBar.SubscriptText = null;
            paygSalesProgressBar.SuperscriptText = null;
            paygSalesProgressBar.InnerColor = Color.White;
            paygSalesProgressBar.Font = new Font("Arial", 25, FontStyle.Bold);

            revenueProgressBar.Size = new Size(300, 300);
            revenueProgressBar.SubscriptText = null;
            revenueProgressBar.SuperscriptText = null;
            revenueProgressBar.InnerColor = Color.White;
            revenueProgressBar.Font = new Font("Arial", 25, FontStyle.Bold);

            broadbandProgressBar.Size = new Size(300, 300);
            broadbandProgressBar.SubscriptText = null;
            broadbandProgressBar.SuperscriptText = null;
            broadbandProgressBar.InnerColor = Color.White;
            broadbandProgressBar.Font = new Font("Arial", 25, FontStyle.Bold);

            accsProgressBar.Size = new Size(300, 300);
            accsProgressBar.SubscriptText = null;
            accsProgressBar.SuperscriptText = null;
            accsProgressBar.InnerColor = Color.White;
            accsProgressBar.Font = new Font("Arial", 25, FontStyle.Bold);

            newMTDBar.Size = new Size(250, 250);
            newMTDBar.SubscriptText = null;
            newMTDBar.SuperscriptText = null;
            newMTDBar.InnerColor = Color.White;
            newMTDBar.Font = new Font("Arial", 20, FontStyle.Bold);

            paygMTDBar.Size = new Size(250, 250);
            paygMTDBar.SubscriptText = null;
            paygMTDBar.SuperscriptText = null;
            paygMTDBar.InnerColor = Color.White;
            paygMTDBar.Font = new Font("Arial", 20, FontStyle.Bold);

            broadMTDBar.Size = new Size(250, 250);
            broadMTDBar.SubscriptText = null;
            broadMTDBar.SuperscriptText = null;
            broadMTDBar.InnerColor = Color.White;
            broadMTDBar.Font = new Font("Arial", 20, FontStyle.Bold);

            revMTDBar.Size = new Size(250, 250);
            revMTDBar.SubscriptText = null;
            revMTDBar.SuperscriptText = null;
            revMTDBar.InnerColor = Color.White;
            revMTDBar.Font = new Font("Arial", 20, FontStyle.Bold);

            accMTDBar.Size = new Size(250, 250);
            accMTDBar.SubscriptText = null;
            accMTDBar.SuperscriptText = null;
            accMTDBar.InnerColor = Color.White;
            accMTDBar.Font = new Font("Arial", 20, FontStyle.Bold);

            if (newSales >= newSalesTarget)
            {
                newSalesProgressBar.ProgressColor = Color.Green;
            }

            if (newSales < newSalesTarget && newSales >= newSalesTarget / 2)
            {
                newSalesProgressBar.ProgressColor = Color.Orange;
            }

            if (newSales < newSalesTarget / 2)
            {
                newSalesProgressBar.ProgressColor = Color.Red;
            }

            //PAYG sales colour conditions

            if (paygSales >= paygSalesTarget)
            {
                paygSalesProgressBar.ProgressColor = Color.Green;
            }

            else if (paygSales < paygSalesTarget && paygSales >= paygSalesTarget / 2)
            {
                paygSalesProgressBar.ProgressColor = Color.Orange;
            }

            else if (paygSales < paygSalesTarget / 2)
            {
                paygSalesProgressBar.ProgressColor = Color.Red;
            }

            //Rev sales colour conditions

            if (revenueSales >= revenueSalesTarget)
            {
                revenueProgressBar.ProgressColor = Color.Green;
            }

            else if (revenueSales < revenueSalesTarget && revenueSales >= revenueSalesTarget / 2)
            {
                revenueProgressBar.ProgressColor = Color.Orange;
            }

            else if (revenueSales < revenueSalesTarget / 2)
            {
                revenueProgressBar.ProgressColor = Color.Red;
            }

            //Broadband sales colour conditions

            if(broadbandSales >= broadbandSalesTarget)
            {
                broadbandProgressBar.ProgressColor = Color.Green;
            }

            else if (broadbandSales < broadbandSalesTarget && broadbandSales >= broadbandSalesTarget / 2)
            {
                broadbandProgressBar.ProgressColor = Color.Orange;
            }

            else if (broadbandSales < broadbandSalesTarget / 2)
            {
                broadbandProgressBar.ProgressColor = Color.Red;
            }

            //Accessory sales colour conditions

            if (accsSales >= accsSalesTarget)
            {
                accsProgressBar.ProgressColor = Color.Green;
            }

            else if (accsSales < accsSalesTarget && accsSales >= accsSalesTarget / 2)
            {
                accsProgressBar.ProgressColor = Color.Orange;
            }

            else if (accsSales < accsSalesTarget / 2)
            {
                accsProgressBar.ProgressColor = Color.Red;
            }

            //New MTD sales colour conditions

            if (newSales >= newMTDTarget)
            {
                newMTDBar.ProgressColor = Color.Green;
            }

            if (newSales < newMTDTarget && newSales >= newMTDTarget / 2)
            {
                newMTDBar.ProgressColor = Color.Orange;
            }

            if (newSales < newMTDTarget / 2)
            {
                newMTDBar.ProgressColor = Color.Red;
            }

            //PAYG sales colour conditions

            if (paygSales >= paygMTDTarget)
            {
                paygMTDBar.ProgressColor = Color.Green;
            }

            else if (paygSales < paygMTDTarget && paygSales >= paygMTDTarget / 2)
            {
                paygMTDBar.ProgressColor = Color.Orange;
            }

            else if (paygSales < paygMTDTarget / 2)
            {
                paygMTDBar.ProgressColor = Color.Red;
            }

            //Revenue sales colour conditions

            if (revenueSales >= revMTDTarget)
            {
                revMTDBar.ProgressColor = Color.Green;
            }

            else if (revenueSales < revMTDTarget && revenueSales >= revMTDTarget / 2)
            {
                revMTDBar.ProgressColor = Color.Orange;
            }

            else if (revenueSales < revMTDTarget / 2)
            {
                revMTDBar.ProgressColor = Color.Red;
            }

            //Broadband sales colour conditions

            if (broadbandSales >= broadMTDTarget)
            {
                broadMTDBar.ProgressColor = Color.Green;
            }

            else if (broadbandSales < broadMTDTarget && broadbandSales >= broadMTDTarget / 2)
            {
                broadMTDBar.ProgressColor = Color.Orange;
            }

            else if (broadbandSales < broadMTDTarget / 2)
            {
                broadMTDBar.ProgressColor = Color.Red;
            }

            if (accsSales >= accMTDTarget)
            {
                accMTDBar.ProgressColor = Color.Green;
            }

            else if (accsSales < accMTDTarget && accsSales >= accMTDTarget / 2)
            {
                accMTDBar.ProgressColor = Color.Orange;
            }

            else if (accsSales < accMTDTarget / 2)
            {
                accMTDBar.ProgressColor = Color.Red;
            }

        }

        private void TabControl2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Everytime the user changes tab this checks whether the store view or region view is active and then checks which tab the user is on
            if (storeSelected == true)
            {
                if (TabControl2.SelectedIndex == 0)
                {
                    getStoreTabData();
                }

                else if (TabControl2.SelectedTab == phasesTab)
                {
                    getDeficitsData();
                }

                else
                {
                    getTabData();
                }
            }

            else if (storeSelected == false)
            {
                if (TabControl2.SelectedIndex == 0)
                {
                    getRegionTabData();
                }

                else if (TabControl2.SelectedTab == phasesTab)
                {
                    getDeficitsData();
                }

                else
                {
                    getRegionStoreData();
                }
            }

        }

        private void backMonth_Click(object sender, EventArgs e)
        {
            //Moves the start date back a month
            startDate_datetime = startDate_datetime.AddMonths(-1);
            now = startDate_datetime;

            dateCompare = DateTime.Compare(startDate_datetime, implementationDate); //Compares the new start date against the implementation date

            //If the date is before the implementation date, add a month back on to the start date and call the no date messagebox method
            if (dateCompare < 0) 
            {
                startDate_datetime = startDate_datetime.AddMonths(1);
                now = startDate_datetime;
                noBackData();
            }

            else if (dateCompare == 0) //If the date back is equal to the implementation date then allows the program to go back a month
            {
                dateBack();
            }

            else
            {
                dateBack();
            }

        }

        private void loginButton_Click(object sender, EventArgs e)
        {
            storeSelected = true; //This is the store login button so it sets the store view to true

            storeCode = storeCodeInput.Text; //Makes the store code variable the same as what the user entered
            getData();
        }

        private void forwardMonth_Click(object sender, EventArgs e)
        {
            //Moves the start date forward a month
            startDate_datetime = startDate_datetime.AddMonths(1);
            now = startDate_datetime;

            dateCompare = DateTime.Compare(startDate_datetime, new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1)); //Compares the new start date against the current date

            //If the date is after the current date, take a month off the start date and call the no date messagebox method
            if (dateCompare > 0)
            {
                startDate_datetime = startDate_datetime.AddMonths(-1);
                now = startDate_datetime;
                noForwardDate();
            }

            else if (dateCompare == 0)//If the date back is equal to the current date then allows the program to go forward a month
            {
                dateForward();
            }

            else
            {
                dateForward();
            }
        }

        private void noStoreokButton_Click(object sender, EventArgs e)
        {
            noStoreBox.Hide(); //Used to hide the box otherwise the user has to click the windows close button, which means that the messagebox cannot be called again
        }

        private void noDataBackOkButton_Click(object sender, EventArgs e)
        {
            noDataBackBox.Hide();//Used to hide the box otherwise the user has to click the windows close button, which means that the messagebox cannot be called again
        }

        private void noDataForwardOkButton_Click(object sender, EventArgs e)
        {
            noDataForwardBox.Hide();//Used to hide the box otherwise the user has to click the windows close button, which means that the messagebox cannot be called again
        }

        private void metroLabel1_Click(object sender, EventArgs e)
        {

        }

        private void metroLabel1_Click_1(object sender, EventArgs e)
        {

        }

        private void regionLoginButton_Click(object sender, EventArgs e)
        {
            storeSelected = false; //This is the region login button so it sets the store view to false

            regionCode = regionCodeBox.Text; //Makes the region code variable the same as what the user entered
            getRegionTabs();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            startDate_datetime = new DateTime(now.Year, now.Month, 1); //Initialises the start date and end date variables to the start and end of the month
            endDate_datetime = startDate_datetime.AddMonths(1).AddDays(-1); //Gets the end of the month by adding a month on to the start date and the taking away a day

            TabControl2.Visible = false; //Hides the tabcontrol until the user has logged in
            this.Controls.Add(TabControl2); //Adds the hidden tabcontrol to the form
            TabControl2.Padding = new System.Drawing.Point(35, 10); //Pads the tabs in the tabcontrol so that the names are visible, instead of being squashed
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized; //Starts the form as a maximised program
            this.AutoScroll = true; //Allows scrollbars to appear on the form
            this.AutoScrollMinSize = new System.Drawing.Size(1600, 992); //If the form goes below this size then the scrollbars appear
        }
    }
}