using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace SailingTripNaviagtor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //Global Variables

        private Dictionary<string, int> windspeed = new Dictionary<string, int>();
        DataSet It = new DataSet();
        

        //This function controls what is displayed when user indiates they do have WiFi (WiFi = Yes)
        private void WiFi_Y_click(object sender, EventArgs e)
        {
            webBrowser1.Visible = true;
            Navigation.Visible = true;
            Distance.Visible = true;
            MarineW.Visible = true;
            Weather.Visible = true;
            currents.Visible = true;
            Dpanel.Visible = false;
            navBox.Visible = false;
            navlabel.Visible = false;
            pictureBox8.Visible = false;
            pictureBox9.Visible = false;
            pictureBox10.Visible = false;
            pictureBox11.Visible = false;
            pictureBox12.Visible = false;
            pictureBox13.Visible = false;
            pictureBox14.Visible = false;
            pictureBox15.Visible = false;
            pictureBox16.Visible = false;
            pictureBox17.Visible = false;
            pictureBox18.Visible = false;
            pictureBox19.Visible = false;
            pictureBox20.Visible = false;
            pictureBox21.Visible = false;
            pictureBox22.Visible = false;
            pictureBox23.Visible = false;
            pictureBox24.Visible = false;
        }

        //This function controls what is displayed when user indicates they do noy have WiFi (WiFi = No)
        private void WiFi_N_click(object sender, EventArgs e)
        {
            webBrowser1.Visible = false;
            Navigation.Visible = false;
            Distance.Visible = true;
            MarineW.Visible = true;
            Weather.Visible = false;
            Dpanel.Visible = false;
            navBox.Visible = true;
            navlabel.Visible = true;
            currents.Visible = false;
        }

        //What user sees when selecting the Navigate button. With WiFi utilizes the web, without WiFi utilizes dropdown menu selection of chart images.
        private void Navigation_click(object sender, EventArgs e)
        {
            webBrowser1.Navigate("http://www.charts.noaa.gov/OnLineViewer/18421.shtml");
        }

        //Navigation alternative when No WiFi 
        private void navBox_click(object sender, EventArgs e)
        {
            string[] colname = { "Barnes and Clark", "Bellingham Bay - Hales Passage", "Bellingham Bay", "Blakely", "Cypress", "Decatur", "Friday Harbor", "Hale Passage", "Lopez-Bottom", "Lopex-Top", "Orcas",
                "San Juans", "Shaw", "Stuart", "Sucia", "Waldron"};

            foreach(string col in colname)
            {
                if (!navBox.Items.Contains(col))
                {
                    navBox.Items.Add(col);
                }
            }

        }

        private void NavBox_Selected(object sender, EventArgs e)
        {
            Dpanel.Visible = false;
            pictureBox8.Visible = false;
            pictureBox9.Visible = false;
            pictureBox10.Visible = false;
            pictureBox11.Visible = false;
            pictureBox12.Visible = false;
            pictureBox13.Visible = false;
            pictureBox14.Visible = false;
            pictureBox15.Visible = false;
            pictureBox16.Visible = false;
            pictureBox17.Visible = false;
            pictureBox18.Visible = false;
            pictureBox19.Visible = false;
            pictureBox20.Visible = false;
            pictureBox21.Visible = false;
            pictureBox22.Visible = false;
            pictureBox23.Visible = false;
            pictureBox24.Visible = false;
           
            switch (navBox.SelectedItem.ToString())
            {
                case "Barnes and Clark":
                    pictureBox8.Visible = true;
                    break;
                case "Bellingham Bay - Hales Passage":
                    pictureBox9.Visible = true;
                    break;
                case "Bellingham Bay":
                    pictureBox10.Visible = true;
                    break;
                case "Blakely":
                    pictureBox11.Visible = true;
                    break;
                case "Cypress":
                    pictureBox12.Visible = true;
                    break;
                case "Decatur":
                    pictureBox13.Visible = true;
                    break;
                case "Friday Harbor":
                    pictureBox14.Visible = true;
                    break;
                case "Hale Passage":
                    pictureBox15.Visible = true;
                    break;
                case "Lopez-Bottom":
                    pictureBox16.Visible = true;
                    break;
                case "Lopex-Top":
                    pictureBox17.Visible = true;
                    break;
                case "Orcas":
                    pictureBox18.Visible = true;
                    break;
                case "San Juans":
                    pictureBox19.Visible = true;
                    break;
                case "Shaw":
                    pictureBox20.Visible = true;
                    break;
                case "Stuart":
                    pictureBox21.Visible = true;
                    break;
                case "Sucia":
                    pictureBox22.Visible = true;
                    break;
                case "Waldron":
                    pictureBox23.Visible = true;
                    break;


            }
        }

        //What user sees when selecting the Weather button.  Takes user to a web page pre-selected for the Bellingham area.
        private void Weather_click(object sender, EventArgs e)
        {
            webBrowser1.Navigate("https://www.wunderground.com/cgi-bin/findweather/getForecast?query=98225");
        }

        //When user selects the Distance button they will be directed to a web page for San Juan Island distances.  
        //Without WiFi, the user selects destinations To and From to get Distance from DropDown menus.
        private void Distance_click(object sender, EventArgs e)
        {
            if (WiFi_Y.Checked == true)
            {
                webBrowser1.Navigate("http://nwcruising.net/");
            }
            if (WiFi_N.Checked == true)
            {
                Dpanel.Visible = true;
                pictureBox8.Visible = false;
                pictureBox9.Visible = false;
                pictureBox10.Visible = false;
                pictureBox11.Visible = false;
                pictureBox12.Visible = false;
                pictureBox13.Visible = false;
                pictureBox14.Visible = false;
                pictureBox15.Visible = false;
                pictureBox16.Visible = false;
                pictureBox17.Visible = false;
                pictureBox18.Visible = false;
                pictureBox19.Visible = false;
                pictureBox20.Visible = false;
                pictureBox21.Visible = false;
                pictureBox22.Visible = false;
                pictureBox23.Visible = false;
                pictureBox24.Visible = false;
               
            }
        }

        //When Marine button is clicked, program naviagtes user to a pre-dialed marine weather web page for the Bellingham area.
        private void Marine_click(object sender, EventArgs e)
        {
            if (WiFi_Y.Checked == true)
            {
                webBrowser1.Navigate("https://www.wunderground.com/MAR/PZ/133.html?MR=1");
            }
            if (WiFi_N.Checked == true)
            {
                pictureBox8.Visible = false;
                pictureBox9.Visible = false;
                pictureBox10.Visible = false;
                pictureBox11.Visible = false;
                pictureBox12.Visible = false;
                pictureBox13.Visible = false;
                pictureBox14.Visible = false;
                pictureBox15.Visible = false;
                pictureBox16.Visible = false;
                pictureBox17.Visible = false;
                pictureBox18.Visible = false;
                pictureBox19.Visible = false;
                pictureBox20.Visible = false;
                pictureBox21.Visible = false;
                pictureBox22.Visible = false;
                pictureBox23.Visible = false;
                pictureBox24.Visible = true;
                Dpanel.Visible = false;

            }
        }

        private void currents_click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://deepzoom.com");
        }

        //Upon clicking the Extra Tools selection, the Drop Down men is populated with choices.
        private void ExtraT_click(object sender, EventArgs e)
        {
            string[] colname = { "Points of Sail", "Cleat Hitch", "Figure Eight", "Jibe", "Round Turn", "Tacking", "Bowline" };

            foreach (string tool in colname)
            {
                if (!ExtraT.Items.Contains(tool))
                {
                    ExtraT.Items.Add(tool);
                }

                ExtraT.ValueMember = "Item";
                ExtraT.DisplayMember = "Item";

            }
        }

        //upon user selecting their Extra Tools option, a corresponsing image will be displayed.
        private void ExtraT_select(object sender, EventArgs e)
        {
            pictureBox1.Visible = false;
            pictureBox2.Visible = false;
            pictureBox3.Visible = false;
            pictureBox4.Visible = false;
            pictureBox5.Visible = false;
            pictureBox6.Visible = false;
            pictureBox7.Visible = false;

            switch (ExtraT.SelectedItem.ToString())
            {
                case "Points of Sail":
                    pictureBox1.Visible = true;
                    break;
                case "Bowline":
                    pictureBox2.Visible = true;
                    break;
                case "Cleat Hitch":
                    pictureBox3.Visible = true;
                    break;
                case "Figure Eight":
                    pictureBox4.Visible = true;
                    break;
                case "Jibe":
                    pictureBox5.Visible = true;
                    break;
                case "Round Turn":
                    pictureBox6.Visible = true;
                    break;
                case "Tacking":
                    pictureBox7.Visible = true;
                    break;

            }
        }// end of ExtraT select


        //Distance controls for when user does not have WiFi.

        //From Index drop down menu is populated from distances.xls being imported and referencing the corresponding column.
        private void FromIndex_click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            ds = ConvertToExcel.ExcelImport.ImportExcelXLS("distances.xls", true);

            for(int intcount = 0; intcount <ds.Tables[0].Rows.Count; intcount++)
            {
                int value = 0;
                var val = ds.Tables[0].Rows[intcount][value].ToString();

                if (!FromIndex.Items.Contains(val))
                {
                    FromIndex.Items.Add(val);
                }
            }

            FromIndex.Sorted = true;
            //FromIndex.ValueMember = "Index1";
            //FromIndex.DisplayMember = "Index1";
        }

        
        //From Destinations are queried from distances.xls based on From Index selected.
        private void From_Click(object sender, EventArgs e)
        {
            string colname = "Index1";
            string Selected = FromIndex.SelectedItem.ToString();

            DataSet ds = new DataSet();
            ds = ConvertToExcel.ExcelImport.ImportExcelXLS_Query("distances.xls", true, Selected, colname);

            for (int intcount = 0; intcount < ds.Tables[0].Rows.Count; intcount++)
            {
                int value = 2;
                var val = ds.Tables[0].Rows[intcount][value].ToString();

                if (!From.Items.Contains(val))
                {
                    From.Items.Add(val);
                }
            }

            From.Sorted = true;
            //From.ValueMember = "From";
            //From.DisplayMember = "From";

        }

        //To Index drop down menu is populated from distances.xls being imported and referencing the corresponding column.
        private void ToIndex_click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            ds = ConvertToExcel.ExcelImport.ImportExcelXLS("distances.xls", true);

            for (int intcount = 0; intcount < ds.Tables[0].Rows.Count; intcount++)
            {
                int value = 1;
                var val = ds.Tables[0].Rows[intcount][value].ToString();

                if (!ToIndex.Items.Contains(val))
                {
                    ToIndex.Items.Add(val);
                }
            }

            ToIndex.Sorted = true;
            //ToIndex.ValueMember = "Index2";
            //ToIndex.DisplayMember = "Index2";
        }

        //To Destinations are queried from distances.xls based on To Index selected.
        private void To_Click(object sender, EventArgs e)
        {
            To.Items.Clear();

            string colname = "Index2";
            string Selected = ToIndex.SelectedItem.ToString();

            DataSet ds = new DataSet();
            ds = ConvertToExcel.ExcelImport.ImportExcelXLS_Query("distances.xls", true, Selected, colname);

            for (int intcount = 0; intcount < ds.Tables[0].Rows.Count; intcount++)
            {
                int value = 3;
                var val = ds.Tables[0].Rows[intcount][value].ToString();

                if (!To.Items.Contains(val))
                {
                    To.Items.Add(val);
                }
            }

            To.Sorted = true;
            //To.ValueMember = "To";
            //To.DisplayMember = "To";
        }

        //Distance text box displays distance based on the concatenation of From and To locations.
        //Queries FromTo column of distances.xls and returns matching Distance 
        private void Dist_changed(object sender, EventArgs e)
        {
            string colname = "FromTo";
            string Selected = From.SelectedItem.ToString() + To.SelectedItem.ToString();

            DataSet ds = new DataSet();
            ds = ConvertToExcel.ExcelImport.ImportExcelXLS_Query("distances.xls", true, Selected, colname);

            try
            {
                
                    int value = 5;
                    var val = ds.Tables[0].Rows[0][value].ToString();
                    Dist.Text = val.ToString();
           
            }
            catch
            {
                MessageBox.Show("Please select a different location.");
            }
         }
            
        

        //THE BELOW FUNCTIONS HANDLE THE ITINERAY SECTION

        private void Itinerary_click(object sender, EventArgs e)
        {
            for(int i = 1; i< 11; i++)
            {
                Itinerary.Items.Add(i);
            }
        }

        private void WindSpeed_click(object sender, EventArgs e)
        {
            windspeed = new Dictionary<string, int>{{"Motor Only (No Wind)", 6}, {"Light Wind (5 mph)", 2}, {"Moderate Wind (10-12 mph)", 4}, { "Strong Wind (15-20 mph)", 7}};

            foreach(string key in windspeed.Keys)
            {
                if (!WindSpeed.Items.Contains(key))
                {
                    WindSpeed.Items.Add(key);
                }             
            }
        }

        private void WindSpeed_Select(object sender, EventArgs e)
        {
            string wind = Convert.ToString(WindSpeed.SelectedItem);
            BoatSpeed.Text = Convert.ToString(windspeed[wind]);
        }

        private void TravelT_fill(object sender, EventArgs e)
        {
            double Time = Convert.ToDouble(D.Text) / Convert.ToDouble(BoatSpeed.Text);
            TravelT.Text = Convert.ToString(Time);
        }

        private void ArrT_Fill(object sender, EventArgs e)
        {
            double arrive = Convert.ToDouble(TravelT.Text) + Convert.ToDouble(DepT.Text);

            ArrT.Text = Convert.ToString(arrive);
        }

        private void StartTrip_click(object sender, EventArgs e)
        {
            It = ConvertToExcel.ExcelImport.ImportExcelXLS("Itinerary.xls", true);

            MessageBox.Show("Ahoy! Enjoy Your Trip!");
        }

        private void AddIT_click(object sender, EventArgs e)
        {
            try
            {
                int i = Convert.ToInt32(Itinerary.Text) - 1;

                It.Tables[0].Rows[i]["Notes"] = Notes.Text;
                It.Tables[0].Rows[i]["DepartureDate"] = Departure.Value;
                It.Tables[0].Rows[i]["ArrivalDate"] = Arrive.Value;
                It.Tables[0].Rows[i]["DepartureTime"] = DepT.Text;
                It.Tables[0].Rows[i]["ArrivalTime"] = ArrT.Text;
                It.Tables[0].Rows[i]["WindSpeed"] = WindSpeed.Text;
                It.Tables[0].Rows[i]["BoatSpeed"] = BoatSpeed.Text;
                It.Tables[0].Rows[i]["Distance"] = D.Text;
                It.Tables[0].Rows[i]["TravelTime"] = TravelT.Text;

                MessageBox.Show("Your Trip has been successfully added to your Itinerary");
            }
            catch
            {
                MessageBox.Show("Remember to select 'Start My Trip', and that all itinerary items are filled out.");
            }
        }

        private void SaveIT_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application oexcel = new Microsoft.Office.Interop.Excel.Application();

            try
            {
                string path = AppDomain.CurrentDomain.BaseDirectory;
                object misValue = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Excel.Workbook obook = oexcel.Workbooks.Add(misValue);
                Microsoft.Office.Interop.Excel.Worksheet osheet = new Microsoft.Office.Interop.Excel.Worksheet();

                osheet = (Microsoft.Office.Interop.Excel.Worksheet)obook.Sheets["Sheet1"];
                obook.Worksheets.Add(misValue);

                int colIndex = 0;
                int rowIndex = 1;

                
                {
                    foreach (DataColumn dc in It.Tables[0].Columns)
                    {
                        colIndex++;
                        osheet.Cells[1, colIndex] = dc.ColumnName;
                    }
                    foreach (DataRow dr in It.Tables[0].Rows)
                    {
                        rowIndex++;
                        colIndex = 0;

                        foreach (DataColumn dc in It.Tables[0].Columns)
                        {
                            colIndex++;
                            osheet.Cells[rowIndex, colIndex] = dr[dc.ColumnName];
                        }
                    }
                }
                
                osheet.Columns.AutoFit();
       
                string filepath = "C:\\Temp\\Itinerary";

                //Release and terminate excel

                obook.SaveAs(filepath);
                obook.Close();
                oexcel.Quit();
             
                GC.Collect();

            }

            catch (Exception ex)
            {
                oexcel.Quit();
               
            }

            MessageBox.Show("Export Successful. File Location = C:\\Temp\\Itinerary");

        }// end of Save IT
           

    }//end of partial class

}// end of namespace