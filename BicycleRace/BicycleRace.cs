using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Text.RegularExpressions;

namespace BicycleRace
{
    public partial class BicycleRace : Form
    {
        private DateTime stop;
        private List<string> categories = new List<string>();
        private List<string> distance = new List<string>();
        private List<string> starttime = new List<string>();
        Regex pattern = new Regex("^[1-9][0-9]?$");
        private TimeSpan elteltido = new TimeSpan(0, 5, 0);

        public BicycleRace()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            categories.Add("JF");
            categories.Add("JL");
            categories.Add("F");
            categories.Add("N");
            categories.Add("SF");
            categories.Add("SN");

            distance.Add("Classic 20 Km");
            distance.Add("Classic 40 Km");
            distance.Add("Family 2 Km");

            starttime.Add("08:00");
            starttime.Add("08:15");
            starttime.Add("08:30");
            starttime.Add("08:45");
            starttime.Add("09:00");
            starttime.Add("09:15");
            starttime.Add("09:30");
            starttime.Add("09:45");
            starttime.Add("10:00");
            starttime.Add("12:00");

            this.dataGridViewStyle();

            dataGridView1.CellClick += (sender2, e2) => timeHandling_CellClick(sender2, e2, dataGridView1);
            dataGridView1.CellContentClick += (sender2, e2) => sexHandling_CellChecked(sender2, e2, dataGridView1);
            dataGridView1.CellValueChanged += (sender2, e2) => ageHandling_ValueChanged(sender2, e2, dataGridView1);
            dataGridView1.CellValueChanged += new DataGridViewCellEventHandler(Classic_StartNumber);

            dataGridView2.CellClick += (sender2, e2) => timeHandling_CellClick(sender2, e2, dataGridView2);
            dataGridView2.CellContentClick += (sender2, e2) => sexHandling_CellChecked(sender2, e2, dataGridView2);
            dataGridView2.CellValueChanged += (sender2, e2) => ageHandling_ValueChanged(sender2, e2, dataGridView2);
            dataGridView2.CellValueChanged += new DataGridViewCellEventHandler(MTB_StartNumber);

            dataGridView3.CellClick += (sender2, e2) => timeHandling_CellClick(sender2, e2, dataGridView3);
            dataGridView3.CellContentClick += (sender2, e2) => sexHandling_CellChecked(sender2, e2, dataGridView3);
            dataGridView3.CellValueChanged += (sender2, e2) => ageHandling_ValueChanged(sender2, e2, dataGridView3);
            dataGridView3.CellValueChanged += new DataGridViewCellEventHandler(ROAD_StartNumber);
        }

        //Sorszám és rajtszám beállítása
        void Classic_StartNumber(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 1)
            {
                dataGridView1.Rows[e.RowIndex].Cells[0].Value = e.RowIndex + 1;
                dataGridView1.Rows[e.RowIndex].Cells[8].Value = Convert.ToInt16(dataGridView1.Rows[e.RowIndex].Cells[0].Value) + 499;
            }
            else { return; }
        }

        void MTB_StartNumber(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 1)
            {
                dataGridView2.Rows[e.RowIndex].Cells[0].Value = e.RowIndex + 1;
                dataGridView2.Rows[e.RowIndex].Cells[6].Value = "MTB 45 Km";
                if (e.RowIndex == 0)
                {
                    dataGridView2.Rows[e.RowIndex].Cells[8].Value = Convert.ToInt16(dataGridView2.Rows[e.RowIndex].Cells[0].Value);
                }
                else
                {
                    dataGridView2.Rows[e.RowIndex].Cells[8].Value = Convert.ToInt16(dataGridView2.Rows[e.RowIndex - 1].Cells[8].Value) + 2;
                }
            }
            else { return; }
        }

        void ROAD_StartNumber(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 1)
            {
                dataGridView3.Rows[e.RowIndex].Cells[0].Value = e.RowIndex + 1;
                dataGridView3.Rows[e.RowIndex].Cells[6].Value = "ROAD 70 Km";
                if (e.RowIndex == 0)
                {
                    dataGridView3.Rows[e.RowIndex].Cells[8].Value = Convert.ToInt16(dataGridView3.Rows[e.RowIndex].Cells[0].Value) + 1;
                }
                else
                {
                    dataGridView3.Rows[e.RowIndex].Cells[8].Value = Convert.ToInt16(dataGridView3.Rows[e.RowIndex - 1].Cells[8].Value) + 2;
                }
            }
            else { return; }
        }

        //Időkezelés
        void timeHandling_CellClick(object sender, DataGridViewCellEventArgs e, DataGridView dgv)
        {
            if (e.ColumnIndex == 10)
            {
                if (dgv.Rows[e.RowIndex].Cells[11].Value == null || dgv.Rows[e.RowIndex].Cells[11].Value.ToString() == "")
                {
                    dgv.Rows[e.RowIndex].Cells[11].Value = DateTime.Now.ToString("HH:mm:ss");
                }
                else 
                {
                    MessageBox.Show("A versenyzőnek már van elrajtolási ideje", "Információ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (e.ColumnIndex == 12)
            {
                if (dgv.Rows[e.RowIndex].Cells[13].Value == null || dgv.Rows[e.RowIndex].Cells[13].Value.ToString() == "")
                {
                    try
                    {
                        stop = DateTime.Now;
                        dgv.Rows[e.RowIndex].Cells[13].Value = stop.ToString("HH:mm:ss");
                        TimeSpan result = stop - DateTime.Parse(dgv.Rows[e.RowIndex].Cells[11].Value.ToString());
                        dgv.Rows[e.RowIndex].Cells[14].Value = result.ToString(@"hh\:mm\:ss");

                        // Helyezések beállítása színekkel
                        colorResults(dgv);
                    }
                    catch (System.Exception ex)
                    {
                        if (ex is System.NullReferenceException || ex is System.FormatException)
                        {
                            dgv.Rows[e.RowIndex].Cells[13].Value = "";
                            MessageBox.Show("Először a Start gombbal el kell indítani az időt", "Figyelmeztetés!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("A versenyzőnek már van beérkezési ideje", "Információ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }      
            }
            else if (e.ColumnIndex == 15)
            {
                try
                {
                    if ((dgv.Rows[e.RowIndex].Cells[11].Value != null && dgv.Rows[e.RowIndex].Cells[13].Value == null) || (dgv.Rows[e.RowIndex].Cells[11].Value.ToString() != "" && dgv.Rows[e.RowIndex].Cells[13].Value.ToString() == ""))
                    {
                        DialogResult dialogResult = MessageBox.Show("Biztosan törölni szeretné az elrajtolási időt?", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                        if (dialogResult == DialogResult.OK)
                        {
                            dgv.Rows[e.RowIndex].Cells[11].Value = "";
                            dgv.Rows[e.RowIndex].Cells[11].Value = null;
                        }
                        else { return; }
                    }
                    else
                    {
                        MessageBox.Show("Kizárólag véletlen elindítás esetén", "Figyelmeztetés!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                catch (System.NullReferenceException)
                {
                    MessageBox.Show("Kizárólag véletlen elindítás esetén", "Figyelmeztetés!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else { return; }
        }


        //Nő/férfi kiválasztására szolgáló választóboxok kezelése, és a Korcsoport meghatározása a beírt kor és a KIVÁLASZTOTT NEM alapján
        void sexHandling_CellChecked(object sender, DataGridViewCellEventArgs e, DataGridView dgv)
        {
            if (e.ColumnIndex == 3) 
            {
                dgv.Rows[e.RowIndex].Cells["Man"].Value = false;
                dgv.Rows[e.RowIndex].Cells["Woman"].Value = !Convert.ToBoolean(dgv.Rows[e.RowIndex].Cells["Woman"].Value);
                if (Convert.ToBoolean(dgv.Rows[e.RowIndex].Cells["Woman"].Value)==false)
                {
                    dgv.Rows[e.RowIndex].Cells[5].Value = "";
                }


                if ((dgv.Rows[e.RowIndex].Cells[2].Value == null || dgv.Rows[e.RowIndex].Cells[2].Value.ToString() != "") && (Convert.ToInt16(dgv.Rows[e.RowIndex].Cells[2].Value) >= 1 && Convert.ToInt16(dgv.Rows[e.RowIndex].Cells[2].Value) <= 18) && Convert.ToBoolean(dgv.Rows[e.RowIndex].Cells["Woman"].Value))
                {
                    dgv.Rows[e.RowIndex].Cells[5].Value = categories[1];
                }
                else if ((dgv.Rows[e.RowIndex].Cells[2].Value == null || dgv.Rows[e.RowIndex].Cells[2].Value.ToString() != "") && (Convert.ToInt16(dgv.Rows[e.RowIndex].Cells[2].Value) >= 19 && Convert.ToInt16(dgv.Rows[e.RowIndex].Cells[2].Value) <= 55) && Convert.ToBoolean(dgv.Rows[e.RowIndex].Cells["Woman"].Value))
                {
                    dgv.Rows[e.RowIndex].Cells[5].Value = categories[3];
                }
                else if ((dgv.Rows[e.RowIndex].Cells[2].Value == null || dgv.Rows[e.RowIndex].Cells[2].Value.ToString() != "") && Convert.ToInt16(dgv.Rows[e.RowIndex].Cells[2].Value) >= 56 && Convert.ToBoolean(dgv.Rows[e.RowIndex].Cells["Woman"].Value))
                {
                    dgv.Rows[e.RowIndex].Cells[5].Value = categories[5];
                }
                else { return; }
            }
            else if (e.ColumnIndex == 4)
            {
                dgv.Rows[e.RowIndex].Cells["Woman"].Value = false;
                dgv.Rows[e.RowIndex].Cells["Man"].Value = !Convert.ToBoolean(dgv.Rows[e.RowIndex].Cells["Man"].Value);
                if (Convert.ToBoolean(dgv.Rows[e.RowIndex].Cells["Man"].Value) == false)
                {
                    dgv.Rows[e.RowIndex].Cells[5].Value = "";
                }

                if ((dgv.Rows[e.RowIndex].Cells[2].Value == null || dgv.Rows[e.RowIndex].Cells[2].Value.ToString() != "") && (Convert.ToInt16(dgv.Rows[e.RowIndex].Cells[2].Value) >= 1 && Convert.ToInt16(dgv.Rows[e.RowIndex].Cells[2].Value) <= 18) && Convert.ToBoolean(dgv.Rows[e.RowIndex].Cells["Man"].Value))
                {
                    dgv.Rows[e.RowIndex].Cells[5].Value = categories[0];
                }
                else if ((dgv.Rows[e.RowIndex].Cells[2].Value == null || dgv.Rows[e.RowIndex].Cells[2].Value.ToString() != "") && (Convert.ToInt16(dgv.Rows[e.RowIndex].Cells[2].Value) >= 19 && Convert.ToInt16(dgv.Rows[e.RowIndex].Cells[2].Value) <= 55) && Convert.ToBoolean(dgv.Rows[e.RowIndex].Cells["Man"].Value))
                {
                    dgv.Rows[e.RowIndex].Cells[5].Value = categories[2];
                }
                else if ((dgv.Rows[e.RowIndex].Cells[2].Value == null || dgv.Rows[e.RowIndex].Cells[2].Value.ToString() != "") && Convert.ToInt16(dgv.Rows[e.RowIndex].Cells[2].Value) >= 56 && Convert.ToBoolean(dgv.Rows[e.RowIndex].Cells["Man"].Value))
                {
                    dgv.Rows[e.RowIndex].Cells[5].Value = categories[4];
                }
                else { return; }
            }
            else { return; }
        }

        //Korcsoport meghatározása a BEÍRT KOR és a kiválasztott nem alapján
        void ageHandling_ValueChanged(object sender, DataGridViewCellEventArgs e, DataGridView dgv)
        {
            if (e.ColumnIndex == 2) 
            {
                string s;
                try 
                { 
                    s = dgv.Rows[e.RowIndex].Cells[2].Value.ToString(); 
                }
                catch (System.NullReferenceException) 
                { 
                    s = "";  
                }
                if (pattern.IsMatch(s) || s == "")
                {
                    if ((dgv.Rows[e.RowIndex].Cells[2].Value == null || dgv.Rows[e.RowIndex].Cells[2].Value.ToString() != "") && (Convert.ToInt16(dgv.Rows[e.RowIndex].Cells[2].Value) >= 1 && Convert.ToInt16(dgv.Rows[e.RowIndex].Cells[2].Value) <= 18) && Convert.ToBoolean(dgv.Rows[e.RowIndex].Cells["Woman"].Value))
                    {
                        dgv.Rows[e.RowIndex].Cells[5].Value = categories[1];
                    }
                    else if ((dgv.Rows[e.RowIndex].Cells[2].Value == null || dgv.Rows[e.RowIndex].Cells[2].Value.ToString() != "") && (Convert.ToInt16(dgv.Rows[e.RowIndex].Cells[2].Value) >= 19 && Convert.ToInt16(dgv.Rows[e.RowIndex].Cells[2].Value) <= 55) && Convert.ToBoolean(dgv.Rows[e.RowIndex].Cells["Woman"].Value))
                    {
                        dgv.Rows[e.RowIndex].Cells[5].Value = categories[3];
                    }
                    else if ((dgv.Rows[e.RowIndex].Cells[2].Value == null || dgv.Rows[e.RowIndex].Cells[2].Value.ToString() != "") && Convert.ToInt16(dgv.Rows[e.RowIndex].Cells[2].Value) >= 56 && Convert.ToBoolean(dgv.Rows[e.RowIndex].Cells["Woman"].Value))
                    {
                        dgv.Rows[e.RowIndex].Cells[5].Value = categories[5];
                    }
                    else if ((dgv.Rows[e.RowIndex].Cells[2].Value == null || dgv.Rows[e.RowIndex].Cells[2].Value.ToString() != "") && (Convert.ToInt16(dgv.Rows[e.RowIndex].Cells[2].Value) >= 1 && Convert.ToInt16(dgv.Rows[e.RowIndex].Cells[2].Value) <= 18) && Convert.ToBoolean(dgv.Rows[e.RowIndex].Cells["Man"].Value))
                    {
                        dgv.Rows[e.RowIndex].Cells[5].Value = categories[0];
                    }
                    else if ((dgv.Rows[e.RowIndex].Cells[2].Value == null || dgv.Rows[e.RowIndex].Cells[2].Value.ToString() != "") && (Convert.ToInt16(dgv.Rows[e.RowIndex].Cells[2].Value) >= 19 && Convert.ToInt16(dgv.Rows[e.RowIndex].Cells[2].Value) <= 55) && Convert.ToBoolean(dgv.Rows[e.RowIndex].Cells["Man"].Value))
                    {
                        dgv.Rows[e.RowIndex].Cells[5].Value = categories[2];
                    }
                    else if ((dgv.Rows[e.RowIndex].Cells[2].Value == null || dgv.Rows[e.RowIndex].Cells[2].Value.ToString() != "") && Convert.ToInt16(dgv.Rows[e.RowIndex].Cells[2].Value) >= 56 && Convert.ToBoolean(dgv.Rows[e.RowIndex].Cells["Man"].Value))
                    {
                        dgv.Rows[e.RowIndex].Cells[5].Value = categories[4];
                    }
                    else
                    {
                        dgv.Rows[e.RowIndex].Cells[5].Value = "";
                        return;
                    }
                }
                else 
                {
                    dgv.Rows[e.RowIndex].Cells[2].Value = "";
                    MessageBox.Show("Az életkort 1 - 99 tartmányban kell megadni!", "Figyelmeztetés!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        //Adattábla beállítása
        void dataGridViewStyle()
        {
            //Classic DGV ----------------------------------------------------------------------------------------------------------

            DataGridViewTextBoxColumn number =
            new DataGridViewTextBoxColumn();
            number.Name = "Number";
            number.HeaderText = "Sorszám";

            dataGridView1.Columns.Insert(0, number);

            DataGridViewTextBoxColumn name =
            new DataGridViewTextBoxColumn();
            name.Name = "Name";
            name.HeaderText = "Vezetéknév és keresztnév";

            dataGridView1.Columns.Insert(1, name);

            DataGridViewTextBoxColumn age =
            new DataGridViewTextBoxColumn();
            age.Name = "Age";
            age.HeaderText = "Életkor";

            dataGridView1.Columns.Insert(2, age);

            DataGridViewCheckBoxColumn checkboxColumnWoman =
            new DataGridViewCheckBoxColumn();
            checkboxColumnWoman.Name = "Woman";
            checkboxColumnWoman.HeaderText = "Nő";

            dataGridView1.Columns.Insert(3, checkboxColumnWoman);

            DataGridViewCheckBoxColumn checkboxColumnMan =
            new DataGridViewCheckBoxColumn();
            checkboxColumnMan.Name = "Man";
            checkboxColumnMan.HeaderText = "Férfi";

            dataGridView1.Columns.Insert(4, checkboxColumnMan);

            DataGridViewTextBoxColumn ageGroup =
            new DataGridViewTextBoxColumn();
            ageGroup.Name = "AgeGroup";
            ageGroup.HeaderText = "Korcsoport";

            dataGridView1.Columns.Insert(5, ageGroup);

            DataGridViewComboBoxColumn comboboxColumnDistance =
            new DataGridViewComboBoxColumn();
            foreach (string str in distance) { comboboxColumnDistance.Items.Add(str); }
            comboboxColumnDistance.AutoComplete = true;
            comboboxColumnDistance.Name = "Táv";
            comboboxColumnDistance.DropDownWidth = 120;
            //comboboxColumnDistance.DefaultCellStyle.NullValue = "";

            dataGridView1.Columns.Insert(6, comboboxColumnDistance);

            DataGridViewTextBoxColumn team =
            new DataGridViewTextBoxColumn();
            team.Name = "Team";
            team.HeaderText = "Csapatnév";

            dataGridView1.Columns.Insert(7, team);

            DataGridViewTextBoxColumn startNumber =
            new DataGridViewTextBoxColumn();
            startNumber.Name = "StartNumber";
            startNumber.HeaderText = "Rajtszám";

            dataGridView1.Columns.Insert(8, startNumber);

            DataGridViewComboBoxColumn comboboxColumnStartTime =
            new DataGridViewComboBoxColumn();
            foreach (string str in starttime) { comboboxColumnStartTime.Items.Add(str); }
            comboboxColumnStartTime.AutoComplete = true;
            comboboxColumnStartTime.Name = "Elrajtolási idő";
            comboboxColumnStartTime.DropDownWidth = 60;
            //comboboxColumnStartTime.DefaultCellStyle.NullValue = "8:00";

            dataGridView1.Columns.Insert(9, comboboxColumnStartTime);

            DataGridViewButtonColumn buttonColumnStart =
            new DataGridViewButtonColumn();
            buttonColumnStart.Name = "Start";
            buttonColumnStart.HeaderText = "";
            buttonColumnStart.Text = "Start";
            buttonColumnStart.UseColumnTextForButtonValue = true;

            dataGridView1.Columns.Insert(10, buttonColumnStart);

            DataGridViewTextBoxColumn startTime =
            new DataGridViewTextBoxColumn();
            startTime.Name = "StartTime";
            startTime.HeaderText = "Indulás";

            dataGridView1.Columns.Insert(11, startTime);

            DataGridViewButtonColumn buttonColumnStop =
            new DataGridViewButtonColumn();
            buttonColumnStop.Name = "Stop";
            buttonColumnStop.HeaderText = "";
            buttonColumnStop.Text = "Stop";
            buttonColumnStop.UseColumnTextForButtonValue = true;

            dataGridView1.Columns.Insert(12, buttonColumnStop);

            DataGridViewTextBoxColumn stopTime =
            new DataGridViewTextBoxColumn();
            stopTime.Name = "StopTime";
            stopTime.HeaderText = "Érkezés";

            dataGridView1.Columns.Insert(13, stopTime);

            DataGridViewTextBoxColumn result =
            new DataGridViewTextBoxColumn();
            result.Name = "Result";
            result.HeaderText = "Időeredmény";

            dataGridView1.Columns.Insert(14, result);

            DataGridViewButtonColumn buttonColumnNull =
            new DataGridViewButtonColumn();
            buttonColumnNull.Name = "Null";
            buttonColumnNull.HeaderText = "";
            buttonColumnNull.Text = "Null";
            buttonColumnNull.UseColumnTextForButtonValue = true;
            /*buttonColumnNull.CellTemplate.Style.BackColor = Color.Tomato;
            buttonColumnNull.CellTemplate.Style.ForeColor = Color.White;
            buttonColumnNull.FlatStyle = FlatStyle.Flat;*/

            dataGridView1.Columns.Insert(15, buttonColumnNull);

            //MTB DGV -------------------------------------------------------------------------------------------------------------

            DataGridViewTextBoxColumn number2 =
            new DataGridViewTextBoxColumn();
            number2.Name = "Number";
            number2.HeaderText = "Sorszám";

            dataGridView2.Columns.Insert(0, number2);

            DataGridViewTextBoxColumn name2 =
            new DataGridViewTextBoxColumn();
            name2.Name = "Name";
            name2.HeaderText = "Vezetéknév és keresztnév";

            dataGridView2.Columns.Insert(1, name2);

            DataGridViewTextBoxColumn age2 =
            new DataGridViewTextBoxColumn();
            age2.Name = "Age";
            age2.HeaderText = "Életkor";

            dataGridView2.Columns.Insert(2, age2);

            DataGridViewCheckBoxColumn checkboxColumnWoman2 =
            new DataGridViewCheckBoxColumn();
            checkboxColumnWoman2.Name = "Woman";
            checkboxColumnWoman2.HeaderText = "Nő";

            dataGridView2.Columns.Insert(3, checkboxColumnWoman2);

            DataGridViewCheckBoxColumn checkboxColumnMan2 =
            new DataGridViewCheckBoxColumn();
            checkboxColumnMan2.Name = "Man";
            checkboxColumnMan2.HeaderText = "Férfi";

            dataGridView2.Columns.Insert(4, checkboxColumnMan2);

            DataGridViewTextBoxColumn ageGroup2 =
            new DataGridViewTextBoxColumn();
            ageGroup2.Name = "AgeGroup";
            ageGroup2.HeaderText = "Korcsoport";

            dataGridView2.Columns.Insert(5, ageGroup2);

            DataGridViewTextBoxColumn km2 =
            new DataGridViewTextBoxColumn();
            km2.Name = "Km";
            km2.HeaderText = "Táv";

            dataGridView2.Columns.Insert(6, km2);

            DataGridViewTextBoxColumn team2 =
            new DataGridViewTextBoxColumn();
            team2.Name = "Team";
            team2.HeaderText = "Csapatnév";

            dataGridView2.Columns.Insert(7, team2);

            DataGridViewTextBoxColumn startNumber2 =
            new DataGridViewTextBoxColumn();
            startNumber2.Name = "StartNumber";
            startNumber2.HeaderText = "Rajtszám";

            dataGridView2.Columns.Insert(8, startNumber2);

            DataGridViewComboBoxColumn comboboxColumnStartTime2 =
            new DataGridViewComboBoxColumn();
            foreach (string str in starttime) { comboboxColumnStartTime2.Items.Add(str); }
            comboboxColumnStartTime2.AutoComplete = true;
            comboboxColumnStartTime2.Name = "Elrajtolási idő";
            comboboxColumnStartTime2.DropDownWidth = 60;
            //comboboxColumnStartTime2.DefaultCellStyle.NullValue = "8:00";

            dataGridView2.Columns.Insert(9, comboboxColumnStartTime2);

            DataGridViewButtonColumn buttonColumnStart2 =
            new DataGridViewButtonColumn();
            buttonColumnStart2.Name = "Start";
            buttonColumnStart2.HeaderText = "";
            buttonColumnStart2.Text = "Start";
            buttonColumnStart2.UseColumnTextForButtonValue = true;

            dataGridView2.Columns.Insert(10, buttonColumnStart2);

            DataGridViewTextBoxColumn startTime2 =
            new DataGridViewTextBoxColumn();
            startTime2.Name = "StartTime";
            startTime2.HeaderText = "Indulás";

            dataGridView2.Columns.Insert(11, startTime2);

            DataGridViewButtonColumn buttonColumnStop2 =
            new DataGridViewButtonColumn();
            buttonColumnStop2.Name = "Stop";
            buttonColumnStop2.HeaderText = "";
            buttonColumnStop2.Text = "Stop";
            buttonColumnStop2.UseColumnTextForButtonValue = true;

            dataGridView2.Columns.Insert(12, buttonColumnStop2);

            DataGridViewTextBoxColumn stopTime2 =
            new DataGridViewTextBoxColumn();
            stopTime2.Name = "StopTime";
            stopTime2.HeaderText = "Érkezés";

            dataGridView2.Columns.Insert(13, stopTime2);

            DataGridViewTextBoxColumn result2 =
            new DataGridViewTextBoxColumn();
            result2.Name = "Result";
            result2.HeaderText = "Időeredmény";

            dataGridView2.Columns.Insert(14, result2);

            DataGridViewButtonColumn buttonColumnNull2 =
            new DataGridViewButtonColumn();
            buttonColumnNull2.Name = "Null";
            buttonColumnNull2.HeaderText = "";
            buttonColumnNull2.Text = "Null";
            buttonColumnNull2.UseColumnTextForButtonValue = true;
            /*buttonColumnNull2.CellTemplate.Style.BackColor = Color.Tomato;
            buttonColumnNull2.CellTemplate.Style.ForeColor = Color.White;
            buttonColumnNull2.FlatStyle = FlatStyle.Flat;*/

            dataGridView2.Columns.Insert(15, buttonColumnNull2);

            //ROAD DGV --------------------------------------------------------------------------------------------------------------

            DataGridViewTextBoxColumn number3 =
            new DataGridViewTextBoxColumn();
            number3.Name = "Number";
            number3.HeaderText = "Sorszám";

            dataGridView3.Columns.Insert(0, number3);

            DataGridViewTextBoxColumn name3 =
            new DataGridViewTextBoxColumn();
            name3.Name = "Name";
            name3.HeaderText = "Vezetéknév és keresztnév";

            dataGridView3.Columns.Insert(1, name3);

            DataGridViewTextBoxColumn age3 =
            new DataGridViewTextBoxColumn();
            age3.Name = "Age";
            age3.HeaderText = "Életkor";

            dataGridView3.Columns.Insert(2, age3);

            DataGridViewCheckBoxColumn checkboxColumnWoman3 =
            new DataGridViewCheckBoxColumn();
            checkboxColumnWoman3.Name = "Woman";
            checkboxColumnWoman3.HeaderText = "Nő";

            dataGridView3.Columns.Insert(3, checkboxColumnWoman3);

            DataGridViewCheckBoxColumn checkboxColumnMan3 =
            new DataGridViewCheckBoxColumn();
            checkboxColumnMan3.Name = "Man";
            checkboxColumnMan3.HeaderText = "Férfi";

            dataGridView3.Columns.Insert(4, checkboxColumnMan3);

            DataGridViewTextBoxColumn ageGroup3 =
            new DataGridViewTextBoxColumn();
            ageGroup3.Name = "AgeGroup";
            ageGroup3.HeaderText = "Korcsoport";

            dataGridView3.Columns.Insert(5, ageGroup3);

            DataGridViewTextBoxColumn km3 =
            new DataGridViewTextBoxColumn();
            km3.Name = "Km";
            km3.HeaderText = "Táv";

            dataGridView3.Columns.Insert(6, km3);

            DataGridViewTextBoxColumn team3 =
            new DataGridViewTextBoxColumn();
            team3.Name = "Team";
            team3.HeaderText = "Csapatnév";

            dataGridView3.Columns.Insert(7, team3);

            DataGridViewTextBoxColumn startNumber3 =
            new DataGridViewTextBoxColumn();
            startNumber3.Name = "StartNumber";
            startNumber3.HeaderText = "Rajtszám";

            dataGridView3.Columns.Insert(8, startNumber3);

            DataGridViewComboBoxColumn comboboxColumnStartTime3 =
            new DataGridViewComboBoxColumn();
            foreach (string str in starttime) { comboboxColumnStartTime3.Items.Add(str); }
            comboboxColumnStartTime3.AutoComplete = true;
            comboboxColumnStartTime3.Name = "Elrajtolási idő";
            comboboxColumnStartTime3.DropDownWidth = 60;
            //comboboxColumnStartTime3.DefaultCellStyle.NullValue = "8:00";

            dataGridView3.Columns.Insert(9, comboboxColumnStartTime3);

            DataGridViewButtonColumn buttonColumnStart3 =
            new DataGridViewButtonColumn();
            buttonColumnStart3.Name = "Start";
            buttonColumnStart3.HeaderText = "";
            buttonColumnStart3.Text = "Start";
            buttonColumnStart3.UseColumnTextForButtonValue = true;

            dataGridView3.Columns.Insert(10, buttonColumnStart3);

            DataGridViewTextBoxColumn startTime3 =
            new DataGridViewTextBoxColumn();
            startTime3.Name = "StartTime";
            startTime3.HeaderText = "Indulás";

            dataGridView3.Columns.Insert(11, startTime3);

            DataGridViewButtonColumn buttonColumnStop3 =
            new DataGridViewButtonColumn();
            buttonColumnStop3.Name = "Stop";
            buttonColumnStop3.HeaderText = "";
            buttonColumnStop3.Text = "Stop";
            buttonColumnStop3.UseColumnTextForButtonValue = true;

            dataGridView3.Columns.Insert(12, buttonColumnStop3);

            DataGridViewTextBoxColumn stopTime3 =
            new DataGridViewTextBoxColumn();
            stopTime3.Name = "StopTime";
            stopTime3.HeaderText = "Érkezés";

            dataGridView3.Columns.Insert(13, stopTime3);

            DataGridViewTextBoxColumn result3 =
            new DataGridViewTextBoxColumn();
            result3.Name = "Result";
            result3.HeaderText = "Időeredmény";

            dataGridView3.Columns.Insert(14, result3);

            DataGridViewButtonColumn buttonColumnNull3 =
            new DataGridViewButtonColumn();
            buttonColumnNull3.Name = "Null";
            buttonColumnNull3.HeaderText = "";
            buttonColumnNull3.Text = "Null";
            buttonColumnNull3.UseColumnTextForButtonValue = true;
            /*buttonColumnNull3.CellTemplate.Style.BackColor = Color.Tomato;
            buttonColumnNull3.CellTemplate.Style.ForeColor = Color.White;
            buttonColumnNull3.FlatStyle = FlatStyle.Flat;*/

            dataGridView3.Columns.Insert(15, buttonColumnNull3);

            //Összegzés DGV --------------------------------------------------------------------------------------------------------------

            DataGridViewTextBoxColumn number4 =
            new DataGridViewTextBoxColumn();
            number4.Name = "Number";
            number4.HeaderText = "Sorszám";

            dataGridView4.Columns.Insert(0, number4);

            DataGridViewTextBoxColumn name4 =
            new DataGridViewTextBoxColumn();
            name4.Name = "Name";
            name4.HeaderText = "Vezetéknév és keresztnév";

            dataGridView4.Columns.Insert(1, name4);

            DataGridViewTextBoxColumn age4 =
            new DataGridViewTextBoxColumn();
            age4.Name = "Age";
            age4.HeaderText = "Életkor";

            dataGridView4.Columns.Insert(2, age4);

            DataGridViewCheckBoxColumn checkboxColumnWoman4 =
            new DataGridViewCheckBoxColumn();
            checkboxColumnWoman4.Name = "Woman";
            checkboxColumnWoman4.HeaderText = "Nő";

            dataGridView4.Columns.Insert(3, checkboxColumnWoman4);

            DataGridViewCheckBoxColumn checkboxColumnMan4 =
            new DataGridViewCheckBoxColumn();
            checkboxColumnMan4.Name = "Man";
            checkboxColumnMan4.HeaderText = "Férfi";

            dataGridView4.Columns.Insert(4, checkboxColumnMan4);

            DataGridViewTextBoxColumn ageGroup4 =
            new DataGridViewTextBoxColumn();
            ageGroup4.Name = "AgeGroup";
            ageGroup4.HeaderText = "Korcsoport";

            dataGridView4.Columns.Insert(5, ageGroup4);

            DataGridViewTextBoxColumn km4 =
            new DataGridViewTextBoxColumn();
            km4.Name = "Km";
            km4.HeaderText = "Táv";

            dataGridView4.Columns.Insert(6, km4);

            DataGridViewTextBoxColumn team4 =
            new DataGridViewTextBoxColumn();
            team4.Name = "Team";
            team4.HeaderText = "Csapatnév";

            dataGridView4.Columns.Insert(7, team4);

            DataGridViewTextBoxColumn startNumber4 =
            new DataGridViewTextBoxColumn();
            startNumber4.Name = "StartNumber";
            startNumber4.HeaderText = "Rajtszám";

            dataGridView4.Columns.Insert(8, startNumber4);

            DataGridViewTextBoxColumn startTimePlan1 =
            new DataGridViewTextBoxColumn();
            startTimePlan1.Name = "StartTimePlan";
            startTimePlan1.HeaderText = "Elrajtolási idő";

            dataGridView4.Columns.Insert(9, startTimePlan1);

            DataGridViewTextBoxColumn start4 =
            new DataGridViewTextBoxColumn();
            start4.Name = "Start";
            start4.HeaderText = "";

            dataGridView4.Columns.Insert(10, start4);

            DataGridViewTextBoxColumn startTime4 =
            new DataGridViewTextBoxColumn();
            startTime4.Name = "StartTime";
            startTime4.HeaderText = "Indulás";

            dataGridView4.Columns.Insert(11, startTime4);

            DataGridViewTextBoxColumn stop4 =
            new DataGridViewTextBoxColumn();
            stop4.Name = "Stop";
            stop4.HeaderText = "";

            dataGridView4.Columns.Insert(12, stop4);

            DataGridViewTextBoxColumn stopTime4 =
            new DataGridViewTextBoxColumn();
            stopTime4.Name = "StopTime";
            stopTime4.HeaderText = "Érkezés";

            dataGridView4.Columns.Insert(13, stopTime4);

            DataGridViewTextBoxColumn result4 =
            new DataGridViewTextBoxColumn();
            result4.Name = "Result";
            result4.HeaderText = "Időeredmény";

            dataGridView4.Columns.Insert(14, result4);

            DataGridViewTextBoxColumn null4 =
            new DataGridViewTextBoxColumn();
            null4.Name = "Null";
            null4.HeaderText = "";

            dataGridView4.Columns.Insert(15, null4);

            dataGridView1.Columns[0].Width = 55;
            dataGridView1.Columns[1].Width = 150;
            dataGridView1.Columns[2].Width = 55;
            dataGridView1.Columns[3].Width = 55;
            dataGridView1.Columns[4].Width = 55;
            dataGridView1.Columns[5].Width = 70;
            dataGridView1.Columns[6].Width = 110;
            dataGridView1.Columns[7].Width = 150;
            dataGridView1.Columns[8].Width = 70;
            dataGridView1.Columns[9].Width = 80;
            dataGridView1.Columns[10].Width = 70;
            dataGridView1.Columns[11].Width = 85;
            dataGridView1.Columns[12].Width = 70;
            dataGridView1.Columns[13].Width = 85;
            dataGridView1.Columns[14].Width = 85;
            dataGridView1.Columns[15].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            dataGridView2.Columns[0].Width = 55;
            dataGridView2.Columns[1].Width = 150;
            dataGridView2.Columns[2].Width = 55;
            dataGridView2.Columns[3].Width = 55;
            dataGridView2.Columns[4].Width = 55;
            dataGridView2.Columns[5].Width = 70;
            dataGridView2.Columns[6].Width = 110;
            dataGridView2.Columns[7].Width = 150;
            dataGridView2.Columns[8].Width = 70;
            dataGridView2.Columns[9].Width = 80;
            dataGridView2.Columns[10].Width = 70;
            dataGridView2.Columns[11].Width = 85;
            dataGridView2.Columns[12].Width = 70;
            dataGridView2.Columns[13].Width = 85;
            dataGridView2.Columns[14].Width = 85;
            dataGridView2.Columns[15].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            dataGridView3.Columns[0].Width = 55;
            dataGridView3.Columns[1].Width = 150;
            dataGridView3.Columns[2].Width = 55;
            dataGridView3.Columns[3].Width = 55;
            dataGridView3.Columns[4].Width = 55;
            dataGridView3.Columns[5].Width = 70;
            dataGridView3.Columns[6].Width = 110;
            dataGridView3.Columns[7].Width = 150;
            dataGridView3.Columns[8].Width = 70;
            dataGridView3.Columns[9].Width = 80;
            dataGridView3.Columns[10].Width = 70;
            dataGridView3.Columns[11].Width = 85;
            dataGridView3.Columns[12].Width = 70;
            dataGridView3.Columns[13].Width = 85;
            dataGridView3.Columns[14].Width = 85;
            dataGridView3.Columns[15].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            dataGridView4.Columns[0].Width = 55;
            dataGridView4.Columns[1].Width = 150;
            dataGridView4.Columns[2].Width = 55;
            dataGridView4.Columns[3].Width = 55;
            dataGridView4.Columns[4].Width = 55;
            dataGridView4.Columns[5].Width = 70;
            dataGridView4.Columns[6].Width = 110;
            dataGridView4.Columns[7].Width = 150;
            dataGridView4.Columns[8].Width = 70;
            dataGridView4.Columns[9].Width = 80;
            dataGridView4.Columns[10].Width = 70;
            dataGridView4.Columns[11].Width = 85;
            dataGridView4.Columns[12].Width = 70;
            dataGridView4.Columns[13].Width = 85;
            dataGridView4.Columns[14].Width = 85;
            dataGridView4.Columns[15].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            dataGridView1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dataGridView2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dataGridView3.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dataGridView4.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            }

            foreach (DataGridViewColumn column in dataGridView2.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            }

            foreach (DataGridViewColumn column in dataGridView3.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            }

            foreach (DataGridViewColumn column in dataGridView4.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            }

            dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

            dataGridView2.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

            dataGridView3.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

            dataGridView4.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

            dataGridView1.Columns[0].ReadOnly = true;
            dataGridView1.Columns[0].DefaultCellStyle.BackColor = Color.LightGray;
            dataGridView1.Columns[5].ReadOnly = true;
            dataGridView1.Columns[5].DefaultCellStyle.BackColor = Color.LightGray;
            dataGridView1.Columns[8].ReadOnly = true;
            dataGridView1.Columns[8].DefaultCellStyle.BackColor = Color.LightGray;
            dataGridView1.Columns[11].ReadOnly = true;
            dataGridView1.Columns[11].DefaultCellStyle.BackColor = Color.LightGray;
            dataGridView1.Columns[13].ReadOnly = true;
            dataGridView1.Columns[13].DefaultCellStyle.BackColor = Color.LightGray;
            dataGridView1.Columns[14].ReadOnly = true;
            dataGridView1.Columns[14].DefaultCellStyle.BackColor = Color.LightGray;

            dataGridView2.Columns[0].ReadOnly = true;
            dataGridView2.Columns[0].DefaultCellStyle.BackColor = Color.LightGray;
            dataGridView2.Columns[5].ReadOnly = true;
            dataGridView2.Columns[5].DefaultCellStyle.BackColor = Color.LightGray;
            dataGridView2.Columns[6].ReadOnly = true;
            dataGridView2.Columns[8].ReadOnly = true;
            dataGridView2.Columns[8].DefaultCellStyle.BackColor = Color.LightGray;
            dataGridView2.Columns[11].ReadOnly = true;
            dataGridView2.Columns[11].DefaultCellStyle.BackColor = Color.LightGray;
            dataGridView2.Columns[13].ReadOnly = true;
            dataGridView2.Columns[13].DefaultCellStyle.BackColor = Color.LightGray;
            dataGridView2.Columns[14].ReadOnly = true;
            dataGridView2.Columns[14].DefaultCellStyle.BackColor = Color.LightGray;

            dataGridView3.Columns[0].ReadOnly = true;
            dataGridView3.Columns[0].DefaultCellStyle.BackColor = Color.LightGray;
            dataGridView3.Columns[5].ReadOnly = true;
            dataGridView3.Columns[5].DefaultCellStyle.BackColor = Color.LightGray;
            dataGridView3.Columns[6].ReadOnly = true;
            dataGridView3.Columns[8].ReadOnly = true;
            dataGridView3.Columns[8].DefaultCellStyle.BackColor = Color.LightGray;
            dataGridView3.Columns[11].ReadOnly = true;
            dataGridView3.Columns[11].DefaultCellStyle.BackColor = Color.LightGray;
            dataGridView3.Columns[13].ReadOnly = true;
            dataGridView3.Columns[13].DefaultCellStyle.BackColor = Color.LightGray;
            dataGridView3.Columns[14].ReadOnly = true;
            dataGridView3.Columns[14].DefaultCellStyle.BackColor = Color.LightGray;

            dataGridView4.Columns[0].ReadOnly = true;
            dataGridView4.Columns[1].ReadOnly = true;
            dataGridView4.Columns[2].ReadOnly = true;
            dataGridView4.Columns[3].ReadOnly = true;
            dataGridView4.Columns[4].ReadOnly = true;
            dataGridView4.Columns[5].ReadOnly = true;
            dataGridView4.Columns[6].ReadOnly = true;
            dataGridView4.Columns[7].ReadOnly = true;
            dataGridView4.Columns[8].ReadOnly = true;
            dataGridView4.Columns[9].ReadOnly = true;
            dataGridView4.Columns[10].ReadOnly = true;
            dataGridView4.Columns[11].ReadOnly = true;
            dataGridView4.Columns[12].ReadOnly = true;
            dataGridView4.Columns[13].ReadOnly = true;
            dataGridView4.Columns[14].ReadOnly = true;
            dataGridView4.Columns[15].ReadOnly = true;
        }

        //Adatok mentése Excel fájlba
        private void Save_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult_Save = MessageBox.Show("Szeretné menteni az adatokat?", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (dialogResult_Save== DialogResult.OK) 
            {
                Cursor.Current = Cursors.WaitCursor;
                Application.UseWaitCursor = true;

                ExcelClass ex = new ExcelClass();

                ex.CreateNewFile();
                ex.CreateNewSheet();
                ex.CreateNewSheet();

                SaveData(dataGridView1, ex);
                ex.ProtectSheet("guranyi");     //Jelszóval lezárás
                ex.SelectWorksheet(2);
                SaveData(dataGridView2, ex);
                ex.ProtectSheet("guranyi");     //Jelszóval lezárás
                ex.SelectWorksheet(3);
                SaveData(dataGridView3, ex);
                ex.ProtectSheet("guranyi");     //Jelszóval lezárás

                try
                {
                    string time = DateTime.Now.ToString("HH:mm:ss");
                    ex.SaveAs(@"VersenyAdatok" + time.Replace(':', '-') + ".xlsx");
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    return;
                }
                
                ex.Close();
                MessageBox.Show("Sikeres mentés", "Információ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else { return; }
            GC.Collect();                   //Excel folyamatok felszabadítása
            GC.WaitForPendingFinalizers();  //Excel folyamatok felszabadítása
            Application.UseWaitCursor = false;
            this.ActiveControl = txtScan.Control;
        }

        private string[] ColumnName = new string[] { "Sorszám", "Vezetéknév és keresztnév", "Életkor", "Nő", "Férfi", "Korcsoport", "Táv", "Csapatnév", "Rajtszám", "Elrajtolási idő", "", "Indulás", "", "Érkezés", "Időeredmény"};

        private void SaveData(DataGridView dgv, ExcelClass e)   //adatmentés eljárás
        {
            for (int j = 0; j < 15; j++)
            {
                e.WriteToExcel(0, j, ColumnName[j]);
            }

            for (int r = 0; r < dgv.Rows.Count - 1; r++)
            {
                for (int col = 0; col < 16; col++)
                {
                    try
                    {
                        e.WriteToExcel(r+1, col, dgv.Rows[r].Cells[col].Value.ToString());
                    }
                    catch (System.NullReferenceException)
                    {
                        e.WriteToExcel(r+1, col, "");
                    }

                }
            }           
        }

        //Adatok betöltése Excel fájlból
        private void Open_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult_Open = MessageBox.Show("Szeretné betölteni az adatokat?", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (dialogResult_Open == DialogResult.OK)
            {
                try
                {
                    dataGridView1.Rows.Clear();
                    dataGridView1.Refresh();
                    dataGridView2.Rows.Clear();
                    dataGridView2.Refresh();
                    dataGridView3.Rows.Clear();
                    dataGridView3.Refresh();
                    dataGridView4.Rows.Clear();
                    dataGridView4.Refresh();

                    var filePath = string.Empty;

                    using (OpenFileDialog openFileDialog = new OpenFileDialog())
                    {
                        openFileDialog.Filter = "Excel Files|*.xls;*.xlsx"; //csak Excel fájlok nyitása

                        if (openFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            filePath = openFileDialog.FileName;
                        }
                    }

                    Cursor.Current = Cursors.WaitCursor;
                    Application.UseWaitCursor = true;

                    ExcelClass excel = new ExcelClass(filePath, 1);

                    OpenData(dataGridView1, dataGridView4, excel);
                    excel.SelectWorksheet(2);
                    OpenData(dataGridView2, dataGridView4, excel);
                    excel.SelectWorksheet(3);
                    OpenData(dataGridView3, dataGridView4, excel);

                    excel.Close();
                    MessageBox.Show("Sikeres adatbetöltés", "Információ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    MessageBox.Show("Nem létező fájl!", "Figyelmeztetés!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else {return;}
            GC.Collect();                   //Excel folyamatok felszabadítása
            GC.WaitForPendingFinalizers();  //Excel folyamatok felszabadítása
            Application.UseWaitCursor = false;
            this.ActiveControl = txtScan.Control;
        }

        private void OpenData(DataGridView dgv, DataGridView dgv4, ExcelClass e)   //adatbetöltés eljárás
        {
            int i = 0;
            int dgv4Rows = dgv4.Rows.Count - 1;
            while (e.ReadCell(i + 1, 1) != "Empty cell")
            {
                dgv.Rows.Add();   //mindig az utolsó sor elé teszi az új sort
                dgv4.Rows.Add();   //mindig az utolsó sor elé teszi az új sort
                //0 -   Sorszám
                dgv4.Rows[dgv4Rows].Cells[0].Value = e.ReadCell(i + 1, 0);
                dgv.Rows[i].Cells[1].Value = e.ReadCell(i + 1, 1);
                dgv4.Rows[dgv4Rows].Cells[1].Value = e.ReadCell(i + 1, 1);
                dgv.Rows[i].Cells[2].Value = e.ReadCell(i + 1, 2);
                dgv4.Rows[dgv4Rows].Cells[2].Value = e.ReadCell(i + 1, 2);
                try
                {
                    dgv.Rows[i].Cells[3].Value = Convert.ToBoolean(e.ReadCell(i + 1, 3));
                    dgv4.Rows[dgv4Rows].Cells[3].Value = Convert.ToBoolean(e.ReadCell(i + 1, 3));
                }
                catch (System.FormatException)
                {
                    dgv.Rows[i].Cells[3].Value = "FALSE";
                    dgv4.Rows[dgv4Rows].Cells[3].Value = "FALSE";
                }
                try
                {
                    dgv.Rows[i].Cells[4].Value = Convert.ToBoolean(e.ReadCell(i + 1, 4));
                    dgv4.Rows[dgv4Rows].Cells[4].Value = Convert.ToBoolean(e.ReadCell(i + 1, 4));
                }
                catch (System.FormatException)
                {
                    dgv.Rows[i].Cells[4].Value = "FALSE";
                    dgv4.Rows[dgv4Rows].Cells[4].Value = "FALSE";
                }
                dgv.Rows[i].Cells[5].Value = e.ReadCell(i + 1, 5);
                dgv4.Rows[dgv4Rows].Cells[5].Value = e.ReadCell(i + 1, 5);
                if (dgv == dataGridView1)
                {
                    if (e.ReadCell(i + 1, 6) == "Classic 20 Km") 
                    {
                        dgv.Rows[i].Cells[6].Value = "Classic 20 Km";
                    }
                    else if (e.ReadCell(i + 1, 6) == "Classic 40 Km")
                    {
                        dgv.Rows[i].Cells[6].Value = "Classic 40 Km";
                    }
                    else if (e.ReadCell(i + 1, 6) == "Family 2 Km")
                    {
                        dgv.Rows[i].Cells[6].Value = "Family 2 Km";
                    }
                    else
                    {
                        dgv.Rows[i].Cells[6].Value = "";
                    }
                }
                else
                {
                    dgv.Rows[i].Cells[6].Value = e.ReadCell(i + 1, 6);
                }
                dgv4.Rows[dgv4Rows].Cells[6].Value = e.ReadCell(i + 1, 6);
                dgv.Rows[i].Cells[7].Value = e.ReadCell(i + 1, 7);
                dgv4.Rows[dgv4Rows].Cells[7].Value = e.ReadCell(i + 1, 7);
                //8 - Rajtszám
                dgv4.Rows[dgv4Rows].Cells[8].Value = e.ReadCell(i + 1, 8);
                try
                {
                    switch (DateTime.FromOADate(Convert.ToDouble(e.ReadCell(i + 1, 9))).ToString("HH:mm"))
                    {
                        case "08:00":
                            dgv.Rows[i].Cells[9].Value = "08:00";
                            break;
                        case "08:15":
                            dgv.Rows[i].Cells[9].Value = "08:15";
                            break;
                        case "08:30":
                            dgv.Rows[i].Cells[9].Value = "08:30";
                            break;
                        case "08:45":
                            dgv.Rows[i].Cells[9].Value = "08:45";
                            break;
                        case "09:00":
                            dgv.Rows[i].Cells[9].Value = "09:00";
                            break;
                        case "09:15":
                            dgv.Rows[i].Cells[9].Value = "09:15";
                            break;
                        case "09:30":
                            dgv.Rows[i].Cells[9].Value = "09:30";
                            break;
                        case "09:45":
                            dgv.Rows[i].Cells[9].Value = "09:45";
                            break;
                        case "10:00":
                            dgv.Rows[i].Cells[9].Value = "10:00";
                            break;
                        case "12:00":
                            dgv.Rows[i].Cells[9].Value = "12:00";
                            break;
                        default:
                            dgv.Rows[i].Cells[9].Value = "";
                            break;
                    }
                }
                catch (System.FormatException)
                {
                    dgv.Rows[i].Cells[9].Value = "";
                }
                dgv4.Rows[dgv4Rows].Cells[9].Value = DateTime.FromOADate(Convert.ToDouble(e.ReadCell(i + 1, 9))).ToString("HH:mm");
                //10    -   Start gomb
                dgv4.Rows[dgv4Rows].Cells[10].Value = "Start";
                try 
                {
                    dgv.Rows[i].Cells[11].Value = DateTime.FromOADate(Convert.ToDouble(e.ReadCell(i + 1, 11))).ToString("HH:mm:ss");
                    dgv4.Rows[dgv4Rows].Cells[11].Value = DateTime.FromOADate(Convert.ToDouble(e.ReadCell(i + 1, 11))).ToString("HH:mm:ss");
                }
                catch (System.FormatException)
                {
                    dgv.Rows[i].Cells[11].Value = "";
                    dgv4.Rows[dgv4Rows].Cells[11].Value = "";
                }
                dgv4.Rows[dgv4Rows].Cells[12].Value = "Stop";
                //12    -   Stop gomb
                try
                {
                    dgv.Rows[i].Cells[13].Value = DateTime.FromOADate(Convert.ToDouble(e.ReadCell(i + 1, 13))).ToString("HH:mm:ss");
                    dgv4.Rows[dgv4Rows].Cells[13].Value = DateTime.FromOADate(Convert.ToDouble(e.ReadCell(i + 1, 13))).ToString("HH:mm:ss");
                }
                catch (System.FormatException)
                {
                    dgv.Rows[i].Cells[13].Value = "";
                    dgv4.Rows[dgv4Rows].Cells[13].Value = "";
                }
                try 
                {
                    dgv.Rows[i].Cells[14].Value = DateTime.FromOADate(Convert.ToDouble(e.ReadCell(i + 1, 14))).ToString("HH:mm:ss");
                    dgv4.Rows[dgv4Rows].Cells[14].Value = DateTime.FromOADate(Convert.ToDouble(e.ReadCell(i + 1, 14))).ToString("HH:mm:ss");
                }
                catch (System.FormatException)
                {
                    dgv.Rows[i].Cells[14].Value = "";
                    dgv4.Rows[dgv4Rows].Cells[14].Value = "";
                }
                //15    -   Null gomb
                i++;
                dgv4Rows++;
            }
            // Helyezések beállítása színekkel
            colorResults(dgv);
        }

        //A kiválasztott adattáblán (Classic, MTB 25, ROAD 70) rendez időeredmény szerint
        private void Sort_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])
            {
                dataGridView1.Sort(dataGridView1.Columns["Result"], ListSortDirection.Ascending);
                MessageBox.Show("A rendezés megtörtént", "Információ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"])
            {
                dataGridView2.Sort(dataGridView2.Columns["Result"], ListSortDirection.Ascending);
                MessageBox.Show("A rendezés megtörtént", "Információ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            } 
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage3"])
            {
                dataGridView3.Sort(dataGridView3.Columns["Result"], ListSortDirection.Ascending);
                MessageBox.Show("A rendezés megtörtént", "Információ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            } 
        }

        private void txtScan_KeyUp(object sender, KeyEventArgs e)
        {
            string startNumber = "";
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    startNumber = txtScan.Text;
                    if (startNumber.Substring(0, 1) == "C")
                    {
                        timeHandling(dataGridView1, startNumber);
                    }
                    else if (startNumber.Substring(0, 1) == "M")
                    {
                        timeHandling(dataGridView2, startNumber);
                    }
                    else if (startNumber.Substring(0, 1) == "R")
                    {
                        timeHandling(dataGridView3, startNumber);
                    }
                    else
                    {
                        MessageBox.Show("Nem érvényes sorszám", "Információ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    txtScan.Text = String.Empty;
                }
                catch (System.ArgumentOutOfRangeException)
                {
                    this.ActiveControl = txtScan.Control;
                }
            }
        }

        private void timeHandling(DataGridView dgv, string startnumber)
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                try
                {
                    if (row.Cells[8].Value.ToString() == startnumber.Substring(1))
                    {
                        if (row.Cells[11].Value == null || row.Cells[11].Value.ToString() == "")
                        {
                            row.Cells[11].Value = DateTime.Now.ToString("HH:mm:ss");
                            break;
                        }
                        else if (row.Cells[13].Value == null || row.Cells[13].Value.ToString() == "")
                        {
                            stop = DateTime.Now;
                            TimeSpan result = stop - DateTime.Parse(row.Cells[11].Value.ToString());
                            if (TimeSpan.Compare(result, elteltido) == 1)
                            {
                                row.Cells[13].Value = stop.ToString("HH:mm:ss");
                                row.Cells[14].Value = result.ToString(@"hh\:mm\:ss");

                                // Helyezések beállítása színekkel
                                colorResults(dgv);
                            }
                            break;
                        }
                        else
                        { break; }
                    }
                }
                catch (System.Exception ex)
                {
                    if (ex is System.NullReferenceException || ex is System.ArgumentOutOfRangeException)
                    {
                        break;
                    }
                }
            }
        }

        private void colorResults(DataGridView dgv) // Helyezések beállítása színekkel
        {
            string[] array = new string[dgv.Rows.Count - 1];
            string[] sortedArray = new string[dgv.Rows.Count - 1];
            for (short i = 0; i < array.Length; i++)
            {
                if (dgv.Rows[i].Cells[14].Value == null || dgv.Rows[i].Cells[14].Value.ToString() == "")
                {
                    array[i] = "99:99:99";
                    sortedArray[i] = "99:99:99";
                }
                else
                {
                    array[i] = dgv.Rows[i].Cells[14].Value.ToString();
                    sortedArray[i] = dgv.Rows[i].Cells[14].Value.ToString();
                }
            }
            Array.Sort(sortedArray);
            string[] uniqueArray = sortedArray.Distinct().ToArray(); //kiveszi a duplikált elemeket
            for (short i = 0; i < array.Length; i++)
            {
                if (uniqueArray[0] == "99:99:99")
                {
                    break;
                }
                else if (String.Compare(array[i], uniqueArray[0]) == 0)
                {
                    dgv.Rows[i].DefaultCellStyle.BackColor = Color.Gold;
                }
                else if (String.Compare(array[i], uniqueArray[1]) == 0 && uniqueArray[1] != "99:99:99")
                {
                    dgv.Rows[i].DefaultCellStyle.BackColor = Color.Gray;
                }
                else if (uniqueArray.Length >= 3)
                {
                    if (String.Compare(array[i], uniqueArray[2]) == 0 && uniqueArray[2] != "99:99:99")
                    {
                        dgv.Rows[i].DefaultCellStyle.BackColor = Color.SaddleBrown;
                    }
                    else
                    {
                        dgv.Rows[i].DefaultCellStyle.BackColor = Color.Empty;
                    }
                }
                else
                {
                    dgv.Rows[i].DefaultCellStyle.BackColor = Color.Empty;
                }
            }
        }
    }
}