using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using LiveCharts;
using LiveCharts.Wpf;

namespace Tyuiu.MezentsevSE.Project.V6
{
    public partial class FormMain : Form
    {
        private SqlConnection sqlConnection = null;

        private SqlDataAdapter dataAdapter = null;

        private DataSet dataSet = null;

        private SqlCommandBuilder sqlBuilder = null;

        private bool newRowAdding = false;
        public FormMain()
        {
            InitializeComponent();
        }


        private void LoadData()
        {
            try
            {
                dataAdapter = new SqlDataAdapter("Select *, 'Delete' AS [Command] FROM Patients", sqlConnection);

                sqlBuilder = new SqlCommandBuilder(dataAdapter);

                sqlBuilder.GetInsertCommand();
                sqlBuilder.GetUpdateCommand();
                sqlBuilder.GetDeleteCommand();

                dataSet = new DataSet();

                dataAdapter.Fill(dataSet, "Patients");
                dataGridViewEdit_MSE.DataSource = dataSet.Tables["Patients"];

                for (int i = 0; i< dataGridViewEdit_MSE.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridViewEdit_MSE[15, i] = linkCell;
                }



                


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        private void ReloadData()
        {
            try
            {
                dataSet.Tables["Patients"].Clear();

                dataAdapter.Fill(dataSet, "Patients");
                dataGridViewEdit_MSE.DataSource = dataSet.Tables["Patients"];

                for (int i = 0; i < dataGridViewEdit_MSE.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridViewEdit_MSE[15, i] = linkCell;
                }






            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }




        private void FormMain_Load(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["Patients"].ConnectionString);

            sqlConnection.Open();

            SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT * FROM Patients", sqlConnection);

            DataSet database = new DataSet();

            dataAdapter.Fill(database, "Patients");

            

            dataGridViewFilter_MSE.DataSource = database.Tables[0];

            LoadData();
            




        }

        private void buttonInsert_MSE_Click(object sender, EventArgs e)
        {
            SqlCommand command = new SqlCommand("INSERT INTO [Patients] (Num,Surname,Name,Otchestvo,Data_Rozhdeniya,SurnameD,NameD,OtchestvoD,Dolzhnost,Ill,Heal,Time,Dispanser,Info) VALUES (@Num,@Surname,@Name,@Otchestvo,@Data_Rozhdeniya,@SurnameD,@NameD,@OtchestvoD,@Dolzhnost,@Ill,@Heal,@Time,@Dispanser,@Info)", sqlConnection);

            DateTime date = DateTime.Parse(textBoxDate_MSE.Text);

            command.Parameters.AddWithValue("Num", textBoxNum_MSE.Text);
            command.Parameters.AddWithValue("Surname", textBoxSurname_MSE.Text);
            command.Parameters.AddWithValue("Name", textBoxName_MSE.Text);
            command.Parameters.AddWithValue("Otchestvo", textBoxOtchestvo_MSE.Text);
            command.Parameters.AddWithValue("Data_Rozhdeniya", $"{date.Month}.{date.Day}.{date.Year}");
            command.Parameters.AddWithValue("SurnameD", textBoxSurnameD_MSE.Text);
            command.Parameters.AddWithValue("NameD", textBoxNameD_MSE.Text);
            command.Parameters.AddWithValue("OtchestvoD", textBoxOtchestvoD_MSE.Text);
            command.Parameters.AddWithValue("Dolzhnost", textBoxWork_MSE.Text);
            command.Parameters.AddWithValue("Ill", textBoxIll_MSE.Text);
            command.Parameters.AddWithValue("Heal", textBoxHeal_MSE.Text);
            command.Parameters.AddWithValue("Time", textBoxTime_MSE.Text);
            command.Parameters.AddWithValue("Dispanser", textBoxDisp_MSE.Text);
            command.Parameters.AddWithValue("Info", textBoxMoreInfo_MSE.Text);

            MessageBox.Show("Количество строк введено:", command.ExecuteNonQuery().ToString());
        }

        private void buttonSelect_MSE_Click(object sender, EventArgs e)
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter(textBoxSelect_MSE.Text, sqlConnection);

            DataSet dataSet = new DataSet();
            dataAdapter.Fill(dataSet);

            dataGridViewTabl_MSE.DataSource = dataSet.Tables[0];
        }

        private void comboBoxFilter_MSE_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBoxFilter_MSE.SelectedIndex)
            {
                case 0:
                    (dataGridViewFilter_MSE.DataSource as DataTable).DefaultView.RowFilter = $"TIME = 'Хроническое заболевание'";

                    break;
                case 1:
                    (dataGridViewFilter_MSE.DataSource as DataTable).DefaultView.RowFilter = $"TIME = 'Месяц'";

                    break;
                case 2:
                    (dataGridViewFilter_MSE.DataSource as DataTable).DefaultView.RowFilter = $"TIME = 'Полмесяца'";

                    break;
                case 3:
                    (dataGridViewFilter_MSE.DataSource as DataTable).DefaultView.RowFilter = $"TIME = 'Пара дней'";

                    break;
                case 4:
                    (dataGridViewFilter_MSE.DataSource as DataTable).DefaultView.RowFilter = "";

                    break;

            }
        }


        /*private void Search(DataGridView db)
        {
            db.Rows.Clear();
            string search = $"SELECT * FROM Patients WHERE (Num,Surname,Name,Otchestvo,Data_Rozhdeniya,SurnameD,NameD,OtchestvoD,Dolzhnost,Ill,Heal,Time,Dispanser,Info) LIKE '%" + textBoxFilter_MSE.Text + "%'";

           

            SqlCommand command = new SqlCommand(search, sqlConnection );

            sqlConnection.Open();

            SqlDataReader read = command.ExecuteReader();

            while (read.Read())
            {
                ReadSingleRow(db, search);
            }
            read.Close();

        }
        */


        

        private void textBoxFilter_MSE_TextChanged(object sender, EventArgs e)
        {
            (dataGridViewFilter_MSE.DataSource as DataTable).DefaultView.RowFilter = $"Ill LIKE '%{textBoxFilter_MSE.Text}%' OR Surname LIKE '%{textBoxFilter_MSE.Text}%' OR Name LIKE '%{textBoxFilter_MSE.Text}%' OR Otchestvo LIKE '%{textBoxFilter_MSE.Text}%' OR SurnameD LIKE '%{textBoxFilter_MSE.Text}%' OR NameD LIKE '%{textBoxFilter_MSE.Text}%' OR OtchestvoD LIKE '%{textBoxFilter_MSE.Text}%' OR Dolzhnost LIKE '%{textBoxFilter_MSE.Text}%' OR Heal LIKE '%{textBoxFilter_MSE.Text}%' OR Time LIKE '%{textBoxFilter_MSE.Text}%' OR Dispanser LIKE '%{textBoxFilter_MSE.Text}%' OR Info LIKE '%{textBoxFilter_MSE.Text}%'";

            

        }

        private void toolStripButtonGraph_MSE_Click(object sender, EventArgs e)
        {
            dataAdapter.Fill(dataSet, "Patients");

            SeriesCollection series = new SeriesCollection();

            ChartValues<int> num = new ChartValues<int>();

            List<string> birthdate = new List<string>();

           

            
        }

        private void toolStripButtonEdit_MSE_Click(object sender, EventArgs e)
        {
            ReloadData();
        }

        private void dataGridViewEdit_MSE_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 15)
                {
                    string P = dataGridViewEdit_MSE.Rows[e.RowIndex].Cells[15].Value.ToString();

                    if (P == "Delete")
                    {
                        if (MessageBox.Show("Удалить эту строку?", "Удаление" , MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            int rowIndex = e.RowIndex;

                            dataGridViewEdit_MSE.Rows.RemoveAt(rowIndex);

                            dataSet.Tables["Patients"].Rows[rowIndex].Delete();

                            dataAdapter.Update(dataSet,"Patients");

                        }
                    }
                    else if (P == "Insert")
                    {
                        int rowIndex = dataGridViewEdit_MSE.Rows.Count - 2;

                        DataRow row = dataSet.Tables["Patients"].NewRow();

                        row["Num"] = dataGridViewEdit_MSE.Rows[rowIndex].Cells["Num"].Value;
                        row["Surname"] = dataGridViewEdit_MSE.Rows[rowIndex].Cells["Surname"].Value;
                        row["Name"] = dataGridViewEdit_MSE.Rows[rowIndex].Cells["Name"].Value;
                        row["Otchestvo"] = dataGridViewEdit_MSE.Rows[rowIndex].Cells["Otchestvo"].Value;
                        row["Data_Rozhdeniya"] = dataGridViewEdit_MSE.Rows[rowIndex].Cells["Data_Rozhdeniya"].Value;
                        row["SurnameD"] = dataGridViewEdit_MSE.Rows[rowIndex].Cells["SurnameD"].Value;
                        row["NameD"] = dataGridViewEdit_MSE.Rows[rowIndex].Cells["NameD"].Value;
                        row["OtchestvoD"] = dataGridViewEdit_MSE.Rows[rowIndex].Cells["OtchestvoD"].Value;
                        row["Dolzhnost"] = dataGridViewEdit_MSE.Rows[rowIndex].Cells["Dolzhnost"].Value;
                        row["Ill"] = dataGridViewEdit_MSE.Rows[rowIndex].Cells["Ill"].Value;
                        row["Heal"] = dataGridViewEdit_MSE.Rows[rowIndex].Cells["Heal"].Value;
                        row["Time"] = dataGridViewEdit_MSE.Rows[rowIndex].Cells["Time"].Value;
                        row["Dispanser"] = dataGridViewEdit_MSE.Rows[rowIndex].Cells["Dispanser"].Value;
                        row["Info"] = dataGridViewEdit_MSE.Rows[rowIndex].Cells["Info"].Value;

                        dataSet.Tables["Patients"].Rows.Add(row);

                        dataSet.Tables["Patients"].Rows.RemoveAt(dataSet.Tables["Patients"].Rows.Count - 1);

                        dataGridViewEdit_MSE.Rows.RemoveAt(dataGridViewEdit_MSE.Rows.Count - 2);
                        dataGridViewEdit_MSE.Rows[e.RowIndex].Cells[15].Value = "Delete";

                        dataAdapter.Update(dataSet, "Patients");

                        newRowAdding = false;

                    }
                    else if (P == "Update")
                    {
                        int r = e.RowIndex;


                        dataSet.Tables["Patients"].Rows[r]["Num"] = dataGridViewEdit_MSE.Rows[r].Cells["Num"].Value;
                        dataSet.Tables["Patients"].Rows[r]["Surname"] = dataGridViewEdit_MSE.Rows[r].Cells["Surname"].Value;
                        dataSet.Tables["Patients"].Rows[r]["Name"] = dataGridViewEdit_MSE.Rows[r].Cells["Name"].Value;
                        dataSet.Tables["Patients"].Rows[r]["Otchestvo"] = dataGridViewEdit_MSE.Rows[r].Cells["Otchestvo"].Value;
                        dataSet.Tables["Patients"].Rows[r]["Data_Rozhdeniya"] = dataGridViewEdit_MSE.Rows[r].Cells["Data_Rozhdeniya"].Value;
                        dataSet.Tables["Patients"].Rows[r]["SurnameD"] = dataGridViewEdit_MSE.Rows[r].Cells["SurnameD"].Value;
                        dataSet.Tables["Patients"].Rows[r]["NameD"] = dataGridViewEdit_MSE.Rows[r].Cells["NameD"].Value;
                      
                        dataSet.Tables["Patients"].Rows[r]["OtchestvoD"] = dataGridViewEdit_MSE.Rows[r].Cells["OtchestvoD"].Value;
                        dataSet.Tables["Patients"].Rows[r]["Dolzhnost"] = dataGridViewEdit_MSE.Rows[r].Cells["Dolzhnost"].Value;
                        dataSet.Tables["Patients"].Rows[r]["Ill"] = dataGridViewEdit_MSE.Rows[r].Cells["Ill"].Value;
                        dataSet.Tables["Patients"].Rows[r]["Heal"] = dataGridViewEdit_MSE.Rows[r].Cells["Heal"].Value;
                        dataSet.Tables["Patients"].Rows[r]["Time"] = dataGridViewEdit_MSE.Rows[r].Cells["Time"].Value;
                        dataSet.Tables["Patients"].Rows[r]["Dispanser"] = dataGridViewEdit_MSE.Rows[r].Cells["Dispanser"].Value;
                        dataSet.Tables["Patients"].Rows[r]["Info"] = dataGridViewEdit_MSE.Rows[r].Cells["Info"].Value;

                        dataAdapter.Update(dataSet, "Patients");
                        dataGridViewEdit_MSE.Rows[e.RowIndex].Cells[15].Value = "Delete";


                    }

                    ReloadData();
                
                
                
                }






            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void dataGridViewEdit_MSE_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            try
            {
                if (newRowAdding == false)
                {
                    newRowAdding = true;
                    int lastRow = dataGridViewEdit_MSE.Rows.Count - 2;
                    DataGridViewRow row = dataGridViewEdit_MSE.Rows[lastRow];

                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridViewEdit_MSE[15, lastRow] = linkCell;

                    row.Cells["Command"].Value = "Insert";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridViewEdit_MSE_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if(newRowAdding == false)
                {
                    int rowIndex = dataGridViewEdit_MSE.SelectedCells[0].RowIndex;

                    DataGridViewRow editingRow = dataGridViewEdit_MSE.Rows[rowIndex];

                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridViewEdit_MSE[15, rowIndex] = linkCell;

                    editingRow.Cells["Command"].Value = "Update";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void dataGridViewEdit_MSE_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.KeyPress -= new KeyPressEventHandler(Column_KeyPress);

            if (dataGridViewEdit_MSE.CurrentCell.ColumnIndex == 5)
            {
                TextBox textbox = e.Control as TextBox;

                if(textbox != null)
                {
                    textbox.KeyPress += new KeyPressEventHandler(Column_KeyPress);
                }
            }
        }

        private void Column_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }
    }
}
