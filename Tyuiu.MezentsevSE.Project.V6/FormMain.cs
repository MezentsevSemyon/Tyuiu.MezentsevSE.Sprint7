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

namespace Tyuiu.MezentsevSE.Project.V6
{
    public partial class FormMain : Form
    {
        private SqlConnection sqlConnection = null;
        public FormMain()
        {
            InitializeComponent();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["Patients"].ConnectionString);

            sqlConnection.Open();

            SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT * FROM Patients", sqlConnection);

            DataSet database = new DataSet();

            dataAdapter.Fill(database);

            dataGridViewFilter_MSE.DataSource = database.Tables[0];


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

       

        
    }
}
