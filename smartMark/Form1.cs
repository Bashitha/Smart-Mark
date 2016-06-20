using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using MySql.Data.MySqlClient;
using GmailSend;

using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace smartMark
{
    public partial class Form1 : Form
    {
        private static DataTable dbdataset = null;
        public Form1()
        {
            InitializeComponent();
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void butStudentSearch_Click(object sender, EventArgs e)
        {
            string connString = "Server = localhost ; Port = 3306 ; Database = smartMark ; Uid = root ; password = bashitha; ";
            MySqlConnection conn = new MySqlConnection(connString);
            MySqlCommand command = conn.CreateCommand();
            command.CommandText = " SELECT course_code,SUM(lec_hours*attendance*-1) AS total_absent_lectures,total_lec_hours FROM mark_attendance NATURAL JOIN course WHERE reg_no = '"+txtRegNo.Text+"' AND attendance = -1 GROUP BY 1;";

            try
            {
                conn.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            MySqlDataAdapter sda = new MySqlDataAdapter();
            sda.SelectCommand = command;
            DataTable dbdataset = new DataTable();
            sda.Fill(dbdataset);

            DataColumn percentageColumn = new DataColumn("Absent percentage(%)", typeof(int));
            percentageColumn.DefaultValue = 0;
            dbdataset.Columns.Add(percentageColumn);
          
            for (int i = 0; i < dbdataset.Rows.Count; i++)
            {
               dbdataset.Rows[i][3]=(double.Parse( dbdataset.Rows[i][1].ToString())/double.Parse(dbdataset.Rows[i][2].ToString() ))*100;
               

            }

            BindingSource bs = new BindingSource();

            bs.DataSource = dbdataset;
            dataGridView1.DataSource = bs;

            sendingMail(command, dbdataset);
           
            conn.Close();
        }

        private void butCourseSearch_Click(object sender, EventArgs e)
        {
            string connString = "Server = localhost ; Port = 3306 ; Database = smartMark ; Uid = root ; password = bashitha; ";
            MySqlConnection conn = new MySqlConnection(connString);
            MySqlCommand command = conn.CreateCommand();
            command.CommandText = " SELECT reg_no,SUM(lec_hours*attendance*-1) AS total_absent_lectures,total_lec_hours FROM mark_attendance NATURAL JOIN course WHERE course_code = '" + txtCourseCode.Text + "' AND attendance = -1 GROUP BY 1;";

            try
            {
                conn.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            MySqlDataAdapter sda = new MySqlDataAdapter();
            sda.SelectCommand = command;
            dbdataset = new DataTable();//table
            sda.Fill(dbdataset);

            DataColumn percentageColumn = new DataColumn("Absent percentage(%)", typeof(int));
            percentageColumn.DefaultValue = 0;
            dbdataset.Columns.Add(percentageColumn);

            for (int i = 0; i < dbdataset.Rows.Count; i++)
            {
                dbdataset.Rows[i][3] = (double.Parse(dbdataset.Rows[i][1].ToString()) / double.Parse(dbdataset.Rows[i][2].ToString())) * 100;


            }

            
            BindingSource bs = new BindingSource();

            bs.DataSource = dbdataset;
            dataGridView1.DataSource = bs;
            
            conn.Close();
        }

        private void butStudentEmail_Click(object sender, EventArgs e)
        {
            
           
        }

        public void ExportToPdf(DataTable dt)
        {
            Document document = new Document();
            PdfWriter writer = PdfWriter.GetInstance(document, new FileStream("c://sample.pdf", FileMode.Create));
            document.Open();
            iTextSharp.text.Font font5 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 5);

            PdfPTable table = new PdfPTable(dt.Columns.Count);
            
            float[] widths = new float[] { 4f, 4f, 4f, 4f };

            table.SetWidths(widths);

            table.WidthPercentage = 100;
            
            
            PdfPCell cell = new PdfPCell(new Phrase("Products"));

            cell.Colspan = dt.Columns.Count;

            foreach (DataColumn c in dt.Columns)
            {

                table.AddCell(new Phrase(c.ColumnName, font5));
            }

            foreach (DataRow r in dt.Rows)
            {
                if (dt.Rows.Count > 0)
                {
                    table.AddCell(new Phrase(r[0].ToString(), font5));
                    table.AddCell(new Phrase(r[1].ToString(), font5));
                    table.AddCell(new Phrase(r[2].ToString(), font5));
                    table.AddCell(new Phrase(r[3].ToString(), font5));
                }
            } document.Add(table);
            document.Close();
        }
        private  void sendingMail(MySqlCommand command, DataTable dbdataset)
        {
            gmail gmlsnd = new gmail();
            gmlsnd.auth("bashithawije@gmail.com", "1amgenius");

            gmlsnd.Subject = "Low Attendance";

            gmlsnd.Priority = 1;

            for (int i = 0; i < dbdataset.Rows.Count; i++)
            {
                if (int.Parse(dbdataset.Rows[i][3].ToString()) >= 15)
                {
                    //taking email of appropriate student from the database
                    String email;
                   
                    command.CommandText = " SELECT email FROM student WHERE reg_no=" + txtRegNo.Text;

                    
                    MySqlDataReader reader = command.ExecuteReader();
                    reader.Read();
                    email = reader["email"].ToString();

                    Console.Write(email);

                    //sending email
                    gmlsnd.To = email;
                    gmlsnd.Message = "Your attendance for " + dbdataset.Rows[i][0].ToString() + " is nearing the minimum faculty requirement.Please pay your attention for that.";

                    try
                    {
                        gmlsnd.send();
                        MessageBox.Show("Your Mail is sent");
                        reader.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                
            }
        }
    }
}
