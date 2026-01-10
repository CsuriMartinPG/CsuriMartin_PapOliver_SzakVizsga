using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace HealthTrack_CsuriMartin_PapOliver
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string kivalasztottOsztaly;
        string kivalasztottNev;
        List<string> osztalyok = new List<string>();
        List<string> diakok = new List<string>();

        public void Form1_Load(object sender, EventArgs e)
        {
            var serverKapcsolat = new MySqlConnectionStringBuilder { Server = "127.0.0.1", Database = "healthtrack", UserID = "root", Password = "mysql" };
            MySqlConnection kapcsolat;
            kapcsolat = new MySqlConnection(serverKapcsolat.ConnectionString);
            kapcsolat.Open();
            var lekerdezes = kapcsolat.CreateCommand();
            lekerdezes.CommandText = "SELECT nev FROM meresek";

            var olvaso = lekerdezes.ExecuteReader();
            MessageBox.Show("ADATBÁZIS BETÖLTVE");
            olvaso.Close();
            

            lekerdezes.CommandText = $"SELECT DISTINCT osztaly\r\n FROM meresek\r\n";
            olvaso = lekerdezes.ExecuteReader();
            while (olvaso.Read())
            {
                osztalyok.Add(olvaso["osztaly"].ToString());
            }
            olvaso.Close();

            foreach (var item in osztalyok)
            {
                osztalyLista.Items.Add(item);
            }



            osztalyLista.Visible = false;
            diakLista.Visible = false;
            vissza.Visible = false;
            elvalaszto1.Visible = false;
            lekerdezesPanel.Visible = false;
        }

        private void kilepes_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void foLekerdezes_Click(object sender, EventArgs e)
        {
            foLekerdezes.Visible = false;
            foHozzaadas.Visible = false;
            foTorles.Visible = false;
            osztalyLista.Visible = true;
            vissza.Visible = true;

        }

        private void osztalyLista_SelectedIndexChanged(object sender, EventArgs e)
        {
            diakLista.Visible = true;
            kivalasztottOsztaly = osztalyLista.Text;
            diakok.Clear();
            diakLista.Items.Clear();
            lekerdezesPanel.Visible = false;

            var serverKapcsolat = new MySqlConnectionStringBuilder { Server = "127.0.0.1", Database = "healthtrack", UserID = "root", Password = "mysql" };
            MySqlConnection kapcsolat;
            kapcsolat = new MySqlConnection(serverKapcsolat.ConnectionString);
            kapcsolat.Open();
            var lekerdezes = kapcsolat.CreateCommand();
            lekerdezes.CommandText = $"SELECT nev FROM szemelyek";

            var olvaso = lekerdezes.ExecuteReader();
            olvaso.Close();

            lekerdezes.CommandText = $"SELECT nev\r\n FROM meresek\r\nWHERE osztaly = '{kivalasztottOsztaly}'\r\n ORDER BY nev DESC";
            olvaso = lekerdezes.ExecuteReader();
            while (olvaso.Read())
            {
                diakok.Add(olvaso["nev"].ToString());
            }
            olvaso.Close();

            foreach (var item in diakok)
            {
                diakLista.Items.Add(item);
            }
        }

        private void diakLista_SelectedIndexChanged(object sender, EventArgs e)
        {
            kivalasztottNev = diakLista.Text;
            lekerdezesPanel.Visible = true;

            var serverKapcsolat = new MySqlConnectionStringBuilder { Server = "127.0.0.1", Database = "healthtrack", UserID = "root", Password = "mysql" };
            MySqlConnection kapcsolat;
            kapcsolat = new MySqlConnection(serverKapcsolat.ConnectionString);
            kapcsolat.Open();
            var lekerdezes = kapcsolat.CreateCommand();
            lekerdezes.CommandText = $"SELECT nev FROM szemelyek";

            var olvaso = lekerdezes.ExecuteReader();
            olvaso.Close();

            lekerdezes.CommandText = $"SELECT *\r\n FROM meresek\r\n WHERE osztaly = '{kivalasztottOsztaly}' AND nev = '{kivalasztottNev}'";
            olvaso = lekerdezes.ExecuteReader();
            while (olvaso.Read())
            {
                oktAz.Text = olvaso.GetString(1);
                eletkor.Text = olvaso.GetInt32(2).ToString();
                suly.Text = olvaso.GetDouble(3).ToString();
                magassag.Text = olvaso.GetInt32(4).ToString();
                osztaly.Text = kivalasztottOsztaly;
                nev.Text = kivalasztottNev;
            }
            olvaso.Close();

        }

        private void vissza_Click(object sender, EventArgs e)
        {
            vissza.Visible = false;
            foLekerdezes.Visible = true;
            foHozzaadas.Visible = true;
            foTorles.Visible = true;
            osztalyLista.Visible = false;
            diakLista.Visible = false;
            elvalaszto1.Visible = false;
            lekerdezesPanel.Visible = false;
        }
    }
}