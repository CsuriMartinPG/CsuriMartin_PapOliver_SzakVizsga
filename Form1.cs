using MySql.Data.MySqlClient;
using Mysqlx.Crud;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
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
        string neme;
        double sulySz = 0.0;
        double magassagSz = 0.0;
        double bmi = 0.0;
        double testzsirSzazalek = 0.0;
        int eletkora = 0;
        List<string> osztalyok = new List<string>();
        List<string> diakok = new List<string>();
        List<kategoriaBMI> fileadat = new List<kategoriaBMI>();

        static public double testZsirSzazalekSzamitas(string neme, int eletkor, double bmi)
        {
            
            if(neme == "férfi")
            {
                return Math.Round((1.20 * bmi) + (0.23 * eletkor) - 16.2, 2);
            }
            else
            {
                return Math.Round((1.20 * bmi) + (0.23 * eletkor) - 5.4, 2);
            }
        }

        public void Form1_Load(object sender, EventArgs e)
        {
            var serverKapcsolat = new MySqlConnectionStringBuilder { Server = "127.0.0.1", Database = "healthtrack", UserID = "root", Password = "mysql" };
            MySqlConnection kapcsolat;
            kapcsolat = new MySqlConnection(serverKapcsolat.ConnectionString);
            kapcsolat.Open();
            var lekerdezes = kapcsolat.CreateCommand();
            lekerdezes.CommandText = "SELECT nev FROM meresek";

            var olvaso = lekerdezes.ExecuteReader();
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
                osztalyLista2.Items.Add(item);
            }

            StreamReader sr = new StreamReader("bmiKategoria.txt");

            while(!sr.EndOfStream)
            {
                fileadat.Add(new kategoriaBMI(sr.ReadLine()));
            }
            fileadat.Reverse();
            sr.Close();
            
            osztalyLista.Visible = false;
            diakLista.Visible = false;
            vissza.Visible = false;
            elvalaszto1.Visible = false;
            lekerdezesPanel.Visible = false;
            osztalyLista2.Visible = false;
            diakLista2.Visible = false;
        }

        private void kilepes_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void foLekerdezes_Click(object sender, EventArgs e)
        {
            var serverKapcsolat = new MySqlConnectionStringBuilder { Server = "127.0.0.1", Database = "healthtrack", UserID = "root", Password = "mysql" };
            MySqlConnection kapcsolat;
            kapcsolat = new MySqlConnection(serverKapcsolat.ConnectionString);
            kapcsolat.Open();
            var lekerdezes = kapcsolat.CreateCommand();
            lekerdezes.CommandText = "SELECT nev FROM meresek";

            var olvaso = lekerdezes.ExecuteReader();
            olvaso.Close();

            foLekerdezes.Visible = false;
            foHozzaadas.Visible = false;
            foTorles.Visible = false;
            osztalyLista.Visible = true;
            vissza.Visible = true;
            lekerdezesPanel.Visible = true;

            nevTxb.Visible = false;
            osztalyTxb.Visible = false;
            oktAzTxb.Visible = false;
            eletkorTxb.Visible = false;
            nemTxb.Visible = false;
            sulyTxb.Visible = false;
            magassagTxb.Visible = false;
            hozzaadas.Visible = false;
            torles.Visible = false;
            testZsirLabel.Visible = true;
            testzsir.Visible = true;
            bmiErtekLabel.Visible = true;
            bmiErtek.Visible = true;
            bmiKatLabel.Visible = true;
            bmiKat.Visible = true;
            tzsirAlso.Visible = true;
            tzsirFelso.Visible = true;
            tzsirKozep.Visible = true;
            pictureBox2.Visible = true;
            tzsirKat.Visible = true;
            tzsirKatLabel.Visible = true;
        }

        private void osztalyLista_SelectedIndexChanged(object sender, EventArgs e)
        {
            diakLista.Visible = true;
            kivalasztottOsztaly = osztalyLista.Text;
            diakok.Clear();
            diakLista.Items.Clear();
            nev.Text = "-----";
            osztaly.Text = "-----";
            oktAz.Text = "-----";
            eletkor.Text = "-----";
            nem.Text = "-----";
            suly.Text = "-----";
            magassag.Text = "-----";
            bmiErtek.Text = "-----";
            bmiKat.Text = "-----";

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
                sulySz = olvaso.GetDouble(3);
                magassag.Text = olvaso.GetInt32(4).ToString();
                magassagSz = olvaso.GetDouble(4)/100;
                osztaly.Text = kivalasztottOsztaly;
                nev.Text = kivalasztottNev;
                eletkora = int.Parse(eletkor.Text);
                nem.Text = olvaso.GetString(7);
                bmi = Math.Round(sulySz / (magassagSz * magassagSz), 1);
                bmiErtek.Text = bmi.ToString();
                
                testzsirSzazalek = testZsirSzazalekSzamitas(nem.Text, eletkora, bmi);
                testzsir.Text = $"{testzsirSzazalek}";

                if (testzsirSzazalek)

                if (nem.Text == "férfi")
                {
                    tzsirAlso.Text = "2%";
                    tzsirKozep.Text = "15%";
                    tzsirFelso.Text = "40%";
                }
                else
                {
                    tzsirAlso.Text = "10%";
                    tzsirKozep.Text = "20%";
                    tzsirFelso.Text = "45%";
                }

                if (bmi < double.Parse(fileadat[0].ertek.Replace('.', ',')))
                {
                    bmiKat.Text = $"{fileadat[0].kategoria}";
                    break;
                }
                if (bmi > double.Parse(fileadat[fileadat.Count - 1].ertek.Replace('.', ',')))
                {
                    bmiKat.Text = $"{fileadat[fileadat.Count - 1].kategoria}";
                    break;
                }

                for (int i = fileadat.Count - 1; i >= 0; i--)
                {
                    
                    if (bmi >= double.Parse(fileadat[i].ertek.Replace('.', ',')) && bmi < double.Parse(fileadat[i + 1].ertek.Replace('.', ',')))
                    {
                        bmiKat.Text = $"{fileadat[i].kategoria}";
                        break;
                    }
                    else
                    {
                        continue;
                    }
                }
            }
            olvaso.Close();
            
        }


        private void vissza_Click(object sender, EventArgs e)
        {
            var serverKapcsolat = new MySqlConnectionStringBuilder { Server = "127.0.0.1", Database = "healthtrack", UserID = "root", Password = "mysql" };
            MySqlConnection kapcsolat;
            kapcsolat = new MySqlConnection(serverKapcsolat.ConnectionString);
            kapcsolat.Open();
            var lekerdezes = kapcsolat.CreateCommand();
            lekerdezes.CommandText = "SELECT nev FROM meresek";

            var olvaso = lekerdezes.ExecuteReader();
            olvaso.Close();

            vissza.Visible = false;
            foLekerdezes.Visible = true;
            foHozzaadas.Visible = true;
            foTorles.Visible = true;
            osztalyLista.Visible = false;
            diakLista.Visible = false;
            elvalaszto1.Visible = false;
            lekerdezesPanel.Visible = false;
            osztalyLista2.Visible = false;
            diakLista2.Visible = false;
            nev.Text = "-----";
            osztaly.Text = "-----";
            oktAz.Text = "-----";
            eletkor.Text = "-----";
            nem.Text = "-----";
            suly.Text = "-----";
            magassag.Text = "-----";
            bmiErtek.Text = "-----";
            bmiKat.Text = "-----";
            testzsir.Text = "-----";
            tzsirAlso.Text = "---";
            tzsirKozep.Text = "---";
            tzsirFelso.Text = "---";

        }

        private void foHozzaadas_Click_1(object sender, EventArgs e)
        {
            var serverKapcsolat = new MySqlConnectionStringBuilder { Server = "127.0.0.1", Database = "healthtrack", UserID = "root", Password = "mysql" };
            MySqlConnection kapcsolat;
            kapcsolat = new MySqlConnection(serverKapcsolat.ConnectionString);
            kapcsolat.Open();
            var lekerdezes = kapcsolat.CreateCommand();
            lekerdezes.CommandText = "SELECT nev FROM meresek";

            var olvaso = lekerdezes.ExecuteReader();
            olvaso.Close();

            foLekerdezes.Visible = false;
            foHozzaadas.Visible = false;
            foTorles.Visible = false;
            vissza.Visible = true;
            lekerdezesPanel.Visible = true;
            bmiErtekLabel.Visible = false;
            bmiErtek.Visible = false;
            bmiKatLabel.Visible = false;
            bmiKat.Visible = false;

            nevTxb.Visible = true;
            osztalyTxb.Visible = true;
            oktAzTxb.Visible = true;
            eletkorTxb.Visible = true;
            nemTxb.Visible = true;
            sulyTxb.Visible = true;
            magassagTxb.Visible = true;
            hozzaadas.Visible= true;
            torles.Visible = false;
            testZsirLabel.Visible = false;
            testzsir.Visible = false;
            tzsirAlso.Visible = false;
            tzsirFelso.Visible = false;
            tzsirKozep.Visible = false;
            pictureBox2.Visible = false;
            tzsirKat.Visible = false;
            tzsirKatLabel.Visible = false;
            
        }

        private void hozzaadas_Click(object sender, EventArgs e)
        {
            if (nevTxb.Text != "" && osztalyTxb.Text != "" && oktAzTxb.Text != "" && eletkorTxb.Text != "" && sulyTxb.Text != "" && magassagTxb.Text != "")
            {
                var serverKapcsolat = new MySqlConnectionStringBuilder { Server = "127.0.0.1", Database = "healthtrack", UserID = "root", Password = "mysql" };
                MySqlConnection kapcsolat;
                kapcsolat = new MySqlConnection(serverKapcsolat.ConnectionString);
                kapcsolat.Open();
                var lekerdezes = kapcsolat.CreateCommand();
                lekerdezes.CommandText = @"
                INSERT INTO szemelyek (oktatasi_azonosito, nev)
                VALUES (@oktAz, @nev);
    
                INSERT INTO meresek (oktatasi_azonosito, eletkor, suly, magassag, osztaly, nev, nem)
                VALUES (@oktAz, @eletkor, @suly, @magassag, @osztaly, @nev, @nem)";

                lekerdezes.Parameters.AddWithValue("@oktAz", oktAzTxb.Text);
                lekerdezes.Parameters.AddWithValue("@nev", nevTxb.Text);
                lekerdezes.Parameters.AddWithValue("@eletkor", int.Parse(eletkorTxb.Text));
                lekerdezes.Parameters.AddWithValue("@suly", decimal.Parse(sulyTxb.Text));
                lekerdezes.Parameters.AddWithValue("@magassag", decimal.Parse(magassagTxb.Text));
                lekerdezes.Parameters.AddWithValue("@osztaly", osztalyTxb.Text);
                lekerdezes.Parameters.AddWithValue("@nem", nemTxb.Text);
                lekerdezes.ExecuteNonQuery();

                StreamWriter sw = new StreamWriter("healthtrack.sql", true);

                sw.WriteLine("INSERT INTO szemelyek(oktatasi_azonosito, nev) VALUES(@oktAz, @nev);");
                sw.WriteLine("INSERT INTO meresek (oktatasi_azonosito, eletkor, suly, magassag, osztaly, nev, nem)\r\n                VALUES (@oktAz, @eletkor, @suly, @magassag, @osztaly, @nev, @nem)");

                sw.Close();

                MessageBox.Show("Sikeres hozzáadás");
                nevTxb.Text = "";
                osztalyTxb.Text = "";
                oktAzTxb.Text = "";
                eletkorTxb.Text = "";
                nemTxb.Text = "";
                sulyTxb.Text = "";
                magassagTxb.Text = "";

            }
            else
            {
                MessageBox.Show("Sikertelen");
            }
        }

        private void foTorles_Click(object sender, EventArgs e)
        {
            var serverKapcsolat = new MySqlConnectionStringBuilder { Server = "127.0.0.1", Database = "healthtrack", UserID = "root", Password = "mysql" };
            MySqlConnection kapcsolat;
            kapcsolat = new MySqlConnection(serverKapcsolat.ConnectionString);
            kapcsolat.Open();
            var lekerdezes = kapcsolat.CreateCommand();
            lekerdezes.CommandText = "SELECT nev FROM meresek";

            var olvaso = lekerdezes.ExecuteReader();
            olvaso.Close();

            foLekerdezes.Visible = false;
            foHozzaadas.Visible = false;
            foTorles.Visible = false;
            osztalyLista2.Visible = true;
            vissza.Visible = true;
            lekerdezesPanel.Visible = true;

            nevTxb.Visible = false;
            osztalyTxb.Visible = false;
            oktAzTxb.Visible = false;
            eletkorTxb.Visible = false;
            nemTxb.Visible = false;
            sulyTxb.Visible = false;
            magassagTxb.Visible = false;
            hozzaadas.Visible = false;
            torles.Visible = true;

            tzsirAlso.Visible = true;
            tzsirFelso.Visible = true;
            tzsirKozep.Visible = true;
            pictureBox2.Visible = true;
            tzsirKat.Visible = true;
            tzsirKatLabel.Visible = true;

        }

        private void osztalyLista2_SelectedIndexChanged(object sender, EventArgs e)
        {
            diakLista2.Visible = true;
            kivalasztottOsztaly = osztalyLista2.Text;
            diakok.Clear();
            diakLista2.Items.Clear();
            nev.Text = "-----";
            osztaly.Text = "-----";
            oktAz.Text = "-----";
            eletkor.Text = "-----";
            nem.Text = "-----";
            suly.Text = "-----";
            magassag.Text = "-----";
            bmiErtek.Text = "-----";
            bmiKat.Text = "-----";

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
                diakLista2.Items.Add(item);
            }
        }
        private void diakLista2_SelectedIndexChanged(object sender, EventArgs e)
        {
            kivalasztottNev = diakLista2.Text;

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
                sulySz = olvaso.GetDouble(3);
                magassag.Text = olvaso.GetInt32(4).ToString();
                magassagSz = olvaso.GetDouble(4) / 100;
                osztaly.Text = kivalasztottOsztaly;
                nev.Text = kivalasztottNev;
                nem.Text = olvaso.GetString(7);
                bmi = Math.Round(sulySz / (magassagSz * magassagSz), 1);
                bmiErtek.Text = bmi.ToString();

                if (nem.Text == "férfi")
                {
                    tzsirAlso.Text = "2%";
                    tzsirKozep.Text = "15%";
                    tzsirFelso.Text = "40%";
                }
                else
                {
                    tzsirAlso.Text = "10%";
                    tzsirKozep.Text = "20%";
                    tzsirFelso.Text = "45%";
                }

                if (bmi < double.Parse(fileadat[0].ertek.Replace('.', ',')))
                {
                    bmiKat.Text = $"{fileadat[0].kategoria}";
                    break;
                }
                if (bmi > double.Parse(fileadat[fileadat.Count - 1].ertek.Replace('.', ',')))
                {
                    bmiKat.Text = $"{fileadat[fileadat.Count - 1].kategoria}";
                    break;
                }

                for (int i = fileadat.Count - 1; i >= 0; i--)
                {

                    if (bmi >= double.Parse(fileadat[i].ertek.Replace('.', ',')) && bmi < double.Parse(fileadat[i + 1].ertek.Replace('.', ',')))
                    {
                        bmiKat.Text = $"{fileadat[i].kategoria}";
                        break;
                    }
                    else
                    {
                        continue;
                    }
                }
            }
            olvaso.Close();
        }

        private void torles_Click(object sender, EventArgs e)
        {
            nev.Text = "-----";
            osztaly.Text = "-----";
            oktAz.Text = "-----";
            eletkor.Text = "-----";
            nem.Text = "-----";
            suly.Text = "-----";
            magassag.Text = "-----";
            bmiErtek.Text = "-----";
            bmiKat.Text = "-----";
            tzsirAlso.Text = "---";
            tzsirKozep.Text = "---";
            tzsirFelso.Text = "---";

            var serverKapcsolat = new MySqlConnectionStringBuilder { Server = "127.0.0.1", Database = "healthtrack", UserID = "root", Password = "mysql" };
            MySqlConnection kapcsolat;
            kapcsolat = new MySqlConnection(serverKapcsolat.ConnectionString);
            kapcsolat.Open();
            var lekerdezes = kapcsolat.CreateCommand();
            MySqlCommand torlesMeresek = new MySqlCommand(
                    "DELETE FROM meresek WHERE nev = @nev AND osztaly = @osztaly",
                    kapcsolat
                );
            torlesMeresek.Parameters.AddWithValue("@nev", kivalasztottNev);
            torlesMeresek.Parameters.AddWithValue("@osztaly", kivalasztottOsztaly);
            torlesMeresek.ExecuteNonQuery();

            StreamWriter sw = new StreamWriter("healthtrack.sql", true);

            sw.WriteLine("DELETE FROM meresek WHERE nev = @nev AND osztaly = @osztaly",
                    kapcsolat);

            sw.Close();

            

            diakLista2.Items.Remove(kivalasztottNev);
        }
    }
    class kategoriaBMI
    {
        public string ertek;
        public string kategoria;
        public kategoriaBMI (string line)
        {
            string[] szet = line.Split('-');
            this.ertek = szet[0].Trim();
            this.kategoria = szet[1].Trim();
        }
    }
}