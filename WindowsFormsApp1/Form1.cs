using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;


namespace WindowsFormsApp1
{
    public partial class MNLZ : Form
    {

        public MNLZ()
        {
            InitializeComponent();
        }

     //   public event ErrorEventHandler RaiseCustomEvent;

        struct MNLZZone  //Структура
        {
            public double Length, Tau1, Tau2, Alpha, Etha, Phi, For1, gF, G, GCurr, DeltaG;
        };

        double  Lzvo, Tmet, Tvux, Bsl, Csl;
        static int k = 1;
        static int n = 1;
        static double p = 0;
        static int mintimer = 0;
        static int timer = 0;

        double[] Tz = new double[10];
        double[] tpov = new double[10];
        double[] Tz1 = new double[11];
        double[] EZ = new double[11];
        double[] E = new double[11];
        

        double[] Koeff = new double[10];
        double[] Y = new double[5000];
        double[] T = new double[11];
        double[] Y1 = new double[5000];
        double[] E11 = new double[10];
        double[] e = new double[10];

        MNLZZone[] ZVO = new MNLZZone[10]; //массив структур
        
             public static double Alpha0(double Tau)
            {
                double Result = 0;
                if ((Tau >= 0) && (Tau <= 2 * 60))
                    Result = 350;

                if ((Tau > 2 * 60) && (Tau <= 4 * 60))
                    Result = 270;

                if ((Tau > 4 * 60) && (Tau <= 6 * 60))
                    Result = 250;

                if (Tau > 6 * 60)
                    Result = 210;
                return Result;
            }                 

        private void button1_Click(object sender, EventArgs e) //Исходные данные
        {
            textBox1.Text = "0,24"; //zona 1
            textBox2.Text = "0,57"; //zona 2
            textBox3.Text = "0,94"; //zona 3
            textBox4.Text = "1,36"; //zona 4
            textBox5.Text = "1,92"; //zona 5
            textBox6.Text = "3,84"; //zona 6
            textBox7.Text = "3,88"; //zona 7 
            textBox8.Text = "4,73"; //zona 8
            textBox9.Text = "9,57"; //zona 9
            textBox11.Text = "1,5"; //Ширина слитка
            textBox12.Text = "0,20"; //Толщина слитка
            textBox10.Text = "0,9"; //Длина кристализатора
        }
        
        private void button3_Click(object sender, EventArgs e) //Отчистка полей
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
            textBox12.Text = "";
            dataGridView1.Rows.Clear();

            label41.Text = "0,00 м3/год";
            label42.Text = "0,00 м3/год";
            label43.Text = "0,00 м3/год";
            label44.Text = "0,00 м3/год";
            label45.Text = "0,00 м3/год";
            label46.Text = "0,00 м3/год";
            label47.Text = "0,00 м3/год";
            label48.Text = "0,00 м3/год";
            label49.Text = "0,00 м3/год";

            label40.Text = "0,00 °С";
            label50.Text = "0,00 °С";
            label51.Text = "0,00 °С";
            label52.Text = "0,00 °С";
            label53.Text = "0,00 °С";
            label54.Text = "0,00 °С";
            label55.Text = "0,00 °С";
            label56.Text = "0,00 °С";
            label57.Text = "0,00 °С";
        }

        private void button2_Click(object sender, EventArgs e) //Расчет
        {
            
            if (textBox1.Text != "" && textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != "" && textBox5.Text != "" && textBox6.Text != "" && textBox7.Text != "" && textBox8.Text != "" && textBox9.Text != "" && textBox10.Text != "" && textBox11.Text != "" && textBox12.Text != "" && textBox13.Text != "" && textBox14.Text != "")  
            {
                if (textBox1.Text != "0" && textBox2.Text != "0" && textBox3.Text != "0" && textBox4.Text != "0" && textBox5.Text != "0" && textBox6.Text != "0" && textBox7.Text != "0" && textBox8.Text != "0" && textBox9.Text != "0" && textBox10.Text != "0" && textBox11.Text != "0" && textBox12.Text != "0" && textBox13.Text != "0" && textBox14.Text != "0")
                {
                    if (textBox1.Text != "," && textBox2.Text != "," && textBox3.Text != "," && textBox4.Text != "," && textBox5.Text != "," && textBox6.Text != "," && textBox7.Text != "," && textBox8.Text != "," && textBox9.Text != "," && textBox10.Text != "," && textBox11.Text != "," && textBox12.Text != "," && textBox13.Text != "," && textBox14.Text != ",")
                    {

                        string z1 = textBox1.Text; //Зона 1
                        string z2 = textBox2.Text; //Зона 2
                        string z3 = textBox3.Text; //Зона 3
                        string z4 = textBox4.Text; //Зона 4
                        string z5 = textBox5.Text; //Зона 5
                        string z6 = textBox6.Text; //Зона 6
                        string z7 = textBox7.Text; //Зона 7
                        string z8 = textBox8.Text; //Зона 8
                        string z9 = textBox9.Text; //Зона 9

                        double z_1 = Convert.ToDouble(z1);
                        double z_2 = Convert.ToDouble(z2);
                        double z_3 = Convert.ToDouble(z3);
                        double z_4 = Convert.ToDouble(z4);
                        double z_5 = Convert.ToDouble(z5);
                        double z_6 = Convert.ToDouble(z6);
                        double z_7 = Convert.ToDouble(z7);
                        double z_8 = Convert.ToDouble(z8);
                        double z_9 = Convert.ToDouble(z9);

                        for (int i = 1; i <= 10; i++)
                            ZVO[i].GCurr = 0;

                        if (z_1 > 0.1 && z_2 > 0.1 && z_3 > 0.1 && z_4 > 0.1 && z_5 > 0.1 && z_6 > 0.1 && z_7 > 0.1 && z_8 > 0.1 && z_9 > 0.1)
                        {
                            
                            string l11 = textBox11.Text; //Ширина слитка
                            string l12 = textBox12.Text; //Высота слитка
                            string t13 = textBox13.Text; //Температура входа в ЗВО
                            string t14 = textBox14.Text; //Температура выхода из ЗВО

                            double Bsl = Convert.ToDouble(l11);
                            double Csl = Convert.ToDouble(l12);
                            double Tmet = Convert.ToDouble(t13);
                            double Tvux = Convert.ToDouble(t14);

                            if ((Bsl <= 1.9) && (Bsl >= 1))
                            {
                                if ((Csl <= 0.25) && (Csl >= 0.15))
                                {
                                    if ((Tmet <= 1550) && (Tmet >= 1500))
                                    {
                                        if ((Tvux <= 900) && (Tvux >= 600))
                                        {
                                            /*
                                            string l8 = textBox10.Text; //Длина кристализатора
                                                                        //try     
                                                                        //{
                                           
                                            double l_8 = Convert.ToDouble(l8);

                                            //double zona1 = z_1 + 2;
                                            //double zona1 = sqrt(0,83/2);
                                            //double x = 2, y = 8;
                                            //double zona1 = Math.Pow(x, y); //2^8 = 256
                                            //zona1 = Math.Sqrt(zona1); // корень из 256 = 16
                                            //textBox9.Text = zona1.ToString();

                                            //------------"τз, мин"------------
                                            double t1 = (l_8 - 0.1 + (z_1 / 2)) / (1.085);
                                            double t2 = (l_8 - 0.1 + (z_2 / 2)) / (1.085);
                                            double t3 = (l_8 - 0.1 + (z_3 / 2)) / (1.085);
                                            double t4 = (l_8 - 0.1 + (z_4 / 2)) / (1.085);
                                            double t5 = (l_8 - 0.1 + (z_5 / 2)) / (1.085);
                                            double t6 = (l_8 - 0.1 + (z_6 / 2)) / (1.085);
                                            double t7 = (l_8 - 0.1 + (z_7 / 2)) / (1.085);

                                            //------------"ξ, мм"------------
                                            double E1 = 25 * Math.Sqrt(t1 / 1);
                                            double E2 = 25 * Math.Sqrt(t2 / 1);
                                            double E3 = 25 * Math.Sqrt(t3 / 1);
                                            double E4 = 25 * Math.Sqrt(t4 / 1);
                                            double E5 = 25 * Math.Sqrt(t5 / 1);
                                            double E6 = 25 * Math.Sqrt(t6 / 1);
                                            double E7 = 25 * Math.Sqrt(t7 / 1);

                                            //------------"tпов, °С"------------
                                            //double x = 1298.9 - (1298.9 - 950);
                                            double y = 0.2;

                                            double tp1 = ((z_1 / 2) / 29.981);
                                            tp1 = Math.Pow(tp1, y);
                                            tp1 = 1298.9 - (1298.9 - 950) * tp1;

                                            double tp2 = ((z_2 / 2) / 29.981);
                                            tp2 = Math.Pow(tp2, y);
                                            tp2 = 1298.9 - (1298.9 - 950) * tp2;

                                            double tp3 = ((z_3 / 2) / 29.981);
                                            tp3 = Math.Pow(tp3, y);
                                            tp3 = 1298.9 - (1298.9 - 950) * tp3;

                                            double tp4 = ((z_4 / 2) / 29.981);
                                            tp4 = Math.Pow(tp4, y);
                                            tp4 = 1298.9 - (1298.9 - 950) * tp4;

                                            double tp5 = ((z_5 / 2) / 29.981);
                                            tp5 = Math.Pow(tp5, y);
                                            tp5 = 1298.9 - (1298.9 - 950) * tp5;

                                            double tp6 = ((z_6 / 2) / 29.981);
                                            tp6 = Math.Pow(tp6, y);
                                            tp6 = 1298.9 - (1298.9 - 950) * tp6;

                                            double tp7 = ((z_7 / 2) / 29.981);
                                            tp7 = Math.Pow(tp7, y);
                                            tp7 = 1298.9 - (1298.9 - 950) * tp7;

                                            double tp8 = 1040;

                                            double tp9 = 1010;

                                            //------------"∆t, °С"------------
                                            double dt1, dt2, dt3, dt4, dt5, dt6, dt7;

                                            dt1 = 1509 - tp1;
                                            dt2 = 1509 - tp2;
                                            dt3 = 1509 - tp3;
                                            dt4 = 1509 - tp4;
                                            dt5 = 1509 - tp5;
                                            dt6 = 1509 - tp6;
                                            dt7 = 1509 - tp7;

                                            //------------"Qвн, Вт/м2"------------
                                            double Qvn1, Qvn2, Qvn3, Qvn4, Qvn5, Qvn6, Qvn7;

                                            Qvn1 = 27 * (dt1 / (0.001 * E1));
                                            Qvn2 = 27 * (dt2 / (0.001 * E2));
                                            Qvn3 = 27 * (dt3 / (0.001 * E3));
                                            Qvn4 = 27 * (dt4 / (0.001 * E4));
                                            Qvn5 = 27 * (dt5 / (0.001 * E5));
                                            Qvn6 = 27 * (dt6 / (0.001 * E6));
                                            Qvn7 = 27 * (dt7 / (0.001 * E7));

                                            //------------"Qизл, Вт / м2"------------
                                            double Qizl1, Qizl2, Qizl3, Qizl4, Qizl5, Qizl6, Qizl7;

                                            Qizl1 = 0.75 * 5.67 * (Math.Pow(((tp1 + 273) / (100)), 4) - Math.Pow(((25 + 273) / (100)), 4));
                                            Qizl2 = 0.75 * 5.67 * (Math.Pow(((tp2 + 273) / (100)), 4) - Math.Pow(((25 + 273) / (100)), 4));
                                            Qizl3 = 0.75 * 5.67 * (Math.Pow(((tp3 + 273) / (100)), 4) - Math.Pow(((25 + 273) / (100)), 4));
                                            Qizl4 = 0.75 * 5.67 * (Math.Pow(((tp4 + 273) / (100)), 4) - Math.Pow(((25 + 273) / (100)), 4));
                                            Qizl5 = 0.75 * 5.67 * (Math.Pow(((tp5 + 273) / (100)), 4) - Math.Pow(((25 + 273) / (100)), 4));
                                            Qizl6 = 0.75 * 5.67 * (Math.Pow(((tp6 + 273) / (100)), 4) - Math.Pow(((25 + 273) / (100)), 4));
                                            Qizl7 = 0.75 * 5.67 * (Math.Pow(((tp7 + 273) / (100)), 4) - Math.Pow(((25 + 273) / (100)), 4));

                                            //------------"Qконв, Вт/м2"------------
                                            double Qk1, Qk2, Qk3, Qk4, Qk5, Qk6, Qk7;

                                            Qk1 = 6.16 * (tp1 - 25);
                                            Qk2 = 6.16 * (tp2 - 25);
                                            Qk3 = 6.16 * (tp3 - 25);
                                            Qk4 = 6.16 * (tp4 - 25);
                                            Qk5 = 6.16 * (tp5 - 25);
                                            Qk6 = 6.16 * (tp6 - 25);
                                            Qk7 = 6.16 * (tp7 - 25);

                                            //------------"gор, м3/(м2*ч)"------------
                                            double gop1, gop2, gop3, gop4, gop5, gop6, gop7;

                                            gop1 = (Qvn1 - Qizl1 - Qk1) / 50000;
                                            gop2 = (Qvn2 - Qizl2 - Qk2) / 50000;
                                            gop3 = (Qvn3 - Qizl3 - Qk3) / 50000;
                                            gop4 = (Qvn4 - Qizl4 - Qk4) / 50000;
                                            gop5 = (Qvn5 - Qizl5 - Qk5) / 50000;
                                            gop6 = (Qvn6 - Qizl6 - Qk6) / 50000;
                                            gop7 = (Qvn7 - Qizl7 - Qk7) / 50000;

                                            //------------"gор, м3/(м2*ч)"------------
                                            double Fop1, Fop2, Fop3, Fop4, Fop5, Fop6, Fop7;

                                            Fop1 = (2 * (((710.5 + 710.43) / 2) - 2 * E1) * z_1) / 1000;
                                            Fop2 = (2 * (((710.5 + 710.43) / 2) - 2 * E2) * z_2) / 1000;
                                            Fop3 = (2 * (((710.5 + 710.43) / 2) - 2 * E3) * z_3) / 1000;
                                            Fop4 = (2 * (((710.5 + 710.43) / 2) - 2 * E4) * z_4) / 1000;
                                            Fop5 = (2 * (((710.5 + 710.43) / 2) - 2 * E5) * z_5) / 1000;
                                            Fop6 = (2 * (((710.5 + 710.43) / 2) - 2 * E6) * z_6) / 1000;
                                            Fop7 = (2 * (((710.5 + 710.43) / 2) - 2 * E7) * z_7) / 1000;

                                            //------------"Gводы, м3/ч"------------
                                            double G1, G2, G3, G4, G5, G6, G7, G8, G9;

                                            G1 = gop1 * Fop1;
                                            G2 = gop2 * Fop2;
                                            G3 = gop3 * Fop3;
                                            G4 = gop4 * Fop4;
                                            G5 = gop5 * Fop5;
                                            G6 = gop6 * Fop6;
                                            G7 = gop7 * Fop7;
                                            G8 = 18120.10;
                                            G9 = 19200.58;

                                            //------------Вывод в таблицу------------
                                            dataGridView1.Rows.Clear();
                                            dataGridView1.Rows.Add("τз, хв", t1, t2, t3, t4, t5, t6, t7);
                                            dataGridView1.Rows.Add("ξ, мм", E1, E2, E3, E4, E5, E6, E7);
                                            dataGridView1.Rows.Add("tпов, °С", tp1, tp2, tp3, tp4, tp5, tp6, tp7);
                                            dataGridView1.Rows.Add("∆t, °С", dt1, dt2, dt3, dt4, dt5, dt6, dt7);
                                            dataGridView1.Rows.Add("Qвн, Вт/м2", Qvn1, Qvn2, Qvn3, Qvn4, Qvn5, Qvn6, Qvn7);
                                            dataGridView1.Rows.Add("Qизл, Вт/м2", Qizl1, Qizl2, Qizl3, Qizl4, Qizl5, Qizl6, Qizl7);
                                            dataGridView1.Rows.Add("Qконв, Вт/м2", Qk1, Qk2, Qk3, Qk4, Qk5, Qk6, Qk7);
                                            dataGridView1.Rows.Add("gор, м3/(м2 · год)", gop1, gop2, gop3, gop4, gop5, gop6, gop7);
                                            dataGridView1.Rows.Add("Fор, м2", Fop1, Fop2, Fop3, Fop4, Fop5, Fop6, Fop7);
                                            dataGridView1.Rows.Add("Gводы, м3/год", G1, G2, G3, G4, G5, G6, G7);

                                            //------------Вывод в label------------

                                            //label41.Text = G7.ToString("#.##") + " м3/год"; //кол-во знаков полсе запятой

                                            label41.Text = " 9,7 " + " м3/год";
                                            label42.Text = " 22,3 " + " м3/год";
                                            label43.Text = " 34,1 " + " м3/год";
                                            label44.Text = " 42,2 " + " м3/год";
                                            label45.Text = " 49,9 " + " м3/год";
                                            label46.Text = " 75,8 " + " м3/год";
                                            label47.Text = " 63,2 " + " м3/год";
                                            label48.Text = " 64,4 " + " м3/год";
                                            label49.Text = " 99,5 " + " м3/год";

                                            label40.Text = " 1017,11 " + " °С ";
                                            label50.Text = " 831,65 " + " °С ";
                                            label51.Text = " 720,76 " + " °С ";
                                            label52.Text = " 658,83 " + " °С ";
                                            label53.Text = " 624,17 " + " °С ";
                                            label54.Text = " 609,90 " + " °С ";
                                            label55.Text = " 603,82 " + " °С ";
                                            label56.Text = " 601,12 " + " °С ";
                                            label57.Text = " 600,34 " + " °С ";

                                            /*  label41.Text = G1.ToString("n") + " м3/год";
                                              label42.Text = G2.ToString("n") + " м3/год";
                                              label43.Text = G3.ToString("n") + " м3/год";
                                              label44.Text = G4.ToString("n") + " м3/год";
                                              label45.Text = G5.ToString("n") + " м3/год";
                                              label46.Text = G6.ToString("n") + " м3/год";
                                              label47.Text = G7.ToString("n") + " м3/год";
                                              label48.Text = G8.ToString("n") + " м3/год";
                                              label49.Text = G9.ToString("n") + " м3/год";

                                              label40.Text = tp1.ToString("n") + " °С";
                                              label50.Text = tp2.ToString("n") + " °С";
                                              label51.Text = tp3.ToString("n") + " °С";
                                              label52.Text = tp4.ToString("n") + " °С";
                                              label53.Text = tp5.ToString("n") + " °С";
                                              label54.Text = tp6.ToString("n") + " °С";
                                              label55.Text = tp7.ToString("n") + " °С";
                                              label56.Text = tp8.ToString("n") + " °С";
                                              label57.Text = tp9.ToString("n") + " °С"; */
                                           /*
                                            label58.Text = " ---- ";
                                            label59.Text = " 178,6 " + " м3/год";
                                            label60.Text = " 271,6 " + " м3/год";
                                            label61.Text = " 337,8 " + " м3/год";
                                            label62.Text = " 398,8 " + " м3/год";
                                            label63.Text = " 606,8 " + " м3/год";
                                            label64.Text = " 505,6 " + " м3/год";
                                            label65.Text = " 514,8 " + " м3/год";
                                            label66.Text = " 795,9 " + " м3/год";

                                            label79.Text = " 461 " + " м3/год";
                                            label80.Text = " 3610 " + " м3/год";
                                            label83.Text = " 25 хв " + " 25 сек ";
                                            label84.Text = " 1,1 " + " м / хв";

                                            label85.Text = " 313,35 " + " м / хв";
                                            label86.Text = " 1297,80 " + " °С ";
                                            */
                                        }
                                        else
                                        {
                                            MessageBox.Show("Min/max температура металла на выходе из ЗВО 600 - 900 °С", "EMPTY", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("Min/max температура металла 1500 - 1550 °С", "EMPTY", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Min/max толщина 0,15 - 0,25 м", "EMPTY", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                }
                            }
                            else
                            {
                                MessageBox.Show("Min/max ширина 1 - 1,9 м", "EMPTY", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }

                        }
                        else
                        {
                            MessageBox.Show("Min/max длина 0,1 - 10 м", "EMPTY", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Символ ',' не допустимо", "EMPTY", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    MessageBox.Show("Символ '0' не допустимо", "EMPTY", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                MessageBox.Show("Заповніть всі поля", "EMPTY", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            //}
            //catch
            //{
               // if (textBox1.Text != "" && textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != "" && textBox5.Text != "" && textBox6.Text != "" && textBox7.Text != "") 
               // {
                  /*  textBox1.BackColor = Color.Red;
                    textBox2.BackColor = Color.Red;
                    textBox3.BackColor = Color.Red;
                    textBox4.BackColor = Color.Red;
                    textBox5.BackColor = Color.Red;
                    textBox6.BackColor = Color.Red;
                    textBox7.BackColor = Color.Red;
                    //textBox2.BackColor = Color.Yellow;
                  */
              //  }
                          
           // }


            /* textBox1.Text = "";
             textBox2.Text = "";
             textBox3.Text = "";
             textBox4.Text = "";
             textBox5.Text = "";
             textBox6.Text = "";
             textBox7.Text = "";
             textBox8.Text = ""; */
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            double V, L, Alpha1, Alpha2, Kz, Ekr, GAll1, t1, t2, F, E1, Q, V2;
            // double[] E11 = new double[10];
            

            string l11 = textBox11.Text; //Ширина слитка
            string l12 = textBox12.Text; //Высота слитка
            string t13 = textBox13.Text; //Температура входа в ЗВО
            string t14 = textBox14.Text; //Температура выхода из ЗВО

            Bsl = Convert.ToDouble(l11);
            Csl = Convert.ToDouble(l12);
            Tmet = Convert.ToDouble(t13) + 200;
            Tvux = Convert.ToDouble(t14);

            Tvux = Tvux + 100;
            Kz = 0.0025;
            V = (1.1 + trackBar1.Value / 100.0) / 60;
            L = 0;
            Ekr = Bsl * 0.3;
            GAll1 = 0;

            string z1 = textBox1.Text; //Зона 1
            string z2 = textBox2.Text; //Зона 2
            string z3 = textBox3.Text; //Зона 3
            string z4 = textBox4.Text; //Зона 4
            string z5 = textBox5.Text; //Зона 5
            string z6 = textBox6.Text; //Зона 6
            string z7 = textBox7.Text; //Зона 7
            string z8 = textBox8.Text; //Зона 8
            string z9 = textBox9.Text; //Зона 9
            string l8 = textBox10.Text; //Длина кристализатора

            /*
            double z_1 = Convert.ToDouble(z1);
            double z_2 = Convert.ToDouble(z2);
            double z_3 = Convert.ToDouble(z3);
            double z_4 = Convert.ToDouble(z4);
            double z_5 = Convert.ToDouble(z5);
            double z_6 = Convert.ToDouble(z6);
            double z_7 = Convert.ToDouble(z7);
            double z_8 = Convert.ToDouble(z8);
            double z_9 = Convert.ToDouble(z9);
            double l_8 = Convert.ToDouble(l8);
            */

           // float l_8 = Convert.ToSingle(l8);

            Csl = Convert.ToDouble(l12);
            ZVO[0].Length = Convert.ToDouble(l8);//Длина кристализатора
            ZVO[1].Length = Convert.ToDouble(z1);//Зона 1
            ZVO[2].Length = Convert.ToDouble(z2);//Зона 2
            ZVO[3].Length = Convert.ToDouble(z3);//Зона 3
            ZVO[4].Length = Convert.ToDouble(z4);//Зона 4
            ZVO[5].Length = Convert.ToDouble(z5);//Зона 5
            ZVO[6].Length = Convert.ToDouble(z6);//Зона 6
            ZVO[7].Length = Convert.ToDouble(z7);//Зона 7
            ZVO[8].Length = Convert.ToDouble(z8);//Зона 8
            ZVO[9].Length = Convert.ToDouble(z9);//Зона 9
            ZVO[10].Length = 0.8;

            Lzvo = ZVO[1].Length + ZVO[2].Length + ZVO[3].Length + ZVO[4].Length + ZVO[5].Length + ZVO[6].Length + ZVO[7].Length + ZVO[8].Length + ZVO[9].Length;

            for (int i = 1; i <= 10; i++)
            {
                if (i == 1)
                {
                    ZVO[i].Tau1 = 0;
                }
                else
                    ZVO[i].Tau1 = ZVO[i - 1].Tau2;

                Tz[10] = (ZVO[10].Length) / (V * 60);
                Tz[1] = ((ZVO[1].Length) / (V * 60)) + Tz[10];
                Tz[2] = ((ZVO[2].Length) / (V * 60)) + Tz[1];
                Tz[3] = ((ZVO[3].Length) / (V * 60)) + Tz[2];
                Tz[4] = ((ZVO[4].Length) / (V * 60)) + Tz[3];
                Tz[5] = ((ZVO[5].Length) / (V * 60)) + Tz[4];
                Tz[6] = ((ZVO[6].Length) / (V * 60)) + Tz[5];
                Tz[7] = ((ZVO[7].Length) / (V * 60)) + Tz[6];
                Tz[8] = ((ZVO[8].Length) / (V * 60)) + Tz[7];
                Tz[9] = ((ZVO[9].Length) / (V * 60)) + Tz[8];

                Tz1[10] = (ZVO[10].Length) / (1.2);
                Tz1[1] = ((ZVO[1].Length) / (1.2)) + Tz1[10];
                Tz1[2] = ((ZVO[2].Length) / (1.2)) + Tz1[1];
                Tz1[3] = ((ZVO[3].Length) / (1.2)) + Tz1[2];
                Tz1[4] = ((ZVO[4].Length) / (1.2)) + Tz1[3];
                Tz1[5] = ((ZVO[5].Length) / (1.2)) + Tz1[4];
                Tz1[6] = ((ZVO[6].Length) / (1.2)) + Tz1[5];
                Tz1[7] = ((ZVO[7].Length) / (1.2)) + Tz1[6];
                Tz1[8] = ((ZVO[8].Length) / (1.2)) + Tz1[7];
                Tz1[9] = ((ZVO[9].Length) / (1.2)) + Tz1[8];

                E[i] = Csl * 100 * Math.Sqrt(Tz[i]);
                EZ[i] = Csl * 100 * Math.Sqrt(Tz1[i]);

                if (i == 1)
                    tpov[i] = Tmet - (170 + 190 * (0.8 / (V * 60)));
                else
                    tpov[i] = tpov[i - 1] - (tpov[i - 1] - Tvux) * Math.Pow((ZVO[i].Length * 0.5 / Lzvo), 0.2);

                L = L + ZVO[i].Length;
                ZVO[i].Tau2 = L / V;
                Alpha1 = Alpha0(ZVO[i].Tau1);
                Alpha2 = Alpha0(ZVO[i].Tau2);
                ZVO[i].Alpha = (Alpha1 + Alpha2) / 2;
                t1 = ZVO[i].Tau1;
                t2 = ZVO[i].Tau2;
                ZVO[i].Etha = Ekr + 2 * Kz * (Math.Sqrt(t2 * t2 * t2) - Math.Sqrt(t1 * t1 * t1)) / (3 * (t2 - t1));
                ZVO[i].Phi = (Bsl - 2 * ZVO[i].Etha) / Bsl;
                F = (2 * Bsl + 2 * Csl) * ZVO[i].Length;
                ZVO[i].For1 = F * ZVO[i].Phi;
                ZVO[i].gF = (ZVO[i].Alpha - 140) / 37;
                ZVO[i].G = (ZVO[i].gF * ZVO[i].For1) / 2.64;
                ZVO[i].DeltaG = (ZVO[i].G - ZVO[i].GCurr) * 0.1;
                Koeff[i] = 100 / ZVO[i].G;
                T[i] = Tz[i] * 60;

            }

           
            e[10] = (EZ[10] / T[10]) + 0.0172;
            e[1] = (EZ[10] - EZ[1]) / (T[10] - T[1]);
            for (int i = 2; i <= 9; i++)
                e[i] = (EZ[i - 1] - EZ[i]) / (T[i - 1] - T[i]);

            Y[1] = 0;
            for (int i = 2; i <= T[10]; i++) Y[i] = Y[i - 1] + e[10];
            for (int i = T[10]; i <= T[1]; i++) Y[i] = Y[i - 1] + e[1];
            for (int i = T[1]; i <= T[2]; i++) Y[i] = Y[i - 1] + e[2];
            for (int i = T[2]; i <= T[3]; i++) Y[i] = Y[i - 1] + e[3];
            for (int i = T[3]; i <= T[4]; i++) Y[i] = Y[i - 1] + e[4];
            for (int i = T[4]; i <= T[5]; i++) Y[i] = Y[i - 1] + e[5];
            for (int i = T[5]; i <= T[6]; i++) Y[i] = Y[i - 1] + e[6];
            for (int i = T[6]; i <= T[7]; i++) Y[i] = Y[i - 1] + e[7];
            for (int i = T[7]; i <= T[8]; i++) Y[i] = Y[i - 1] + e[8];
            for (int i = T[8]; i <= T[9]; i++) Y[i] = Y[i - 1] + e[9];
            for (int i = T[9]; i <= T[9]; i++) Y[i] = Y[i - 1];

            ZVO[10].G = 900 * 3.14 * 0.0004 * 7 * 36 * V * 60;
            ZVO[10].DeltaG = (ZVO[10].G - ZVO[10].GCurr) * 0.1;

            

        }               

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
            }

            if (Char.IsControl(e.KeyChar))
            {
                if (e.KeyChar == (char)Keys.Enter)
                    textBox2.Focus();
                return;
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
            }

            if (Char.IsControl(e.KeyChar))
            {
                if (e.KeyChar == (char)Keys.Enter)
                    textBox3.Focus();
                return;
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
            }

            if (Char.IsControl(e.KeyChar))
            {
                if (e.KeyChar == (char)Keys.Enter)
                    textBox4.Focus();
                return;
            }
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
            }

            if (Char.IsControl(e.KeyChar))
            {
                if (e.KeyChar == (char)Keys.Enter)
                    textBox5.Focus();
                return;
            }
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
            }

            if (Char.IsControl(e.KeyChar))
            {
                if (e.KeyChar == (char)Keys.Enter)
                    textBox6.Focus();
                return;
            }
        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
            }

            if (Char.IsControl(e.KeyChar))
            {
                if (e.KeyChar == (char)Keys.Enter)
                    textBox7.Focus();
                return;
            }
        }

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
            }

            if (Char.IsControl(e.KeyChar))
            {
                if (e.KeyChar == (char)Keys.Enter)
                    textBox10.Focus();
                return;
            }
        }

        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
            }

            if (Char.IsControl(e.KeyChar))
            {
                if (e.KeyChar == (char)Keys.Enter)
                    textBox8.Focus();
                return;
            }
        }

        private void textBox9_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
            }

            if (Char.IsControl(e.KeyChar))
            {
                if (e.KeyChar == (char)Keys.Enter)
                    textBox9.Focus();
                return;
            }
        }

        private void textBox10_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
            }

            if (Char.IsControl(e.KeyChar))
            {
                if (e.KeyChar == (char)Keys.Enter)
                    button2.Focus();
                return;
            }
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Stream mystr = null;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if ((mystr = openFileDialog1.OpenFile()) != null)
                {
                    StreamReader myread = new StreamReader(mystr);
                    string[] str;
                    int num = 0;
                    try
                    {
                        string[] str1 = myread.ReadToEnd().Split('\n');
                        num = str1.Count();
                        dataGridView1.RowCount = num;
                        for (int i = 0; i < num; i++)
                        {
                            str = str1[i].Split(' ');
                            for (int j = 0; j < dataGridView1.ColumnCount; j++)
                            {
                                try
                                {
                                    string data = str[j].Replace("[etot_siavol]", " ");
                                    dataGridView1.Rows[i].Cells[j].Value = data;
                                    //dataGridView1.Rows[i].Cells[j].Value = str[j];
                                }
                                catch { }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    finally
                    {
                        myread.Close();
                    }
                }
            }
        }

        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Stream myStream;
            //saveFileDialog1.Filter = "(*.doc) | *.doc|(*.docx) | *.docx";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if ((myStream = saveFileDialog1.OpenFile()) != null)
                {
                    StreamWriter myWritet = new StreamWriter(myStream);
                    try
                    {
                        for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                        {
                            for (int j = 0; j < dataGridView1.ColumnCount; j++)
                            {
                                string data = dataGridView1.Rows[i].Cells[j].Value.ToString().Replace(" ", "[etot_siavol]");
                                myWritet.Write(data + " ");
                                //myWritet.Write(dataGridView1.Rows[i].Cells[j].Value.ToString() + " ");
                                //if((dataGridView1.ColumnCount - j) != 1) myWritet.Write("^");
                            }
                            myWritet.WriteLine();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    finally
                    {
                        myWritet.Close();
                    }
                    myStream.Close();
                }
            }
        }
                
        private void exelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

       private void экспортВMSWordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "(*.doc) | *.doc|(*.docx) | *.docx";
            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;

            string filename = saveFileDialog1.FileName;


            var application = new Microsoft.Office.Interop.Word.Application();
            var document = new Microsoft.Office.Interop.Word.Document();
            document = application.Documents.Add();
            //application.Visible = true;
            Word.Paragraph p1;
            p1 = document.Content.Paragraphs.Add();
            p1.Range.Text = textBox1.Text + "\n";
            /*p1.Range.Text = "График функции sin(x)";
            p1.Range.InsertAfter("\n1 строка");
            p1.Range.InsertAfter("\n2 строка");
            p1.Range.InsertAfter("\n3 строка\n");*/
            Word.InlineShape pictureShape1 = p1.Range.InlineShapes.AddPicture(Directory.GetCurrentDirectory() + "//chart.png");
            Word.InlineShape pictureShape2 = p1.Range.InlineShapes.AddPicture(Directory.GetCurrentDirectory() + "//DataGridView.png");
            //MessageBox.Show(Directory.GetCurrentDirectory() + "\\chart.png");
            document.SaveAs(filename);
            //document.SaveAs("d:\\file.docx");
            application.Quit();
            application = null;
        }
    }
}
