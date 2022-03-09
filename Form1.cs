using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using Microsoft.Office.Interop.Excel;
using ZedGraph;
using org.mariuszgromada.math.mxparser;
using System.Threading;

namespace lab_3
{
    public partial class Form1 : Form
    {
        double maximum = -1 * Math.Pow(2, 69);
        double minimum = Math.Pow(2, 69);

        List<double> listX = new List<double>();
        public Form1()
        {
            InitializeComponent();
        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog1.FileName = String.Empty;
            DialogResult res = openFileDialog1.ShowDialog();
            if (res != DialogResult.OK) return;

            try
            {
                dataGridView1.Rows.Clear();
                Application AppExcel = new Application();
                Workbook workSpaceExcel = AppExcel.Workbooks.Open(openFileDialog1.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
                Worksheet sheetExcel = (Worksheet)workSpaceExcel.Sheets[1];
                if (sheetExcel.Rows.CurrentRegion.EntireRow.Count == 1)
                {
                    MessageBox.Show("В Excel-файле только одна строка");
                }
                else
                {
                    dataGridView1.Rows.Clear();
                    var lastCell = sheetExcel.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);  //...SpecialCells(XlCellType.xlCellTypeLastCell) - возвращает диапазон ячеек
                    string x = String.Empty;
                    string y = String.Empty;
                    for (int i = 0; i < lastCell.Row; i++)
                    {
                        x = sheetExcel.Cells[i + 1, 1].Text.ToString();
                        y = sheetExcel.Cells[i + 1, 2].Text.ToString();
                        if (x.Trim() != String.Empty && y.Trim() != String.Empty) {
                            dataGridView1.Rows.Add();
                            dataGridView1[0, i].Value = x;
                            dataGridView1[1, i].Value = y;
                        }
                    }
                }
                workSpaceExcel.Close(false, Type.Missing, Type.Missing);
                AppExcel.Quit();

            }
            catch (Exception exception)
            {
                MessageBox.Show("Ошибка при загрузке файла!");
            }
        }

        private void рассчитатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = false;
          
            GraphPane graph = zedGraphControl1.GraphPane;
            graph.CurveList.Clear();
            PointPairList list_points = new PointPairList();
            List<double> values = new List<double>();

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                list_points.Add(Convert.ToDouble(dataGridView1[0, i].Value), Convert.ToDouble(dataGridView1[1, i].Value));
            }
            dataGridView1.AllowUserToAddRows = true;

            LineItem line = graph.AddCurve("Точки", list_points, Color.Green, SymbolType.Circle);
            line.Line.IsVisible = false;


            Thread graphic = new Thread(new ThreadStart(asyncGraph));
            graphic.Start();


            zedGraphControl1.AxisChange();
            zedGraphControl1.Invalidate();
        }

        private void quadraticFunc()
        {
            double xi = 0;
            double yi = 0;
            double xy = 0;
            double x2 = 0;
            double x3 = 0;
            double x4 = 0;
            double x2y = 0;
            double yifyi = 0; // сумма (yi - y^) 
            double yimyavg = 0; // сумма (yi - yсреднее)
            double yavg; // y среднее
            double determ;
            int tableCount = dataGridView1.RowCount - 1;
            double a, b, c;
            double det, deta, detb, detc;
            string exp = String.Empty;
            string exp_output = String.Empty;
            
            PointPairList quadraticList = new PointPairList();

            for (int i = 0; i < tableCount; i++)
            {

                    xi += Convert.ToDouble(dataGridView1[0, i].Value);
                    yi += Convert.ToDouble(dataGridView1[1, i].Value);
                    xy += Convert.ToDouble(dataGridView1[0, i].Value) * Convert.ToDouble(dataGridView1[1, i].Value);
                    x2 += Math.Pow(Convert.ToDouble(dataGridView1[0, i].Value), 2);
                    x3 += Math.Pow(Convert.ToDouble(dataGridView1[0, i].Value), 3);
                    x4 += Math.Pow(Convert.ToDouble(dataGridView1[0, i].Value), 4);
                    x2y += Math.Pow(Convert.ToDouble(dataGridView1[0, i].Value), 2) * Convert.ToDouble(dataGridView1[1, i].Value);
                
            }
            yavg = yi / tableCount; // y среднее

            det = (x2 * x2 * x2) + (xi * xi * x4) + (x3 * x3 * tableCount) - (tableCount * x2 * x4) - (xi * x2 * x3) - (xi * x3 * x2);
            deta = (yi * x2 * x2) + (xi * xi * x2y) + (xy * x3 * tableCount) - (x2y * x2 * tableCount) - (xi * x2 * xy) - (yi * xi * x3);
            a = deta / det;
            detb = (x2 * x2 * xy) + (yi * xi * x4) + (x3 * x2y * tableCount) - (x4 * xy * tableCount) - (yi * x2 * x3) - (x2 * xi * x2y);
            b = detb / det;
            detc = (x2 * x2 * x2y) + (xi * xy * x4) + (x3 * x3 * yi) - (x4 * x2 * yi) - (xi * x2y * x3) - (x2 * xy * x3);
            c = detc / det;
            if (b >= 0)
            {
                exp += Convert.ToString(a).Replace(',', '.') + "*x^2+" + Convert.ToString(b).Replace(',', '.') + "*x";
                exp_output += Convert.ToString(Math.Round(a, 3)).Replace(',', '.') + "*x^2+" + Convert.ToString(Math.Round(b, 3)).Replace(',', '.') + "*x";
                if (c >= 0)
                {
                    exp += "+" + Convert.ToString(c).Replace(',', '.');
                    exp_output += "+" + Convert.ToString(Math.Round(c, 3)).Replace(',', '.');
                }
                else
                {
                    exp += Convert.ToString(c).Replace(',', '.');
                    exp_output += Convert.ToString(Math.Round(c, 3)).Replace(',', '.');
                }
            }
            else
            {
                exp = Convert.ToString(a).Replace(',', '.') + "*x^2" + Convert.ToString(b).Replace(',', '.') + "*x";
                exp_output = Convert.ToString(Math.Round(a, 3)).Replace(',', '.') + "*x^2" + Convert.ToString(Math.Round(b, 3)).Replace(',', '.') + "*x";
                if (c >= 0)
                {
                    exp += "+" + Convert.ToString(c).Replace(',', '.');
                    exp_output += "+" + Convert.ToString(Math.Round(c, 3)).Replace(',', '.');
                }
                else
                {
                    exp += Convert.ToString(c).Replace(',', '.');
                    exp_output += Convert.ToString(Math.Round(c, 3)).Replace(',', '.');
                }
            }

             
            
            for (int i = 0; i < tableCount; i++)  // цикл для нахождения числителя и знаменателя для расчёта коэффициента корреляции
            {
                    yifyi += Math.Pow(Convert.ToDouble(dataGridView1[1, i].Value) - func(Convert.ToDouble(dataGridView1[0, i].Value), exp), 2);
                    yimyavg += Math.Pow(Convert.ToDouble(dataGridView1[1, i].Value) - yavg, 2);
    
            }

            determ = Math.Pow(Math.Sqrt(1 - yifyi / yimyavg), 2); // Коэффициент детерминации

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (Convert.ToDouble(dataGridView1[0, i].Value) > maximum)
                {
                    maximum = Convert.ToDouble(dataGridView1[0, i].Value);
                }
                if (Convert.ToDouble(dataGridView1[0, i].Value) < minimum)
                {
                    minimum = Convert.ToDouble(dataGridView1[0, i].Value);
                }
            }

            for (int i = Convert.ToInt32(minimum); i <= maximum; i++)
            {
                quadraticList.Add(i, func(i, exp));
            }

            label6.Invoke((MethodInvoker)delegate { label6.Text = "Коэффициент детерминации: " + determ.ToString(); });
            label4.Invoke((MethodInvoker)delegate { label4.Text = "Функция: " + exp_output; });
            addPoints(quadraticList, "Квадратичная функция");
        }

        private void linearFunc()
        {
            double xi = 0;
            double yi = 0;
            double xy = 0;
            double x2 = 0;
            double y2 = 0;
            double determ; // Коэфициент детерминации
            int textCount = dataGridView1.RowCount - 1;
            double a, b;
            string exp;
            string exp_output;
            PointPairList linearList = new PointPairList();

            for (int i = 0; i < textCount; i++)
            {
                 xi += Convert.ToDouble(dataGridView1[0, i].Value);
                 yi += Convert.ToDouble(dataGridView1[1, i].Value);
                 xy += Convert.ToDouble(dataGridView1[0, i].Value) * Convert.ToDouble(dataGridView1[1, i].Value);
                 x2 += Math.Pow(Convert.ToDouble(dataGridView1[0, i].Value), 2);
                 y2 += Math.Pow(Convert.ToDouble(dataGridView1[1, i].Value), 2);
                
            }
            a = ((xi * yi) - textCount * xy) / (Math.Pow(xi, 2) - textCount * x2);
            b = (xi * xy - x2 * yi) / (Math.Pow(xi, 2) - textCount * x2);
            determ = Math.Pow((textCount * xy - xi * yi) / Math.Sqrt((textCount * x2 - Math.Pow(xi, 2)) * (textCount * y2 - Math.Pow(yi, 2))), 2); // нахождение коэфициента детерминации
            if (b >= 0)
            {
                exp = Convert.ToString(a).Replace(',', '.') + "*x+" + Convert.ToString(b).Replace(',', '.');
                exp_output = Convert.ToString(Math.Round(a, 3)).Replace(',', '.') + "*x+" + Convert.ToString(Math.Round(b, 3)).Replace(',', '.');
            }
            else
            {
                exp = Convert.ToString(a).Replace(',', '.') + "*x" + Convert.ToString(b).Replace(',', '.');
                exp_output = Convert.ToString(Math.Round(a, 3)).Replace(',', '.') + "*x" + Convert.ToString(Math.Round(b, 3)).Replace(',', '.');
            }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (Convert.ToDouble(dataGridView1[0, i].Value) > maximum)
                {
                    maximum = Convert.ToDouble(dataGridView1[0, i].Value);
                }
                if (Convert.ToDouble(dataGridView1[0, i].Value) < minimum)
                {
                    minimum = Convert.ToDouble(dataGridView1[0, i].Value);
                }
            }

            for (int i = Convert.ToInt32(minimum); i <= maximum; i++)
            {
                
                linearList.Add(i, func(i, exp));
            }

            label2.Invoke((MethodInvoker)delegate { label2.Text = "Коэффициент детерминации: " + determ.ToString(); });
            label3.Invoke((MethodInvoker)delegate { label3.Text = "Функция: " + exp_output; });
            addPoints(linearList, "Линейная функция");

        }

        private double func(double x, string exp)
        {

            try
            {
                Argument xmain = new Argument("x");
                Expression y = new Expression(exp.ToLower(), xmain);

                xmain.setArgumentValue(x);
                return y.calculate();
            }
            catch (Exception e)
            {
                return 0;
            }

        }
        private void addPoints(PointPairList l, string name)
        {
            GraphPane graph = zedGraphControl1.GraphPane;

            if (name == "Линейная функция")
            {
                graph.AddCurve(name, l, Color.Blue, SymbolType.None);
            }
            else
            {
                graph.AddCurve(name, l, Color.Red, SymbolType.None);
            }

            zedGraphControl1.AxisChange();
            zedGraphControl1.Invalidate();
        }

        private void очиститьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0) { dataGridView1.Rows.Clear(); }
            zedGraphControl1.GraphPane.CurveList.Clear();
            listX.Clear();

            zedGraphControl1.AxisChange();  //Масштабирование и обновление данных об j
            zedGraphControl1.Invalidate();  //Обновление графика
        }

        private void сгенерироватьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Thread thread = new Thread(new ThreadStart(generation));
            thread.Start();

        }

        private void dataGridViewGeneration()
        {
            try
            {
                dataGridView1.Invoke((MethodInvoker)delegate { dataGridView1.Rows.Clear(); });

                GraphPane graph = zedGraphControl1.GraphPane;  //Создание системы координат
                graph.CurveList.Clear();

                if (textBox1.Text == String.Empty)
                {
                    MessageBox.Show("Введите кол-во точек для генерации");
                }
                else
                {

                    int n = Convert.ToInt32(textBox1.Text);

                    Random rnd = new Random();

                    for (int i = 0; i < n; i++)
                    {
                        dataGridView1.Invoke((MethodInvoker)delegate
                        {
                            dataGridView1.Rows.Add();
                            dataGridView1[0, i].Value = rnd.Next(-50, 50);
                            dataGridView1[1, i].Value = rnd.Next(-50, 50);
                        });
                    }


                }
            }
            catch (Exception)
            {
                MessageBox.Show("Введите целочисленное значение");
            }

        }

        private async void asyncGraph()
        {
            await Task.Run(() => linearFunc());
            await Task.Run(() => quadraticFunc());
        }

        private async void generation()
        {
            await Task.Run(() => dataGridViewGeneration());
        }

    }
}

