using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Math_Model_All_Work.min_el
{
    /// <summary>
    /// Логика взаимодействия для MinElement.xaml
    /// </summary>
    public partial class MinElement : Window
    {
        private DataTable costTable;
        public MinElement()
        {
            InitializeComponent();
        }
        // Создание матрицы
        private void BtnCreateCostMatrix_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int n = int.Parse(TxtNumSuppliers.Text);
                int m = int.Parse(TxtNumConsumers.Text);

                if (n <= 0 || m <= 0)
                {
                    MessageBox.Show("Введите положительные числа.");
                    return;
                }

                costTable = new DataTable();

                for (int j = 0; j < m; j++)
                    costTable.Columns.Add($"C{j + 1}", typeof(int));

                for (int i = 0; i < n; i++)
                {
                    var row = costTable.NewRow();
                    for (int j = 0; j < m; j++)
                        row[j] = 0;
                    costTable.Rows.Add(row);
                }

                DataGridCost.ItemsSource = costTable.DefaultView;
            }
            catch
            {
                MessageBox.Show("Ошибка при создании матрицы.");
            }
        }

        // Проверки
        private bool ValidateInputs()
        {
            if (costTable == null)
            {
                MessageBox.Show("Сначала создайте матрицу.");
                return false;
            }

            if (string.IsNullOrWhiteSpace(TxtSupply.Text))
            {
                MessageBox.Show("Введите запасы.");
                return false;
            }

            if (string.IsNullOrWhiteSpace(TxtDemand.Text))
            {
                MessageBox.Show("Введите потребности.");
                return false;
            }

            foreach (DataRow row in costTable.Rows)
            {
                foreach (var item in row.ItemArray)
                {
                    if (item == DBNull.Value || string.IsNullOrWhiteSpace(item.ToString()))
                    {
                        MessageBox.Show("Заполните все ячейки матрицы.");
                        return false;
                    }

                    if (!int.TryParse(item.ToString(), out _))
                    {
                        MessageBox.Show("Матрица должна содержать только числа.");
                        return false;
                    }
                }
            }

            return true;
        }

        // Получение матрицы стоимостей
        private int[,] GetCost()
        {
            int n = costTable.Rows.Count;
            int m = costTable.Columns.Count;

            int[,] cost = new int[n, m];

            for (int i = 0; i < n; i++)
                for (int j = 0; j < m; j++)
                    cost[i, j] = Convert.ToInt32(costTable.Rows[i][j]);

            return cost;
        }

        // Метод минимальных элементов
        private int[,] LeastCost(int[] supply, int[] demand, int[,] cost)
        {
            int n = supply.Length;
            int m = demand.Length;

            int[,] alloc = new int[n, m];

            var s = supply.ToArray();
            var d = demand.ToArray();

            while (s.Sum() > 0 && d.Sum() > 0)
            {
                int min = int.MaxValue;
                int mi = -1, mj = -1;

                for (int i = 0; i < n; i++)
                {
                    if (s[i] == 0) continue;

                    for (int j = 0; j < m; j++)
                    {
                        if (d[j] == 0) continue;

                        if (cost[i, j] < min)
                        {
                            min = cost[i, j];
                            mi = i;
                            mj = j;
                        }
                    }
                }

                if (mi == -1) break;

                int q = Math.Min(s[mi], d[mj]);
                alloc[mi, mj] = q;

                s[mi] -= q;
                d[mj] -= q;
            }

            return alloc;
        }

        // Подсчёт стоимости
        private int CalculateCost(int[,] alloc, int[,] cost)
        {
            int total = 0;

            for (int i = 0; i < alloc.GetLength(0); i++)
                for (int j = 0; j < alloc.GetLength(1); j++)
                    total += alloc[i, j] * cost[i, j];

            return total;
        }

        // Вывод результата
        private void ShowResult(int[,] alloc, int totalCost)
        {
            int n = alloc.GetLength(0);
            int m = alloc.GetLength(1);

            StringBuilder sb = new StringBuilder();

            sb.AppendLine("Матрица распределения:\n");

            for (int i = 0; i < n; i++)
            {
                sb.Append($"S{i + 1} | ");

                for (int j = 0; j < m; j++)
                    sb.Append($"{alloc[i, j],5}");

                sb.AppendLine();
            }

            sb.AppendLine();
            sb.AppendLine("Общая стоимость: " + totalCost);

            TxtResult.Text = sb.ToString();
        }

        // Решение
        private void BtnSolve_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!ValidateInputs()) return;

                int[] supply = TxtSupply.Text.Split(',')
                    .Select(s =>
                    {
                        if (!int.TryParse(s.Trim(), out int val))
                            throw new Exception();
                        return val;
                    }).ToArray();

                int[] demand = TxtDemand.Text.Split(',')
                    .Select(s =>
                    {
                        if (!int.TryParse(s.Trim(), out int val))
                            throw new Exception();
                        return val;
                    }).ToArray();

                if (supply.Length != costTable.Rows.Count)
                {
                    MessageBox.Show("Количество запасов должно совпадать со строками.");
                    return;
                }

                if (demand.Length != costTable.Columns.Count)
                {
                    MessageBox.Show("Количество потребностей должно совпадать со столбцами.");
                    return;
                }

                var cost = GetCost();
                var alloc = LeastCost(supply, demand, cost);
                int total = CalculateCost(alloc, cost);

                ShowResult(alloc, total);
            }
            catch
            {
                MessageBox.Show("Ошибка ввода данных.");
            }
        }

        // Очистка
        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            costTable = null;
            DataGridCost.ItemsSource = null;

            TxtSupply.Clear();
            TxtDemand.Clear();
            TxtResult.Text = "";

            MessageBox.Show("Данные очищены.");
        }
    }
}

