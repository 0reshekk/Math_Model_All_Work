using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace Math_Model_All_Work
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
            InitializeCostMatrix();
            TaskBox.SelectionChanged += TaskBox_SelectionChanged;
            SmoModeBox.SelectionChanged += SmoModeBox_SelectionChanged;
            UpdateSmoModeUi();
            UpdatePanels();
        }

        private void InitializeCostMatrix()
        {
            var costRows = new List<CostRow>
            {
                new CostRow { V1 = 2, V2 = 3, V3 = 1, V4 = 4, V5 = 0 },
                new CostRow { V1 = 5, V2 = 4, V3 = 8, V4 = 6, V5 = 0 },
                new CostRow { V1 = 5, V2 = 6, V3 = 8, V4 = 7, V5 = 0 }
            };

            CostMatrix.ItemsSource = costRows;
        }

        private void TaskBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdatePanels();
            ResultText.Text = "";
        }

        private void SmoModeBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateSmoModeUi();
            ResultText.Text = "";
        }

        private void UpdatePanels()
        {
            SmoPanel.Visibility = Visibility.Collapsed;
            TransportPanel.Visibility = Visibility.Collapsed;
            SimplexPanel.Visibility = Visibility.Collapsed;

            switch (TaskBox.SelectedIndex)
            {
                case 0:
                    SmoPanel.Visibility = Visibility.Visible;
                    break;
                case 1:
                case 2:
                    TransportPanel.Visibility = Visibility.Visible;
                    break;
                case 3:
                    SimplexPanel.Visibility = Visibility.Visible;
                    break;
            }
        }

        private void UpdateSmoModeUi()
        {
            TserviceLabel.Text = "t обслуживания (сек):";
            NLabel.Text = "n (каналов):";
            KLabel.Text = "K (макс. размер системы):";

            switch (SmoModeBox.SelectedIndex)
            {
                case 0:
                    SmoHintText.Text =
                        "Универсальный режим M/M/n/K";
                    break;
                case 1:
                    SmoHintText.Text =
                        "M/M/1/1: одноканальная система с отказами. Поля n и K !Не участвуют!";
                    break;
                case 2:
                    SmoHintText.Text =
                        "M/M/n/n: многоканальная система с отказами по формуле Эрланга B. Используются λ, μ и n.";
                    break;
                case 3:
                    SmoHintText.Text =
                        "M/M/1: одноканальная система с бесконечной очередью. Используются λ и μ; требуется ρ < 1.";
                    break;
                case 4:
                    SmoHintText.Text =
                        "M/M/n: многоканальная система с бесконечной очередью по формуле Эрланга C. Используются λ, μ и n.";
                    break;
                case 5:
                    SmoHintText.Text =
                        "M/D/1: одноканальная система с детерминированным временем обслуживания. Если t > 0, то μ берется как 1/t.";
                    break;
                case 6:
                    TserviceLabel.Text = "t ожидания (сек):";
                    SmoHintText.Text =
                        "Нетерпеливые заявки: t трактуется как максимально допустимое время ожидания. " +
                        "Используются λ, μ и t ожидания; поля n и K в этом режиме не участвуют.";
                    break;
                default:
                    SmoHintText.Text = "";
                    break;
            }
        }

        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            ResultText.Text = "";
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(ResultText.Text))
                {
                    MessageBox.Show("Нет данных для экспорта.");
                    return;
                }

                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Текстовый файл (*.txt)|*.txt|CSV (*.csv)|*.csv",
                    FileName = "result.txt"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    File.WriteAllText(saveFileDialog.FileName, ResultText.Text, Encoding.UTF8);
                    MessageBox.Show("Экспорт выполнен.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка экспорта: " + ex.Message);
            }
        }

        private void CalcBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                switch (TaskBox.SelectedIndex)
                {
                    case 0:
                        SolveSmo();
                        break;
                    case 1:
                        SolveTransportNorthWest();
                        break;
                    case 2:
                        SolveTransportLeastCost();
                        break;
                    case 3:
                        SolveSimplexTask();
                        break;
                    default:
                        ResultText.Text = "Выбери тип задачи.";
                        break;
                }
            }
            catch (Exception ex)
            {
                ResultText.Text = "Ошибка: " + ex.Message;
            }
        }

        // =========================
        // СМО
        // =========================
        private void SolveSmo()
        {
            double lambda = ReadDouble(LambdaBox.Text);
            double muInput = ReadDouble(MuBox.Text);
            double timeInput = ReadDouble(TserviceBox.Text);
            int n = ReadInt(NBox.Text, min: 1);
            int K = ReadInt(KBox.Text, min: 1);

            if (lambda <= 0) throw new Exception("λ должно быть > 0.");
            if (timeInput < 0) throw new Exception("t должно быть ≥ 0.");

            var sb = new StringBuilder();
            sb.AppendLine("Тип: Система массового обслуживания");
            sb.AppendLine();

            switch (SmoModeBox.SelectedIndex)
            {
                case 0:
                    {
                        double mu = ResolveServiceRate(muInput, timeInput);
                        ValidateServiceRate(mu);

                        sb.AppendLine("Режим расчета: Универсальный M/M/n/K");
                        sb.AppendLine();
                        Print(sb, ("λ", lambda), ("μ", mu), ("n", n), ("K", K));
                        sb.AppendLine();
                        AppendFinite(lambda, mu, n, K, sb);
                        break;
                    }
                case 1:
                    {
                        double mu = ResolveServiceRate(muInput, timeInput);
                        ValidateServiceRate(mu);

                        sb.AppendLine("Режим расчета: Одноканальная с отказами (M/M/1/1)");
                        sb.AppendLine();
                        Print(sb, ("λ", lambda), ("μ", mu));
                        sb.AppendLine();
                        AppendSingleLoss(lambda, mu, sb);
                        break;
                    }
                case 2:
                    {
                        double mu = ResolveServiceRate(muInput, timeInput);
                        ValidateServiceRate(mu);

                        sb.AppendLine("Режим расчета: Многоканальная с отказами (M/M/n/n)");
                        sb.AppendLine();
                        Print(sb, ("λ", lambda), ("μ", mu), ("n", n));
                        sb.AppendLine();
                        AppendMultiLoss(lambda, mu, n, sb);
                        break;
                    }
                case 3:
                    {
                        double mu = ResolveServiceRate(muInput, timeInput);
                        ValidateServiceRate(mu);

                        sb.AppendLine("Режим расчета: Одноканальная с бесконечной очередью (M/M/1)");
                        sb.AppendLine();
                        Print(sb, ("λ", lambda), ("μ", mu));
                        sb.AppendLine();
                        AppendMM1(lambda, mu, sb);
                        break;
                    }
                case 4:
                    {
                        double mu = ResolveServiceRate(muInput, timeInput);
                        ValidateServiceRate(mu);

                        sb.AppendLine("Режим расчета: Многоканальная с бесконечной очередью (M/M/n)");
                        sb.AppendLine();
                        Print(sb, ("λ", lambda), ("μ", mu), ("n", n));
                        sb.AppendLine();
                        AppendMMn(lambda, mu, n, sb);
                        break;
                    }
                case 5:
                    {
                        double mu = ResolveServiceRate(muInput, timeInput);
                        ValidateServiceRate(mu);

                        sb.AppendLine("Режим расчета: Одноканальная с фиксированным временем обслуживания (M/D/1)");
                        sb.AppendLine();
                        Print(sb, ("λ", lambda), ("μ", mu), ("t", timeInput));
                        sb.AppendLine();
                        AppendFixed(lambda, mu, sb);
                        break;
                    }
                case 6:
                    {
                        ValidateServiceRate(muInput);
                        if (timeInput <= 0) throw new Exception("t ожидания должно быть > 0.");

                        sb.AppendLine("Режим расчета: Ограниченное время ожидания");
                        sb.AppendLine();
                        Print(sb, ("λ", lambda), ("μ", muInput), ("tожид", timeInput));
                        sb.AppendLine();
                        AppendImpatient(lambda, muInput, timeInput, sb);
                        break;
                    }
                default:
                    throw new Exception("Выбери режим СМО.");
            }

            ResultText.Text = sb.ToString().TrimEnd();
        }

        private static double ResolveServiceRate(double mu, double serviceTime)
        {
            return serviceTime > 0 ? 1.0 / serviceTime : mu;
        }

        private static void ValidateServiceRate(double mu)
        {
            if (mu <= 0) throw new Exception("μ должно быть > 0.");
        }

        private void AppendSingleLoss(double lambda, double mu, StringBuilder sb)
        {
            double p0 = mu / (lambda + mu);
            double p1 = lambda / (lambda + mu);

            double pBlock = p1;
            double q = 1 - pBlock;
            double a = lambda * q;
            double kBar = p1;

            Print(sb,
                ("p0", p0), ("p1", p1),
                ("Pотк", pBlock), ("Q", q), ("A", a), ("k̄", kBar));
        }

        private void AppendMultiLoss(double lambda, double mu, int n, StringBuilder sb)
        {
            double rho = lambda / mu;

            double p0 = 1.0 / Enumerable.Range(0, n + 1)
                                         .Select(k => Math.Pow(rho, k) / Fact(k))
                                         .Sum();

            double pBlock = Math.Pow(rho, n) / Fact(n) * p0;
            double q = 1 - pBlock;
            double a = lambda * q;
            double kBar = a / mu;

            Print(sb,
                ("ρ", rho), ("p0", p0), ("Pотк", pBlock),
                ("Q", q), ("A", a), ("k̄", kBar));
        }

        private void AppendMM1(double lambda, double mu, StringBuilder sb)
        {
            double rho = lambda / mu;
            if (rho >= 1.0) throw new Exception("Для M/M/1 требуется ρ < 1.");

            double p0 = 1 - rho;
            double lSys = rho / (1 - rho);
            double lQ = rho * rho / (1 - rho);
            double tSys = lSys / lambda;
            double tQ = lQ / lambda;
            double pBusy = rho;

            Print(sb,
                ("ρ", rho), ("p0", p0),
                ("Pотк", 0.0), ("Q", 1.0), ("A", lambda),
                ("Pзан", pBusy), ("Lоч", lQ), ("Lсист", lSys),
                ("Tоч", tQ), ("Tсист", tSys));
        }

        private void AppendMMn(double lambda, double mu, int n, StringBuilder sb)
        {
            double rho = lambda / mu;
            if (rho / n >= 1.0)
                throw new Exception("Для M/M/n требуется ρ/n < 1.");

            double sum = Enumerable.Range(0, n)
                                   .Select(k => Math.Pow(rho, k) / Fact(k))
                                   .Sum();
            double tail = Math.Pow(rho, n) / Fact(n) * (n / (n - rho));
            double p0 = 1.0 / (sum + tail);

            double pWait = Math.Pow(rho, n) / Fact(n) * (n / (n - rho)) * p0;
            double lQ = pWait * (rho / (n - rho));
            double lService = rho;
            double lSys = lQ + lService;
            double tQ = lQ / lambda;
            double tSys = lSys / lambda;

            Print(sb,
                ("ρ", rho), ("p0", p0), ("Pожид", pWait),
                ("Pотк", 0.0), ("Q", 1.0), ("A", lambda),
                ("k̄", lService), ("Lоч", lQ), ("Lсист", lSys),
                ("Tоч", tQ), ("Tсист", tSys));
        }

        private void AppendFinite(double lambda, double mu, int n, int K, StringBuilder sb)
        {
            double[] p = new double[K + 1];
            p[0] = 1.0;

            for (int k = 1; k <= K; k++)
            {
                double lambdaK = lambda;
                double muK = Math.Min(k, n) * mu;
                p[k] = p[k - 1] * (lambdaK / muK);
            }

            double norm = p.Sum();
            for (int k = 0; k <= K; k++)
                p[k] /= norm;

            double p0 = p[0];
            double pBlock = p[K];
            double q = 1 - pBlock;
            double a = lambda * q;
            double kBar = 0.0;
            double lSys = 0.0;

            for (int k = 0; k <= K; k++)
            {
                lSys += k * p[k];
                kBar += Math.Min(k, n) * p[k];
            }

            double lQ = lSys - kBar;
            double tSys = a > 0 ? lSys / a : 0.0;
            double tQ = a > 0 ? lQ / a : 0.0;

            Print(sb,
                ("p0", p0), ("Pотк", pBlock), ("Q", q), ("A", a),
                ("k̄", kBar), ("Lоч", lQ), ("Lсист", lSys),
                ("Tоч", tQ), ("Tсист", tSys));
        }

        private void AppendFixed(double lambda, double mu, StringBuilder sb)
        {
            double rho = lambda / mu;
            if (rho >= 1.0) throw new Exception("Для M/D/1 требуется ρ < 1.");

            double lQ = (lambda * lambda) / (2 * mu * (mu - lambda));
            double tQ = lQ / lambda;
            double lSys = lQ + lambda / mu;
            double tSys = tQ + 1.0 / mu;

            Print(sb,
                ("ρ", rho),
                ("Lоч", lQ), ("Tоч", tQ),
                ("Lсист", lSys), ("Tсист", tSys));
        }

        private void AppendImpatient(double lambda, double mu, double maxWaitTime, StringBuilder sb)
        {
            double rho = lambda / mu;
            if (rho >= 1.0) throw new Exception("Для расчета нетерпеливых заявок требуется ρ < 1.");

            double lQ = (rho * rho) / (1 - rho);
            double tQ = lQ / lambda;
            double pStay = Math.Exp(-tQ / maxWaitTime);
            double pOut = 1 - pStay;
            double q = pStay;
            double a = lambda * q;

            Print(sb,
                ("ρ", rho),
                ("Tоч", tQ),
                ("Pтерп", pStay),
                ("Pотк", pOut),
                ("Q", q),
                ("A", a));
        }

        private static long Fact(int n)
        {
            if (n < 2) return 1;

            long result = 1;
            for (int i = 2; i <= n; i++)
                result *= i;

            return result;
        }

        // =========================
        // Транспортная задача
        // =========================
        private void SolveTransportNorthWest()
        {
            int[] supply = ParseIntArray(SupplyInput.Text);
            int[] demand = ParseIntArray(DemandInput.Text);
            List<CostRow> costMatrix = (List<CostRow>)CostMatrix.ItemsSource;

            int[,] cost = GetCostMatrix(costMatrix, supply.Length, demand.Length);

            if (supply.Sum() != demand.Sum())
                throw new Exception("Сумма запасов должна быть равна сумме потребностей.");

            int[,] result = NorthWestCornerMethod((int[])supply.Clone(), (int[])demand.Clone(), cost);
            int totalCost = CalculateTransportCost(result, cost);

            ResultText.Text =
                "Тип: Транспортная задача\n" +
                "Метод: Северо-западный угол\n\n" +
                "План перевозок:\n" +
                MatrixToString(result) +
                "\nТранспортные издержки: " + totalCost;
        }

        private void SolveTransportLeastCost()
        {
            int[] supply = ParseIntArray(SupplyInput.Text);
            int[] demand = ParseIntArray(DemandInput.Text);
            List<CostRow> costMatrix = (List<CostRow>)CostMatrix.ItemsSource;

            int[,] cost = GetCostMatrix(costMatrix, supply.Length, demand.Length);

            if (supply.Sum() != demand.Sum())
                throw new Exception("Сумма запасов должна быть равна сумме потребностей.");

            int[,] result = LeastCostMethod((int[])supply.Clone(), (int[])demand.Clone(), cost);
            int totalCost = CalculateTransportCost(result, cost);

            ResultText.Text =
                "Тип: Транспортная задача\n" +
                "Метод: Минимальный элемент\n\n" +
                "План перевозок:\n" +
                MatrixToString(result) +
                "\nТранспортные издержки: " + totalCost;
        }

        private int[,] GetCostMatrix(List<CostRow> costMatrix, int rows, int cols)
        {
            if (rows > costMatrix.Count)
                throw new Exception("Количество поставщиков больше числа строк матрицы.");

            if (cols > 5)
                throw new Exception("В этой версии доступно максимум 5 столбцов потребителей.");

            int[,] cost = new int[rows, cols];

            for (int i = 0; i < rows; i++)
            {
                int[] rowValues = new int[]
                {
                    costMatrix[i].V1,
                    costMatrix[i].V2,
                    costMatrix[i].V3,
                    costMatrix[i].V4,
                    costMatrix[i].V5
                };

                for (int j = 0; j < cols; j++)
                    cost[i, j] = rowValues[j];
            }

            return cost;
        }

        private int[,] NorthWestCornerMethod(int[] supply, int[] demand, int[,] cost)
        {
            int rows = supply.Length;
            int cols = demand.Length;
            int[,] result = new int[rows, cols];

            int i = 0, j = 0;

            while (i < rows && j < cols)
            {
                int allocation = Math.Min(supply[i], demand[j]);
                result[i, j] = allocation;

                supply[i] -= allocation;
                demand[j] -= allocation;

                if (supply[i] == 0 && demand[j] == 0)
                {
                    i++;
                    j++;
                }
                else if (supply[i] == 0)
                {
                    i++;
                }
                else
                {
                    j++;
                }
            }

            return result;
        }

        private int[,] LeastCostMethod(int[] supply, int[] demand, int[,] cost)
        {
            int rows = supply.Length;
            int cols = demand.Length;
            int[,] result = new int[rows, cols];

            bool[] rowClosed = new bool[rows];
            bool[] colClosed = new bool[cols];

            while (supply.Sum() > 0 && demand.Sum() > 0)
            {
                int minCost = int.MaxValue;
                int minI = -1, minJ = -1;

                for (int i = 0; i < rows; i++)
                {
                    if (rowClosed[i]) continue;

                    for (int j = 0; j < cols; j++)
                    {
                        if (colClosed[j]) continue;

                        if (cost[i, j] < minCost)
                        {
                            minCost = cost[i, j];
                            minI = i;
                            minJ = j;
                        }
                    }
                }

                if (minI == -1 || minJ == -1)
                    break;

                int allocation = Math.Min(supply[minI], demand[minJ]);
                result[minI, minJ] = allocation;

                supply[minI] -= allocation;
                demand[minJ] -= allocation;

                if (supply[minI] == 0) rowClosed[minI] = true;
                if (demand[minJ] == 0) colClosed[minJ] = true;
            }

            return result;
        }

        private int CalculateTransportCost(int[,] result, int[,] cost)
        {
            int totalCost = 0;

            for (int i = 0; i < result.GetLength(0); i++)
                for (int j = 0; j < result.GetLength(1); j++)
                    totalCost += result[i, j] * cost[i, j];

            return totalCost;
        }

        private string MatrixToString(int[,] matrix)
        {
            var sb = new StringBuilder();

            for (int i = 0; i < matrix.GetLength(0); i++)
            {
                for (int j = 0; j < matrix.GetLength(1); j++)
                    sb.Append(matrix[i, j]).Append('\t');

                sb.AppendLine();
            }

            return sb.ToString();
        }

        // =========================
        // Симплекс
        // =========================
        private void SolveSimplexTask()
        {
            double[] c = ParseDoubleArraySpace(ObjectiveInput.Text);
            double[,] a = ParseMatrix(MatrixInput.Text);
            double[] b = ParseDoubleArraySpace(BInput.Text);

            int m = a.GetLength(0);
            int n = a.GetLength(1);

            if (c.Length != n)
                throw new Exception("Количество коэффициентов целевой функции должно совпадать с числом столбцов матрицы A.");

            if (b.Length != m)
                throw new Exception("Размер вектора b должен совпадать с числом строк матрицы A.");

            if (b.Any(x => x < 0))
                throw new Exception("Правая часть b должна быть неотрицательной.");

            var result = SolveSimplex(a, b, c);

            var sb = new StringBuilder();
            sb.AppendLine("Тип: Симплекс-метод");
            sb.AppendLine("Результат:");
            sb.AppendLine();

            for (int i = 0; i < result.solution.Length; i++)
                sb.AppendLine($"x{i + 1,-5} = {Math.Round(result.solution[i], 3):0.###}");

            sb.AppendLine($"Zmax  = {Math.Round(result.optimum, 3):0.###}");

            ResultText.Text = sb.ToString();
        }

        private (double[] solution, double optimum) SolveSimplex(double[,] A, double[] b, double[] c)
        {
            int m = A.GetLength(0);
            int n = A.GetLength(1);

            int rows = m + 1;
            int cols = n + m + 1;

            double[,] tableau = new double[rows, cols];

            for (int i = 0; i < m; i++)
            {
                for (int j = 0; j < n; j++)
                    tableau[i, j] = A[i, j];

                tableau[i, n + i] = 1;
                tableau[i, cols - 1] = b[i];
            }

            for (int j = 0; j < n; j++)
                tableau[m, j] = -c[j];

            while (true)
            {
                int pivotCol = -1;
                double minValue = 0;

                for (int j = 0; j < cols - 1; j++)
                {
                    if (tableau[m, j] < minValue)
                    {
                        minValue = tableau[m, j];
                        pivotCol = j;
                    }
                }

                if (pivotCol == -1)
                    break;

                int pivotRow = -1;
                double minRatio = double.MaxValue;

                for (int i = 0; i < m; i++)
                {
                    if (tableau[i, pivotCol] > 1e-9)
                    {
                        double ratio = tableau[i, cols - 1] / tableau[i, pivotCol];
                        if (ratio < minRatio)
                        {
                            minRatio = ratio;
                            pivotRow = i;
                        }
                    }
                }

                if (pivotRow == -1)
                    throw new Exception("Целевая функция не ограничена сверху.");

                Pivot(tableau, pivotRow, pivotCol);
            }

            double[] solution = new double[n];

            for (int j = 0; j < n; j++)
            {
                int oneRow = -1;
                bool isBasic = true;

                for (int i = 0; i < m; i++)
                {
                    double value = tableau[i, j];

                    if (Math.Abs(value - 1) < 1e-9)
                    {
                        if (oneRow == -1)
                            oneRow = i;
                        else
                        {
                            isBasic = false;
                            break;
                        }
                    }
                    else if (Math.Abs(value) > 1e-9)
                    {
                        isBasic = false;
                        break;
                    }
                }

                solution[j] = (isBasic && oneRow != -1) ? tableau[oneRow, cols - 1] : 0;
            }

            double optimum = tableau[m, cols - 1];
            return (solution, optimum);
        }

        private void Pivot(double[,] tableau, int pivotRow, int pivotCol)
        {
            int rows = tableau.GetLength(0);
            int cols = tableau.GetLength(1);

            double pivot = tableau[pivotRow, pivotCol];

            for (int j = 0; j < cols; j++)
                tableau[pivotRow, j] /= pivot;

            for (int i = 0; i < rows; i++)
            {
                if (i == pivotRow) continue;

                double factor = tableau[i, pivotCol];
                for (int j = 0; j < cols; j++)
                    tableau[i, j] -= factor * tableau[pivotRow, j];
            }
        }

        // =========================
        // Вспомогательное
        // =========================
        private static double ReadDouble(string s)
        {
            s = s.Replace(',', '.');
            if (!double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var v))
                throw new Exception($"Не число: \"{s}\".");
            return v;
        }

        private static int ReadInt(string s, int min = 0)
        {
            if (!int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v))
                throw new Exception($"Не целое число: \"{s}\".");
            if (v < min) throw new Exception($"Значение должно быть ≥ {min}.");
            return v;
        }

        private int[] ParseIntArray(string text)
        {
            return text.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries)
                       .Select(int.Parse)
                       .ToArray();
        }

        private double[] ParseDoubleArraySpace(string text)
        {
            return text.Split(new[] { ' ', ';', '\t', ',', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                       .Select(x => double.Parse(x.Replace(',', '.'), CultureInfo.InvariantCulture))
                       .ToArray();
        }

        private double[,] ParseMatrix(string text)
        {
            var lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length == 0)
                throw new Exception("Матрица пустая.");

            var first = lines[0].Split(new[] { ' ', ';', '\t', ',' }, StringSplitOptions.RemoveEmptyEntries);
            int rows = lines.Length;
            int cols = first.Length;

            double[,] matrix = new double[rows, cols];

            for (int i = 0; i < rows; i++)
            {
                var row = lines[i].Split(new[] { ' ', ';', '\t', ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (row.Length != cols)
                    throw new Exception("Во всех строках матрицы должно быть одинаковое число элементов.");

                for (int j = 0; j < cols; j++)
                    matrix[i, j] = double.Parse(row[j].Replace(',', '.'), CultureInfo.InvariantCulture);
            }

            return matrix;
        }

        private static void Print(StringBuilder sb, params (string name, double val)[] pairs)
        {
            foreach (var (name, val) in pairs)
                sb.AppendLine($"{name,-10} = {Math.Round(val, 3):0.###}");
        }
    }

    public class CostRow
    {
        public int V1 { get; set; }
        public int V2 { get; set; }
        public int V3 { get; set; }
        public int V4 { get; set; }
        public int V5 { get; set; }
    }
}
