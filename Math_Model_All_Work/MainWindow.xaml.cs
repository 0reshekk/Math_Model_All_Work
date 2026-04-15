using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Xml.Linq;

namespace Math_Model_All_Work
{
    public partial class MainWindow : Window
    {
        #region Template Defaults

        private const int DefaultTaskIndex = 0;
        private const int DefaultSmoModeIndex = 0;
        private const int DefaultTransportRowCount = 3;
        private const string DefaultLambda = "2";
        private const string DefaultMu = "3";
        private const string DefaultServiceTime = "0";
        private const string DefaultChannels = "1";
        private const string DefaultSystemCapacity = "2";
        private const string DefaultSupply = "30,40,20";
        private const string DefaultDemand = "20,30,25,15,0";
        private const string DefaultObjective = "3 5";
        private const string DefaultMatrix = "1 0\r\n0 2\r\n3 2";
        private const string DefaultB = "4 12 18";

        #endregion

        #region Initialization

        public MainWindow()
        {
            InitializeComponent();
            CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;

            HookUiEvents();
            ResetTemplateToDefaults();
        }

        private void HookUiEvents()
        {
            TaskBox.SelectionChanged += TaskBox_SelectionChanged;
            SmoModeBox.SelectionChanged += SmoModeBox_SelectionChanged;
        }

        private void ResetTemplateToDefaults()
        {
            TaskBox.SelectedIndex = DefaultTaskIndex;
            SmoModeBox.SelectedIndex = DefaultSmoModeIndex;

            LambdaBox.Text = DefaultLambda;
            MuBox.Text = DefaultMu;
            TserviceBox.Text = DefaultServiceTime;
            NBox.Text = DefaultChannels;
            KBox.Text = DefaultSystemCapacity;

            SupplyInput.Text = DefaultSupply;
            DemandInput.Text = DefaultDemand;
            ObjectiveInput.Text = DefaultObjective;
            MatrixInput.Text = DefaultMatrix;
            BInput.Text = DefaultB;

            SetCostMatrixRows(BuildDefaultCostRows());
            ResultText.Text = string.Empty;

            UpdatePanels();
            UpdateSmoModeUi();
        }

        private static List<CostRow> BuildDefaultCostRows()
        {
            return new List<CostRow>
            {
                new CostRow { V1 = 2, V2 = 3, V3 = 1, V4 = 4, V5 = 0 },
                new CostRow { V1 = 5, V2 = 4, V3 = 8, V4 = 6, V5 = 0 },
                new CostRow { V1 = 5, V2 = 6, V3 = 8, V4 = 7, V5 = 0 }
            };
        }

        private static List<CostRow> BuildEmptyCostRows(int rowCount = DefaultTransportRowCount)
        {
            int normalizedCount = Math.Max(rowCount, DefaultTransportRowCount);
            var rows = new List<CostRow>();

            for (int index = 0; index < normalizedCount; index++)
                rows.Add(new CostRow());

            return rows;
        }

        #endregion

        #region Common UI Events

        private void TaskBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdatePanels();
            ResultText.Text = string.Empty;
        }

        private void SmoModeBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateSmoModeUi();
            ResultText.Text = string.Empty;
        }

        #endregion

        #region Common Template Helpers

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
                    SmoHintText.Text = "Универсальный режим M/M/n/K.";
                    break;
                case 1:
                    SmoHintText.Text = "M/M/1/1: одноканальная система с отказами. Поля n и K не участвуют.";
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
                    SmoHintText.Text = string.Empty;
                    break;
            }
        }

        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            ClearTemplateFields();
        }

        private void ClearTemplateFields()
        {
            LambdaBox.Text = string.Empty;
            MuBox.Text = string.Empty;
            TserviceBox.Text = string.Empty;
            NBox.Text = string.Empty;
            KBox.Text = string.Empty;

            SupplyInput.Text = string.Empty;
            DemandInput.Text = string.Empty;
            ObjectiveInput.Text = string.Empty;
            MatrixInput.Text = string.Empty;
            BInput.Text = string.Empty;

            SetCostMatrixRows(BuildEmptyCostRows(GetCurrentCostRowCount()));
            ResultText.Text = string.Empty;

            UpdateSmoModeUi();
        }

        private AppTemplateData CaptureTemplateData()
        {
            return new AppTemplateData
            {
                TaskIndex = Math.Max(TaskBox.SelectedIndex, 0),
                TaskName = GetSelectedItemText(TaskBox),
                SmoModeIndex = Math.Max(SmoModeBox.SelectedIndex, 0),
                SmoModeName = GetSelectedItemText(SmoModeBox),
                Lambda = LambdaBox.Text,
                Mu = MuBox.Text,
                ServiceTime = TserviceBox.Text,
                Channels = NBox.Text,
                SystemCapacity = KBox.Text,
                Supply = SupplyInput.Text,
                Demand = DemandInput.Text,
                Objective = ObjectiveInput.Text,
                Matrix = MatrixInput.Text,
                B = BInput.Text,
                ResultText = ResultText.Text,
                CostRows = GetCostMatrixRows()
            };
        }

        private void ApplyTemplateData(AppTemplateData data)
        {
            if (data == null)
                throw new ArgumentNullException(nameof(data));

            TaskBox.SelectedIndex = NormalizeIndex(data.TaskIndex, TaskBox.Items.Count, DefaultTaskIndex);
            SmoModeBox.SelectedIndex = NormalizeIndex(data.SmoModeIndex, SmoModeBox.Items.Count, DefaultSmoModeIndex);

            LambdaBox.Text = data.Lambda ?? string.Empty;
            MuBox.Text = data.Mu ?? string.Empty;
            TserviceBox.Text = data.ServiceTime ?? string.Empty;
            NBox.Text = data.Channels ?? string.Empty;
            KBox.Text = data.SystemCapacity ?? string.Empty;

            SupplyInput.Text = data.Supply ?? string.Empty;
            DemandInput.Text = data.Demand ?? string.Empty;
            ObjectiveInput.Text = data.Objective ?? string.Empty;
            MatrixInput.Text = data.Matrix ?? string.Empty;
            BInput.Text = data.B ?? string.Empty;

            SetCostMatrixRows(data.CostRows.Count > 0 ? data.CostRows : BuildEmptyCostRows());
            ResultText.Text = data.ResultText ?? string.Empty;

            UpdatePanels();
            UpdateSmoModeUi();
        }

        private static int NormalizeIndex(int value, int itemCount, int fallback)
        {
            return value >= 0 && value < itemCount ? value : fallback;
        }

        private static string GetSelectedItemText(ComboBox comboBox)
        {
            return (comboBox.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? string.Empty;
        }

        private int GetCurrentCostRowCount()
        {
            var rows = CostMatrix.ItemsSource as IEnumerable<CostRow>;
            int count = rows?.Count() ?? 0;
            return Math.Max(count, DefaultTransportRowCount);
        }

        private List<CostRow> GetCostMatrixRows()
        {
            var rows = CostMatrix.ItemsSource as IEnumerable<CostRow>;
            if (rows == null)
                return BuildDefaultCostRows();

            return rows.Select(row => row.Clone()).ToList();
        }

        private void SetCostMatrixRows(IEnumerable<CostRow> rows)
        {
            CostMatrix.ItemsSource = rows.Select(row => row.Clone()).ToList();
        }

        #endregion

        #region Import / Export

        private void ImportExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var openFileDialog = new OpenFileDialog
                {
                    Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                    Title = "Импорт параметров из Excel"
                };

                if (openFileDialog.ShowDialog() != true)
                    return;

                ApplyTemplateData(ExcelTemplateService.Import(openFileDialog.FileName));
                MessageBox.Show("Импорт Excel выполнен.", "Импорт", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка импорта Excel: " + ex.Message, "Импорт", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Workbook (*.xlsx)|*.xlsx|Текстовый файл (*.txt)|*.txt|CSV (*.csv)|*.csv",
                    DefaultExt = "xlsx",
                    FileName = "math-model-template.xlsx",
                    Title = "Экспорт данных"
                };

                if (saveFileDialog.ShowDialog() != true)
                    return;

                string extension = Path.GetExtension(saveFileDialog.FileName).ToLowerInvariant();
                if (extension == ".xlsx")
                {
                    ExcelTemplateService.Export(saveFileDialog.FileName, CaptureTemplateData());
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(ResultText.Text))
                        throw new Exception("Нет данных для текстового экспорта. Сначала выполните расчет или выберите Excel.");

                    File.WriteAllText(saveFileDialog.FileName, ResultText.Text, Encoding.UTF8);
                }

                MessageBox.Show("Экспорт выполнен.", "Экспорт", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка экспорта: " + ex.Message, "Экспорт", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region Calculation Dispatcher

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
                        ResultText.Text = "Выберите тип задачи.";
                        break;
                }
            }
            catch (Exception ex)
            {
                ResultText.Text = "Ошибка: " + ex.Message;
            }
        }

        #endregion

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

        private static int[] ParseIntArray(string text)
        {
            return text.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries)
                       .Select(part => ReadInt(part.Trim()))
                       .ToArray();
        }

        private static double[] ParseDoubleArraySpace(string text)
        {
            return text.Split(new[] { ' ', ';', '\t', ',', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                       .Select(ReadDouble)
                       .ToArray();
        }

        private static double[,] ParseMatrix(string text)
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
                    matrix[i, j] = ReadDouble(row[j]);
            }

            return matrix;
        }

        private static void Print(StringBuilder sb, params (string name, double val)[] pairs)
        {
            foreach (var (name, val) in pairs)
                sb.AppendLine($"{name,-10} = {Math.Round(val, 3):0.###}");
        }
    }

    internal sealed class AppTemplateData
    {
        public int TaskIndex { get; set; }
        public string TaskName { get; set; } = string.Empty;
        public int SmoModeIndex { get; set; }
        public string SmoModeName { get; set; } = string.Empty;
        public string Lambda { get; set; } = string.Empty;
        public string Mu { get; set; } = string.Empty;
        public string ServiceTime { get; set; } = string.Empty;
        public string Channels { get; set; } = string.Empty;
        public string SystemCapacity { get; set; } = string.Empty;
        public string Supply { get; set; } = string.Empty;
        public string Demand { get; set; } = string.Empty;
        public string Objective { get; set; } = string.Empty;
        public string Matrix { get; set; } = string.Empty;
        public string B { get; set; } = string.Empty;
        public string ResultText { get; set; } = string.Empty;
        public List<CostRow> CostRows { get; set; } = new List<CostRow>();
    }

    public sealed class CostRow
    {
        public int V1 { get; set; }
        public int V2 { get; set; }
        public int V3 { get; set; }
        public int V4 { get; set; }
        public int V5 { get; set; }

        public CostRow Clone()
        {
            return new CostRow
            {
                V1 = V1,
                V2 = V2,
                V3 = V3,
                V4 = V4,
                V5 = V5
            };
        }
    }

    internal static class ExcelTemplateService
    {
        private const string InputSheetName = "Ввод";
        private const string ResultSheetName = "Решение";
        private const string LegacyMetaSheetName = "Meta";
        private const string LegacyInputsSheetName = "Inputs";
        private const string LegacyTransportSheetName = "TransportCost";
        private const string LegacyResultSheetName = "Result";

        private static readonly XNamespace SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static readonly XNamespace RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace PackageRelationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
        private static readonly XNamespace ContentTypeNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";
        private static readonly XNamespace CorePropertyNamespace = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        private static readonly XNamespace DcNamespace = "http://purl.org/dc/elements/1.1/";
        private static readonly XNamespace DctermsNamespace = "http://purl.org/dc/terms/";
        private static readonly XNamespace DcmitypeNamespace = "http://purl.org/dc/dcmitype/";
        private static readonly XNamespace XsiNamespace = "http://www.w3.org/2001/XMLSchema-instance";
        private static readonly XNamespace ExtendedPropertyNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

        public static void Export(string filePath, AppTemplateData data)
        {
            if (data == null)
                throw new ArgumentNullException(nameof(data));

            if (File.Exists(filePath))
                File.Delete(filePath);

            using (var archive = ZipFile.Open(filePath, ZipArchiveMode.Create))
            {
                WriteEntry(archive, "[Content_Types].xml", BuildContentTypesXml());
                WriteEntry(archive, "_rels/.rels", BuildRootRelationshipsXml());
                WriteEntry(archive, "docProps/core.xml", BuildCorePropertiesXml());
                WriteEntry(archive, "docProps/app.xml", BuildAppPropertiesXml());
                WriteEntry(archive, "xl/workbook.xml", BuildWorkbookXml());
                WriteEntry(archive, "xl/_rels/workbook.xml.rels", BuildWorkbookRelationshipsXml());
                WriteEntry(archive, "xl/worksheets/sheet1.xml", BuildWorksheetXml(BuildInputRows(data)));
                WriteEntry(archive, "xl/worksheets/sheet2.xml", BuildWorksheetXml(BuildResultRows(data)));
            }
        }

        public static AppTemplateData Import(string filePath)
        {
            using (var archive = ZipFile.OpenRead(filePath))
            {
                Dictionary<string, string> workbookSheets = GetWorkbookSheetMap(archive);
                IReadOnlyList<string> sharedStrings = GetSharedStrings(archive);

                if (workbookSheets.ContainsKey(InputSheetName))
                    return ImportReadableWorkbook(archive, workbookSheets, sharedStrings);

                return ImportLegacyWorkbook(archive, workbookSheets, sharedStrings);
            }
        }

        private static string BuildContentTypesXml()
        {
            var document = new XDocument(
                new XElement(ContentTypeNamespace + "Types",
                    new XElement(ContentTypeNamespace + "Default",
                        new XAttribute("Extension", "rels"),
                        new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
                    new XElement(ContentTypeNamespace + "Default",
                        new XAttribute("Extension", "xml"),
                        new XAttribute("ContentType", "application/xml")),
                    new XElement(ContentTypeNamespace + "Override",
                        new XAttribute("PartName", "/xl/workbook.xml"),
                        new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")),
                    new XElement(ContentTypeNamespace + "Override",
                        new XAttribute("PartName", "/xl/worksheets/sheet1.xml"),
                        new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")),
                    new XElement(ContentTypeNamespace + "Override",
                        new XAttribute("PartName", "/xl/worksheets/sheet2.xml"),
                        new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")),
                    new XElement(ContentTypeNamespace + "Override",
                        new XAttribute("PartName", "/docProps/core.xml"),
                        new XAttribute("ContentType", "application/vnd.openxmlformats-package.core-properties+xml")),
                    new XElement(ContentTypeNamespace + "Override",
                        new XAttribute("PartName", "/docProps/app.xml"),
                        new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.extended-properties+xml"))));

            return document.Declaration + Environment.NewLine + document;
        }

        private static string BuildRootRelationshipsXml()
        {
            var document = new XDocument(
                new XElement(PackageRelationshipNamespace + "Relationships",
                    new XElement(PackageRelationshipNamespace + "Relationship",
                        new XAttribute("Id", "rId1"),
                        new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
                        new XAttribute("Target", "xl/workbook.xml")),
                    new XElement(PackageRelationshipNamespace + "Relationship",
                        new XAttribute("Id", "rId2"),
                        new XAttribute("Type", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"),
                        new XAttribute("Target", "docProps/core.xml")),
                    new XElement(PackageRelationshipNamespace + "Relationship",
                        new XAttribute("Id", "rId3"),
                        new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"),
                        new XAttribute("Target", "docProps/app.xml"))));

            return document.Declaration + Environment.NewLine + document;
        }

        private static string BuildCorePropertiesXml()
        {
            string timestamp = DateTime.UtcNow.ToString("s", CultureInfo.InvariantCulture) + "Z";

            var document = new XDocument(
                new XElement(CorePropertyNamespace + "coreProperties",
                    new XAttribute(XNamespace.Xmlns + "cp", CorePropertyNamespace),
                    new XAttribute(XNamespace.Xmlns + "dc", DcNamespace),
                    new XAttribute(XNamespace.Xmlns + "dcterms", DctermsNamespace),
                    new XAttribute(XNamespace.Xmlns + "dcmitype", DcmitypeNamespace),
                    new XAttribute(XNamespace.Xmlns + "xsi", XsiNamespace),
                    new XElement(DcNamespace + "creator", "Math_Model_All_Work"),
                    new XElement(CorePropertyNamespace + "lastModifiedBy", "Math_Model_All_Work"),
                    new XElement(DctermsNamespace + "created",
                        new XAttribute(XsiNamespace + "type", "dcterms:W3CDTF"),
                        timestamp),
                    new XElement(DctermsNamespace + "modified",
                        new XAttribute(XsiNamespace + "type", "dcterms:W3CDTF"),
                        timestamp)));

            return document.Declaration + Environment.NewLine + document;
        }

        private static string BuildAppPropertiesXml()
        {
            var document = new XDocument(
                new XElement(ExtendedPropertyNamespace + "Properties",
                    new XAttribute(XNamespace.Xmlns + "vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"),
                    new XElement(ExtendedPropertyNamespace + "Application", "Math_Model_All_Work"),
                    new XElement(ExtendedPropertyNamespace + "DocSecurity", "0"),
                    new XElement(ExtendedPropertyNamespace + "ScaleCrop", "false"),
                    new XElement(ExtendedPropertyNamespace + "HeadingPairs"),
                    new XElement(ExtendedPropertyNamespace + "TitlesOfParts"),
                    new XElement(ExtendedPropertyNamespace + "Company", string.Empty),
                    new XElement(ExtendedPropertyNamespace + "LinksUpToDate", "false"),
                    new XElement(ExtendedPropertyNamespace + "SharedDoc", "false"),
                    new XElement(ExtendedPropertyNamespace + "HyperlinksChanged", "false"),
                    new XElement(ExtendedPropertyNamespace + "AppVersion", "1.0")));

            return document.Declaration + Environment.NewLine + document;
        }

        private static string BuildWorkbookXml()
        {
            var document = new XDocument(
                new XElement(SpreadsheetNamespace + "workbook",
                    new XAttribute(XNamespace.Xmlns + "r", RelationshipNamespace),
                    new XElement(SpreadsheetNamespace + "sheets",
                        new XElement(SpreadsheetNamespace + "sheet",
                            new XAttribute("name", InputSheetName),
                            new XAttribute("sheetId", "1"),
                            new XAttribute(RelationshipNamespace + "id", "rId1")),
                        new XElement(SpreadsheetNamespace + "sheet",
                            new XAttribute("name", ResultSheetName),
                            new XAttribute("sheetId", "2"),
                            new XAttribute(RelationshipNamespace + "id", "rId2")))));

            return document.Declaration + Environment.NewLine + document;
        }

        private static string BuildWorkbookRelationshipsXml()
        {
            var document = new XDocument(
                new XElement(PackageRelationshipNamespace + "Relationships",
                    new XElement(PackageRelationshipNamespace + "Relationship",
                        new XAttribute("Id", "rId1"),
                        new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"),
                        new XAttribute("Target", "worksheets/sheet1.xml")),
                    new XElement(PackageRelationshipNamespace + "Relationship",
                        new XAttribute("Id", "rId2"),
                        new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"),
                        new XAttribute("Target", "worksheets/sheet2.xml"))));

            return document.Declaration + Environment.NewLine + document;
        }

        private static string BuildWorksheetXml(IReadOnlyList<IReadOnlyList<string>> rows)
        {
            var sheetData = new XElement(SpreadsheetNamespace + "sheetData");

            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                var rowElement = new XElement(SpreadsheetNamespace + "row",
                    new XAttribute("r", rowIndex + 1));

                IReadOnlyList<string> row = rows[rowIndex];
                for (int columnIndex = 0; columnIndex < row.Count; columnIndex++)
                {
                    string cellReference = GetCellReference(columnIndex + 1, rowIndex + 1);
                    string cellValue = row[columnIndex] ?? string.Empty;

                    rowElement.Add(
                        new XElement(SpreadsheetNamespace + "c",
                            new XAttribute("r", cellReference),
                            new XAttribute("t", "inlineStr"),
                            new XElement(SpreadsheetNamespace + "is",
                                new XElement(SpreadsheetNamespace + "t",
                                    new XAttribute(XNamespace.Xml + "space", "preserve"),
                                    cellValue))));
                }

                sheetData.Add(rowElement);
            }

            var document = new XDocument(
                new XElement(SpreadsheetNamespace + "worksheet",
                    sheetData));

            return document.Declaration + Environment.NewLine + document;
        }

        private static List<IReadOnlyList<string>> BuildInputRows(AppTemplateData data)
        {
            var costRows = data.CostRows?.Count > 0
                ? data.CostRows
                : new List<CostRow> { new CostRow(), new CostRow(), new CostRow() };

            int matrixStartRow = 19;
            int simplexHeaderRow = matrixStartRow + costRows.Count + 1;
            var rows = new List<IReadOnlyList<string>>
            {
                new[] { "Шаблон задач математического моделирования" },
                new[] { string.Empty },
                new[] { "Тип задачи", data.TaskIndex.ToString(CultureInfo.InvariantCulture), data.TaskName ?? string.Empty },
                new[] { "Режим СМО", data.SmoModeIndex.ToString(CultureInfo.InvariantCulture), data.SmoModeName ?? string.Empty },
                new[] { string.Empty },
                new[] { "Параметры СМО", "Значение" },
                new[] { "λ (интенсивность заявок)", data.Lambda ?? string.Empty },
                new[] { "μ (интенсивность обслуживания)", data.Mu ?? string.Empty },
                new[] { "t (время)", data.ServiceTime ?? string.Empty },
                new[] { "n (каналы)", data.Channels ?? string.Empty },
                new[] { "K (размер системы)", data.SystemCapacity ?? string.Empty },
                new[] { string.Empty },
                new[] { "Транспортная задача", "Значение" },
                new[] { "Запасы", data.Supply ?? string.Empty },
                new[] { "Потребности", data.Demand ?? string.Empty },
                new[] { "Количество строк матрицы", costRows.Count.ToString(CultureInfo.InvariantCulture) },
                new[] { string.Empty },
                new[] { "B1", "B2", "B3", "B4", "B5" }
            };

            foreach (var row in costRows)
            {
                rows.Add(new[]
                {
                    row.V1.ToString(CultureInfo.InvariantCulture),
                    row.V2.ToString(CultureInfo.InvariantCulture),
                    row.V3.ToString(CultureInfo.InvariantCulture),
                    row.V4.ToString(CultureInfo.InvariantCulture),
                    row.V5.ToString(CultureInfo.InvariantCulture)
                });
            }

            while (rows.Count < simplexHeaderRow - 1)
                rows.Add(new[] { string.Empty });

            rows.Add(new[] { "Симплекс-метод", "Значение" });
            rows.Add(new[] { "Целевая функция", data.Objective ?? string.Empty });
            rows.Add(new[] { "Матрица A", data.Matrix ?? string.Empty });
            rows.Add(new[] { "Правая часть b", data.B ?? string.Empty });

            return rows;
        }

        private static List<IReadOnlyList<string>> BuildResultRows(AppTemplateData data)
        {
            var rows = new List<IReadOnlyList<string>>
            {
                new[] { "Итоговое решение" },
                new[] { string.Empty },
                new[] { "Тип задачи", data.TaskName ?? string.Empty },
                new[] { "Режим СМО", data.SmoModeName ?? string.Empty },
                new[] { string.Empty },
                new[] { "Результат расчета" }
            };

            string[] lines = (data.ResultText ?? string.Empty)
                .Replace("\r\n", "\n")
                .Replace('\r', '\n')
                .Split('\n');

            foreach (string line in lines)
                rows.Add(new[] { line });

            return rows;
        }

        private static AppTemplateData ImportReadableWorkbook(
            ZipArchive archive,
            IReadOnlyDictionary<string, string> workbookSheets,
            IReadOnlyList<string> sharedStrings)
        {
            List<Dictionary<int, string>> inputRows = ReadWorksheetRows(archive, workbookSheets, InputSheetName, sharedStrings);
            int matrixRowCount = ReadWorksheetInt(inputRows, 16, 2, 3);
            int matrixStartRow = 19;
            int simplexHeaderRow = matrixStartRow + matrixRowCount + 1;

            var costRows = new List<CostRow>();
            for (int offset = 0; offset < matrixRowCount; offset++)
            {
                int rowNumber = matrixStartRow + offset;
                costRows.Add(new CostRow
                {
                    V1 = ReadWorksheetInt(inputRows, rowNumber, 1, 0),
                    V2 = ReadWorksheetInt(inputRows, rowNumber, 2, 0),
                    V3 = ReadWorksheetInt(inputRows, rowNumber, 3, 0),
                    V4 = ReadWorksheetInt(inputRows, rowNumber, 4, 0),
                    V5 = ReadWorksheetInt(inputRows, rowNumber, 5, 0)
                });
            }

            return new AppTemplateData
            {
                TaskIndex = ReadWorksheetInt(inputRows, 3, 2, 0),
                TaskName = ReadWorksheetString(inputRows, 3, 3),
                SmoModeIndex = ReadWorksheetInt(inputRows, 4, 2, 0),
                SmoModeName = ReadWorksheetString(inputRows, 4, 3),
                Lambda = ReadWorksheetString(inputRows, 7, 2),
                Mu = ReadWorksheetString(inputRows, 8, 2),
                ServiceTime = ReadWorksheetString(inputRows, 9, 2),
                Channels = ReadWorksheetString(inputRows, 10, 2),
                SystemCapacity = ReadWorksheetString(inputRows, 11, 2),
                Supply = ReadWorksheetString(inputRows, 14, 2),
                Demand = ReadWorksheetString(inputRows, 15, 2),
                Objective = ReadWorksheetString(inputRows, simplexHeaderRow + 1, 2),
                Matrix = ReadWorksheetString(inputRows, simplexHeaderRow + 2, 2),
                B = ReadWorksheetString(inputRows, simplexHeaderRow + 3, 2),
                ResultText = ReadReadableResultText(archive, workbookSheets, sharedStrings),
                CostRows = costRows
            };
        }

        private static AppTemplateData ImportLegacyWorkbook(
            ZipArchive archive,
            IReadOnlyDictionary<string, string> workbookSheets,
            IReadOnlyList<string> sharedStrings)
        {
            Dictionary<string, string> meta = ReadKeyValueSheet(archive, workbookSheets, LegacyMetaSheetName, sharedStrings);
            Dictionary<string, string> inputs = ReadKeyValueSheet(archive, workbookSheets, LegacyInputsSheetName, sharedStrings);
            List<CostRow> costRows = ReadTransportRows(archive, workbookSheets, sharedStrings);
            string resultText = ReadResultText(archive, workbookSheets, sharedStrings);

            return new AppTemplateData
            {
                TaskIndex = ReadInt(meta, "TaskIndex", 0),
                TaskName = ReadString(meta, "TaskName"),
                SmoModeIndex = ReadInt(meta, "SmoModeIndex", 0),
                SmoModeName = ReadString(meta, "SmoModeName"),
                Lambda = ReadString(inputs, "Lambda"),
                Mu = ReadString(inputs, "Mu"),
                ServiceTime = ReadString(inputs, "ServiceTime"),
                Channels = ReadString(inputs, "Channels"),
                SystemCapacity = ReadString(inputs, "SystemCapacity"),
                Supply = ReadString(inputs, "Supply"),
                Demand = ReadString(inputs, "Demand"),
                Objective = ReadString(inputs, "Objective"),
                Matrix = ReadString(inputs, "Matrix"),
                B = ReadString(inputs, "B"),
                ResultText = resultText,
                CostRows = costRows
            };
        }

        private static void WriteEntry(ZipArchive archive, string entryPath, string content)
        {
            var entry = archive.CreateEntry(entryPath, CompressionLevel.Optimal);
            using (var writer = new StreamWriter(entry.Open()))
                writer.Write(content);
        }

        private static Dictionary<string, string> GetWorkbookSheetMap(ZipArchive archive)
        {
            XDocument workbook = LoadXml(archive, "xl/workbook.xml");
            XDocument relationships = LoadXml(archive, "xl/_rels/workbook.xml.rels");

            Dictionary<string, string> relationMap = relationships.Root?
                .Elements(PackageRelationshipNamespace + "Relationship")
                .ToDictionary(
                    element => (string)element.Attribute("Id"),
                    element => NormalizeWorkbookTarget((string)element.Attribute("Target")))
                ?? new Dictionary<string, string>();

            var sheetMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            IEnumerable<XElement> sheets = workbook.Root?
                .Element(SpreadsheetNamespace + "sheets")?
                .Elements(SpreadsheetNamespace + "sheet")
                ?? Enumerable.Empty<XElement>();

            foreach (XElement sheet in sheets)
            {
                string sheetName = (string)sheet.Attribute("name");
                string relationId = (string)sheet.Attribute(RelationshipNamespace + "id");
                if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(relationId))
                    continue;

                if (relationMap.TryGetValue(relationId, out string target))
                    sheetMap[sheetName] = target;
            }

            return sheetMap;
        }

        private static IReadOnlyList<string> GetSharedStrings(ZipArchive archive)
        {
            var entry = archive.GetEntry("xl/sharedStrings.xml");
            if (entry == null)
                return Array.Empty<string>();

            using (var stream = entry.Open())
            {
                XDocument document = XDocument.Load(stream);
                return document.Root?
                    .Elements(SpreadsheetNamespace + "si")
                    .Select(ReadSharedStringItem)
                    .ToList()
                    ?? new List<string>();
            }
        }

        private static string ReadSharedStringItem(XElement item)
        {
            if (item == null)
                return string.Empty;

            var directText = item.Element(SpreadsheetNamespace + "t");
            if (directText != null)
                return directText.Value;

            return string.Concat(item.Elements(SpreadsheetNamespace + "r")
                                     .Select(run => run.Element(SpreadsheetNamespace + "t")?.Value ?? string.Empty));
        }

        private static Dictionary<string, string> ReadKeyValueSheet(
            ZipArchive archive,
            IReadOnlyDictionary<string, string> workbookSheets,
            string sheetName,
            IReadOnlyList<string> sharedStrings)
        {
            List<Dictionary<int, string>> rows = ReadWorksheetRows(archive, workbookSheets, sheetName, sharedStrings);
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            for (int i = 1; i < rows.Count; i++)
            {
                string key = GetCell(rows[i], 1);
                string value = GetCell(rows[i], 2);

                if (!string.IsNullOrWhiteSpace(key))
                    result[key] = value ?? string.Empty;
            }

            return result;
        }

        private static List<CostRow> ReadTransportRows(
            ZipArchive archive,
            IReadOnlyDictionary<string, string> workbookSheets,
            IReadOnlyList<string> sharedStrings)
        {
            List<Dictionary<int, string>> rows = ReadWorksheetRows(archive, workbookSheets, LegacyTransportSheetName, sharedStrings);
            var result = new List<CostRow>();

            for (int i = 1; i < rows.Count; i++)
            {
                if (rows[i].Count == 0)
                    continue;

                result.Add(new CostRow
                {
                    V1 = ReadInt(rows[i], 1),
                    V2 = ReadInt(rows[i], 2),
                    V3 = ReadInt(rows[i], 3),
                    V4 = ReadInt(rows[i], 4),
                    V5 = ReadInt(rows[i], 5)
                });
            }

            return result;
        }

        private static string ReadResultText(
            ZipArchive archive,
            IReadOnlyDictionary<string, string> workbookSheets,
            IReadOnlyList<string> sharedStrings)
        {
            List<Dictionary<int, string>> rows = ReadWorksheetRows(archive, workbookSheets, LegacyResultSheetName, sharedStrings);
            var lines = new List<string>();

            for (int i = 1; i < rows.Count; i++)
                lines.Add(GetCell(rows[i], 1) ?? string.Empty);

            return string.Join(Environment.NewLine, lines);
        }

        private static string ReadReadableResultText(
            ZipArchive archive,
            IReadOnlyDictionary<string, string> workbookSheets,
            IReadOnlyList<string> sharedStrings)
        {
            List<Dictionary<int, string>> rows = ReadWorksheetRows(archive, workbookSheets, ResultSheetName, sharedStrings);
            var lines = new List<string>();

            for (int rowNumber = 7; rowNumber <= rows.Count; rowNumber++)
            {
                string value = ReadWorksheetString(rows, rowNumber, 1);
                if (!string.IsNullOrEmpty(value) || lines.Count > 0)
                    lines.Add(value);
            }

            return string.Join(Environment.NewLine, lines).TrimEnd();
        }

        private static string ReadWorksheetString(List<Dictionary<int, string>> rows, int rowNumber, int columnIndex)
        {
            if (rowNumber < 1 || rowNumber > rows.Count)
                return string.Empty;

            return GetCell(rows[rowNumber - 1], columnIndex) ?? string.Empty;
        }

        private static int ReadWorksheetInt(List<Dictionary<int, string>> rows, int rowNumber, int columnIndex, int defaultValue)
        {
            string value = ReadWorksheetString(rows, rowNumber, columnIndex);
            if (string.IsNullOrWhiteSpace(value))
                return defaultValue;

            if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedInt))
                return parsedInt;

            if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsedDouble))
                return Convert.ToInt32(Math.Round(parsedDouble, MidpointRounding.AwayFromZero));

            throw new InvalidDataException($"Не удалось прочитать целое число из Excel: \"{value}\".");
        }

        private static List<Dictionary<int, string>> ReadWorksheetRows(
            ZipArchive archive,
            IReadOnlyDictionary<string, string> workbookSheets,
            string sheetName,
            IReadOnlyList<string> sharedStrings)
        {
            if (!workbookSheets.TryGetValue(sheetName, out string sheetPath))
                throw new InvalidDataException($"В файле Excel не найден лист \"{sheetName}\".");

            XDocument worksheet = LoadXml(archive, sheetPath);
            IEnumerable<XElement> rowElements = worksheet.Root?
                .Element(SpreadsheetNamespace + "sheetData")?
                .Elements(SpreadsheetNamespace + "row")
                ?? Enumerable.Empty<XElement>();

            var rows = new List<Dictionary<int, string>>();
            foreach (XElement row in rowElements)
            {
                var cells = new Dictionary<int, string>();

                foreach (XElement cell in row.Elements(SpreadsheetNamespace + "c"))
                {
                    string reference = (string)cell.Attribute("r");
                    int columnIndex = GetColumnIndex(reference);
                    cells[columnIndex] = ReadCellValue(cell, sharedStrings);
                }

                rows.Add(cells);
            }

            return rows;
        }

        private static XDocument LoadXml(ZipArchive archive, string entryPath)
        {
            var entry = archive.GetEntry(entryPath);
            if (entry == null)
                throw new InvalidDataException($"В Excel-файле отсутствует часть {entryPath}.");

            using (var stream = entry.Open())
                return XDocument.Load(stream);
        }

        private static string NormalizeWorkbookTarget(string target)
        {
            string normalized = (target ?? string.Empty).Replace('\\', '/');
            if (normalized.StartsWith("/"))
                return normalized.TrimStart('/');

            return "xl/" + normalized.TrimStart('/');
        }

        private static string ReadCellValue(XElement cell, IReadOnlyList<string> sharedStrings)
        {
            string type = (string)cell.Attribute("t");

            if (type == "inlineStr")
            {
                return string.Concat(cell.Descendants(SpreadsheetNamespace + "t")
                                         .Select(element => element.Value));
            }

            string value = cell.Element(SpreadsheetNamespace + "v")?.Value ?? string.Empty;
            if (type == "s" && int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int sharedIndex))
            {
                if (sharedIndex >= 0 && sharedIndex < sharedStrings.Count)
                    return sharedStrings[sharedIndex];
            }

            if (type == "b")
                return value == "1" ? "TRUE" : "FALSE";

            return value;
        }

        private static int GetColumnIndex(string cellReference)
        {
            if (string.IsNullOrWhiteSpace(cellReference))
                return 1;

            int index = 0;
            foreach (char character in cellReference)
            {
                if (!char.IsLetter(character))
                    break;

                index = (index * 26) + (char.ToUpperInvariant(character) - 'A' + 1);
            }

            return index == 0 ? 1 : index;
        }

        private static string GetCell(Dictionary<int, string> row, int columnIndex)
        {
            return row.TryGetValue(columnIndex, out string value) ? value : string.Empty;
        }

        private static int ReadInt(Dictionary<int, string> row, int columnIndex)
        {
            string rawValue = GetCell(row, columnIndex);
            if (string.IsNullOrWhiteSpace(rawValue))
                return 0;

            if (int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value))
                return value;

            if (double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double doubleValue))
                return Convert.ToInt32(Math.Round(doubleValue, MidpointRounding.AwayFromZero));

            throw new InvalidDataException($"Не удалось прочитать целое число из Excel: \"{rawValue}\".");
        }

        private static int ReadInt(IReadOnlyDictionary<string, string> values, string key, int defaultValue)
        {
            if (!values.TryGetValue(key, out string rawValue) || string.IsNullOrWhiteSpace(rawValue))
                return defaultValue;

            if (int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int result))
                return result;

            throw new InvalidDataException($"Поле \"{key}\" содержит некорректное целое значение.");
        }

        private static string ReadString(IReadOnlyDictionary<string, string> values, string key)
        {
            return values.TryGetValue(key, out string rawValue) ? rawValue ?? string.Empty : string.Empty;
        }

        private static string GetCellReference(int columnIndex, int rowIndex)
        {
            return GetColumnName(columnIndex) + rowIndex.ToString(CultureInfo.InvariantCulture);
        }

        private static string GetColumnName(int columnIndex)
        {
            var characters = new Stack<char>();
            int value = columnIndex;

            while (value > 0)
            {
                value--;
                characters.Push((char)('A' + (value % 26)));
                value /= 26;
            }

            return new string(characters.ToArray());
        }
    }
}
