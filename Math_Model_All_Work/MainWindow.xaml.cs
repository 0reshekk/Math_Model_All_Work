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

        /// <summary>
        /// Создает главное окно, настраивает культуру и подготавливает интерфейс к работе.
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
            CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;

            HookUiEvents();
            ResetTemplateToDefaults();
        }

        /// <summary>
        /// Подписывает элементы интерфейса на события, которые должны обновлять шаблон на лету.
        /// </summary>
        private void HookUiEvents()
        {
            TaskBox.SelectionChanged += TaskBox_SelectionChanged;
            SmoModeBox.SelectionChanged += SmoModeBox_SelectionChanged;
        }

        /// <summary>
        /// Возвращает шаблон к начальному состоянию, чтобы пользователь всегда стартовал с корректного примера.
        /// </summary>
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

        /// <summary>
        /// Формирует демонстрационную матрицу стоимостей для транспортной задачи.
        /// </summary>
        private static List<CostRow> BuildDefaultCostRows()
        {
            return new List<CostRow>
            {
                new CostRow { V1 = 2, V2 = 3, V3 = 1, V4 = 4, V5 = 0 },
                new CostRow { V1 = 5, V2 = 4, V3 = 8, V4 = 6, V5 = 0 },
                new CostRow { V1 = 5, V2 = 6, V3 = 8, V4 = 7, V5 = 0 }
            };
        }

        /// <summary>
        /// Создает пустую матрицу стоимостей с минимальным количеством строк,
        /// чтобы таблица не исчезала полностью после очистки.
        /// </summary>
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

        // Реагирует на смену типа задачи, переключает нужные панели и очищает старый результат.
        private void TaskBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdatePanels();
            ResultText.Text = string.Empty;
        }

        // Реагирует на смену режима СМО, обновляет подсказки и очищает старый результат.
        private void SmoModeBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateSmoModeUi();
            ResultText.Text = string.Empty;
        }

        #endregion

        #region Common Template Helpers

        /// <summary>
        /// Показывает только ту панель ввода, которая относится к выбранному типу задачи.
        /// </summary>
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

        /// <summary>
        /// Обновляет подписи и пояснения для выбранного режима СМО,
        /// чтобы пользователь видел, какие поля реально участвуют в расчете.
        /// </summary>
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

        // Кнопка очистки: сбрасывает введенные пользователем значения, но оставляет структуру формы.
        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            ClearTemplateFields();
        }

        /// <summary>
        /// Очищает все поля ввода и сбрасывает результат, сохраняя рабочую структуру таблиц и панелей.
        /// </summary>
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

        /// <summary>
        /// Снимает текущие значения с интерфейса и сохраняет их в объекте,
        /// который затем используется при экспорте.
        /// </summary>
        private AppTemplateData CaptureTemplateData()
        {
            return new AppTemplateData
            {
                TaskIndex = Math.Max(TaskBox.SelectedIndex, 0),
                SmoModeIndex = Math.Max(SmoModeBox.SelectedIndex, 0),
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

        /// <summary>
        /// Применяет импортированные данные к интерфейсу и восстанавливает выбранный режим задачи.
        /// </summary>
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

        /// <summary>
        /// Возвращает индекс, если он попадает в диапазон элементов списка;
        /// иначе подставляет безопасное значение по умолчанию.
        /// </summary>
        private static int NormalizeIndex(int value, int itemCount, int fallback)
        {
            return value >= 0 && value < itemCount ? value : fallback;
        }

        /// <summary>
        /// Определяет текущее количество строк в матрице стоимостей,
        /// чтобы при очистке и импорте не потерять размер таблицы.
        /// </summary>
        private int GetCurrentCostRowCount()
        {
            var rows = CostMatrix.ItemsSource as IEnumerable<CostRow>;
            int count = rows?.Count() ?? 0;
            return Math.Max(count, DefaultTransportRowCount);
        }

        /// <summary>
        /// Копирует строки матрицы стоимостей из таблицы в обычный список,
        /// чтобы экспорт и расчеты работали с независимыми данными.
        /// </summary>
        private List<CostRow> GetCostMatrixRows()
        {
            var rows = CostMatrix.ItemsSource as IEnumerable<CostRow>;
            if (rows == null)
                return BuildDefaultCostRows();

            return rows.Select(row => row.Clone()).ToList();
        }

        /// <summary>
        /// Загружает набор строк в таблицу стоимостей через копию,
        /// чтобы изменения в UI не портили исходные объекты.
        /// </summary>
        private void SetCostMatrixRows(IEnumerable<CostRow> rows)
        {
            CostMatrix.ItemsSource = rows.Select(row => row.Clone()).ToList();
        }

        #endregion

        #region Import / Export

        // Кнопка импорта Excel: загружает сохраненный шаблон и восстанавливает его на форме.
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

        // Кнопка экспорта: сохраняет либо Excel-шаблон для будущего импорта, либо текстовый результат расчета.
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

        // Кнопка расчета: запускает нужный алгоритм в зависимости от выбранного типа задачи.
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
        #region // СМО
        // =========================
        /// <summary>
        /// Выполняет расчет для выбранного режима СМО и записывает в результат только итоговые вычисленные показатели.
        /// </summary>
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

            switch (SmoModeBox.SelectedIndex)
            {
                case 0:
                    {
                        double mu = ResolveServiceRate(muInput, timeInput);
                        ValidateServiceRate(mu);

                        AppendFinite(lambda, mu, n, K, sb);
                        break;
                    }
                case 1:
                    {
                        double mu = ResolveServiceRate(muInput, timeInput);
                        ValidateServiceRate(mu);

                        AppendSingleLoss(lambda, mu, sb);
                        break;
                    }
                case 2:
                    {
                        double mu = ResolveServiceRate(muInput, timeInput);
                        ValidateServiceRate(mu);

                        AppendMultiLoss(lambda, mu, n, sb);
                        break;
                    }
                case 3:
                    {
                        double mu = ResolveServiceRate(muInput, timeInput);
                        ValidateServiceRate(mu);

                        AppendMM1(lambda, mu, sb);
                        break;
                    }
                case 4:
                    {
                        double mu = ResolveServiceRate(muInput, timeInput);
                        ValidateServiceRate(mu);

                        AppendMMn(lambda, mu, n, sb);
                        break;
                    }
                case 5:
                    {
                        double mu = ResolveServiceRate(muInput, timeInput);
                        ValidateServiceRate(mu);

                        AppendFixed(lambda, mu, sb);
                        break;
                    }
                case 6:
                    {
                        ValidateServiceRate(muInput);
                        if (timeInput <= 0) throw new Exception("t ожидания должно быть > 0.");

                        AppendImpatient(lambda, muInput, timeInput, sb);
                        break;
                    }
                default:
                    throw new Exception("Выбери режим СМО.");
            }

            ResultText.Text = sb.ToString().TrimEnd();
        }

        /// <summary>
        /// Определяет интенсивность обслуживания:
        /// либо берет введенное μ, либо вычисляет его как 1 / t обслуживания.
        /// </summary>
        private static double ResolveServiceRate(double mu, double serviceTime)
        {
            return serviceTime > 0 ? 1.0 / serviceTime : mu;
        }

        /// <summary>
        /// Проверяет, что интенсивность обслуживания допустима для формул СМО.
        /// </summary>
        private static void ValidateServiceRate(double mu)
        {
            if (mu <= 0) throw new Exception("μ должно быть > 0.");
        }

        /// <summary>
        /// Добавляет показатели для одноканальной СМО с отказами M/M/1/1.
        /// </summary>
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

        /// <summary>
        /// Добавляет показатели для многоканальной СМО с отказами по формуле Эрланга B.
        /// </summary>
        private void AppendMultiLoss(double lambda, double mu, int n, StringBuilder sb)
        {
            double rho = lambda / mu;

            // Находим вероятность пустой системы через сумму вероятностей всех допустимых состояний.
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

        /// <summary>
        /// Добавляет показатели для одноканальной СМО с бесконечной очередью M/M/1.
        /// </summary>
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

        /// <summary>
        /// Добавляет показатели для многоканальной СМО с бесконечной очередью M/M/n.
        /// </summary>
        private void AppendMMn(double lambda, double mu, int n, StringBuilder sb)
        {
            double rho = lambda / mu;
            if (rho / n >= 1.0)
                throw new Exception("Для M/M/n требуется ρ/n < 1.");

            // Первая сумма учитывает состояния без очереди, хвостовой член — состояния ожидания.
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

        /// <summary>
        /// Добавляет показатели для конечной системы M/M/n/K,
        /// где число заявок в системе ограничено сверху.
        /// </summary>
        private void AppendFinite(double lambda, double mu, int n, int K, StringBuilder sb)
        {
            double[] p = new double[K + 1];
            p[0] = 1.0;

            // Рекуррентно восстанавливаем вероятности состояний до K заявок.
            for (int k = 1; k <= K; k++)
            {
                double lambdaK = lambda;
                double muK = Math.Min(k, n) * mu;
                p[k] = p[k - 1] * (lambdaK / muK);
            }

            // После нормировки массив p превращается в корректное распределение вероятностей.
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

        /// <summary>
        /// Добавляет показатели для модели M/D/1 с фиксированным временем обслуживания.
        /// </summary>
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

        /// <summary>
        /// Добавляет показатели для модели с ограниченным временем ожидания,
        /// где часть заявок покидает систему, не дождавшись обслуживания.
        /// </summary>
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

        /// <summary>
        /// Вычисляет факториал натурального числа для формул Эрланга.
        /// </summary>
        private static long Fact(int n)
        {
            if (n < 2) return 1;

            long result = 1;
            for (int i = 2; i <= n; i++)
                result *= i;

            return result;
        }
        #endregion

        // =========================
        #region // Транспортная задача
        // =========================
        /// <summary>
        /// Решает транспортную задачу методом северо-западного угла
        /// и показывает только итоговый план перевозок и суммарные издержки.
        /// </summary>
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
                "План перевозок:\n" +
                MatrixToString(result) +
                "\nТранспортные издержки: " + totalCost;
        }

        /// <summary>
        /// Решает транспортную задачу методом минимального элемента
        /// и показывает только итоговый план перевозок и суммарные издержки.
        /// </summary>
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
                "План перевозок:\n" +
                MatrixToString(result) +
                "\nТранспортные издержки: " + totalCost;
        }

        /// <summary>
        /// Преобразует табличный ввод стоимости перевозок в двумерную матрицу,
        /// с которой уже работают алгоритмы транспортной задачи.
        /// </summary>
        private int[,] GetCostMatrix(List<CostRow> costMatrix, int rows, int cols)
        {
            if (rows > costMatrix.Count)
                throw new Exception("Количество поставщиков больше числа строк матрицы.");

            if (cols > 5)
                throw new Exception("В этой версии доступно максимум 5 столбцов потребителей.");

            int[,] cost = new int[rows, cols];

            for (int i = 0; i < rows; i++)
            {
                // DataGrid хранит строку как объект с пятью полями,
                // поэтому сначала собираем их в обычный массив.
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

        /// <summary>
        /// Строит опорный план методом северо-западного угла:
        /// последовательно заполняет таблицу, начиная с левой верхней клетки.
        /// </summary>
        private int[,] NorthWestCornerMethod(int[] supply, int[] demand, int[,] cost)
        {
            int rows = supply.Length;
            int cols = demand.Length;
            int[,] result = new int[rows, cols];

            int i = 0, j = 0;

            while (i < rows && j < cols)
            {
                // В текущую клетку отправляем максимально возможный объем,
                // не превышая ни запас поставщика, ни спрос потребителя.
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

        /// <summary>
        /// Строит план перевозок методом минимального элемента:
        /// на каждом шаге выбирает самую дешевую еще доступную клетку.
        /// </summary>
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

                // Ищем клетку с минимальной стоимостью среди еще не закрытых строк и столбцов.
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

        /// <summary>
        /// Вычисляет суммарные транспортные издержки для готового плана перевозок.
        /// </summary>
        private int CalculateTransportCost(int[,] result, int[,] cost)
        {
            int totalCost = 0;

            for (int i = 0; i < result.GetLength(0); i++)
                for (int j = 0; j < result.GetLength(1); j++)
                    totalCost += result[i, j] * cost[i, j];

            return totalCost;
        }

        /// <summary>
        /// Преобразует числовую матрицу в многострочный текст,
        /// который удобно показывать в поле результата и экспортировать в Excel.
        /// </summary>
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

        #endregion

        // =========================
        #region // Симплекс
        // =========================
        /// <summary>
        /// Проверяет ввод симплекс-задачи, решает ее и выводит только найденный оптимальный план и значение целевой функции.
        /// </summary>
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

            for (int i = 0; i < result.solution.Length; i++)
                sb.AppendLine($"x{i + 1,-5} = {Math.Round(result.solution[i], 3):0.###}");

            sb.AppendLine($"Zmax  = {Math.Round(result.optimum, 3):0.###}");

            ResultText.Text = sb.ToString();
        }

        /// <summary>
        /// Решает задачу линейного программирования симплекс-методом
        /// в стандартной форме с ограничениями вида A * x <= b.
        /// </summary>
        private (double[] solution, double optimum) SolveSimplex(double[,] A, double[] b, double[] c)
        {
            int m = A.GetLength(0);
            int n = A.GetLength(1);

            int rows = m + 1;
            int cols = n + m + 1;

            double[,] tableau = new double[rows, cols];

            // Верхняя часть таблицы содержит матрицу ограничений,
            // добавленные базисные переменные и правую часть.
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

                // Опорный столбец выбираем по самому отрицательному коэффициенту в строке цели.
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

                // Опорная строка определяется по правилу минимального положительного отношения.
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

            // Восстанавливаем значения исходных переменных по базисным столбцам итоговой таблицы.
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

        /// <summary>
        /// Выполняет один шаг Жорданова исключения для выбранного разрешающего элемента.
        /// </summary>
        private void Pivot(double[,] tableau, int pivotRow, int pivotCol)
        {
            int rows = tableau.GetLength(0);
            int cols = tableau.GetLength(1);

            double pivot = tableau[pivotRow, pivotCol];

            // Сначала нормируем опорную строку, чтобы в разрешающем столбце получить единицу.
            for (int j = 0; j < cols; j++)
                tableau[pivotRow, j] /= pivot;

            // Затем зануляем остальные элементы разрешающего столбца.
            for (int i = 0; i < rows; i++)
            {
                if (i == pivotRow) continue;

                double factor = tableau[i, pivotCol];
                for (int j = 0; j < cols; j++)
                    tableau[i, j] -= factor * tableau[pivotRow, j];
            }
        }
        #endregion


        // =========================
        // Вспомогательное
        // =========================
        /// <summary>
        /// Читает вещественное число из текстового поля, поддерживая и точку, и запятую.
        /// </summary>
        private static double ReadDouble(string s)
        {
            s = s.Replace(',', '.');
            if (!double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var v))
                throw new Exception($"Не число: \"{s}\".");
            return v;
        }

        /// <summary>
        /// Читает целое число и дополнительно проверяет нижнюю границу допустимого значения.
        /// </summary>
        private static int ReadInt(string s, int min = 0)
        {
            if (!int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v))
                throw new Exception($"Не целое число: \"{s}\".");
            if (v < min) throw new Exception($"Значение должно быть ≥ {min}.");
            return v;
        }

        /// <summary>
        /// Разбирает список целых чисел, разделенных запятыми, пробелами или точками с запятой.
        /// </summary>
        private static int[] ParseIntArray(string text)
        {
            return text.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries)
                       .Select(part => ReadInt(part.Trim()))
                       .ToArray();
        }

        /// <summary>
        /// Разбирает одномерный массив вещественных коэффициентов для целевой функции и вектора b.
        /// </summary>
        private static double[] ParseDoubleArraySpace(string text)
        {
            return text.Split(new[] { ' ', ';', '\t', ',', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                       .Select(ReadDouble)
                       .ToArray();
        }

        /// <summary>
        /// Преобразует многострочный текст в матрицу коэффициентов ограничений.
        /// </summary>
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

        /// <summary>
        /// Печатает набор числовых показателей в виде столбца `имя = значение`.
        /// </summary>
        private static void Print(StringBuilder sb, params (string name, double val)[] pairs)
        {
            foreach (var (name, val) in pairs)
                sb.AppendLine($"{name,-10} = {Math.Round(val, 3):0.###}");
        }
    }

    /// <summary>
    /// Хранит все данные шаблона, которые могут быть сохранены в Excel или восстановлены из него.
    /// </summary>
    internal sealed class AppTemplateData
    {
        public int TaskIndex { get; set; }
        public int SmoModeIndex { get; set; }
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

    /// <summary>
    /// Представляет одну строку матрицы стоимостей транспортной задачи.
    /// </summary>
    public sealed class CostRow
    {
        public int V1 { get; set; }
        public int V2 { get; set; }
        public int V3 { get; set; }
        public int V4 { get; set; }
        public int V5 { get; set; }

        /// <summary>
        /// Создает копию строки, чтобы UI и экспорт работали с независимыми объектами.
        /// </summary>
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

    /// <summary>
    /// Выполняет экспорт и импорт упрощенного Excel-шаблона без лишних служебных полей в видимой таблице.
    /// </summary>
    internal static class ExcelTemplateService
    {
        private const string FormatKey = "__format";
        private const string FormatValueV1 = "MathModelAllWork.Excel.v1";
        private const string FormatValueV2 = "MathModelAllWork.Excel.v2";
        private const string TransportMethodKey = "TransportMethod";
        private const string TransportMethodNorthWest = "NorthWest";
        private const string TransportMethodLeastCost = "LeastCost";
        private const string ResultLineKeyPrefix = "Result.";
        private static readonly XNamespace SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        /// <summary>
        /// Сохраняет данные шаблона в XLSX-файл OpenXML, не используя внешние Excel-библиотеки.
        /// </summary>
        public static void Export(string filePath, AppTemplateData data)
        {
            if (data == null)
                throw new ArgumentNullException(nameof(data));

            if (File.Exists(filePath))
                File.Delete(filePath);

            List<string[]> rows = BuildRows(data);

            using (var archive = ZipFile.Open(filePath, ZipArchiveMode.Create))
            {
                WriteEntry(archive, "[Content_Types].xml",
                    "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                    "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\" />" +
                    "<Default Extension=\"xml\" ContentType=\"application/xml\" />" +
                    "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\" />" +
                    "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\" />" +
                    "</Types>");
                WriteEntry(archive, "_rels/.rels",
                    "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\" />" +
                    "</Relationships>");
                WriteEntry(archive, "xl/workbook.xml",
                    "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                    "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                    "<sheets><sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\" /></sheets>" +
                    "</workbook>");
                WriteEntry(archive, "xl/_rels/workbook.xml.rels",
                    "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\" />" +
                    "</Relationships>");
                WriteEntry(archive, "xl/worksheets/sheet1.xml", BuildWorksheetXml(rows));
            }
        }

        /// <summary>
        /// Загружает Excel-файл, восстановливает данные формы и поддерживает старые версии формата.
        /// </summary>
        public static AppTemplateData Import(string filePath)
        {
            using (var archive = ZipFile.OpenRead(filePath))
            {
                Dictionary<string, string> rows = ReadRows(archive);
                string format = ReadString(rows, FormatKey);

                if (!string.IsNullOrWhiteSpace(format) &&
                    !string.Equals(format, FormatValueV1, StringComparison.Ordinal) &&
                    !string.Equals(format, FormatValueV2, StringComparison.Ordinal))
                {
                    throw new InvalidDataException("Поддерживается только Excel-файл, экспортированный этой программой в поддерживаемом формате.");
                }

                return new AppTemplateData
                {
                    TaskIndex = InferTaskIndex(rows),
                    SmoModeIndex = ReadInt(rows, nameof(AppTemplateData.SmoModeIndex)),
                    Lambda = ReadString(rows, nameof(AppTemplateData.Lambda)),
                    Mu = ReadString(rows, nameof(AppTemplateData.Mu)),
                    ServiceTime = ReadString(rows, nameof(AppTemplateData.ServiceTime)),
                    Channels = ReadString(rows, nameof(AppTemplateData.Channels)),
                    SystemCapacity = ReadString(rows, nameof(AppTemplateData.SystemCapacity)),
                    Supply = ReadString(rows, nameof(AppTemplateData.Supply)),
                    Demand = ReadString(rows, nameof(AppTemplateData.Demand)),
                    Objective = ReadString(rows, nameof(AppTemplateData.Objective)),
                    Matrix = ReadString(rows, nameof(AppTemplateData.Matrix)),
                    B = ReadString(rows, nameof(AppTemplateData.B)),
                    ResultText = ReadResultText(rows),
                    CostRows = ParseCostRows(ReadString(rows, nameof(AppTemplateData.CostRows)))
                };
            }
        }

        /// <summary>
        /// Формирует набор строк Excel только из нужных входных данных текущей задачи и строк результата.
        /// </summary>
        private static List<string[]> BuildRows(AppTemplateData data)
        {
            var rows = new List<string[]>();

            switch (data.TaskIndex)
            {
                case 0:
                    rows.Add(Row(nameof(AppTemplateData.SmoModeIndex), data.SmoModeIndex));
                    rows.Add(Row(nameof(AppTemplateData.Lambda), data.Lambda));
                    rows.Add(Row(nameof(AppTemplateData.Mu), data.Mu));
                    rows.Add(Row(nameof(AppTemplateData.ServiceTime), data.ServiceTime));
                    rows.Add(Row(nameof(AppTemplateData.Channels), data.Channels));
                    rows.Add(Row(nameof(AppTemplateData.SystemCapacity), data.SystemCapacity));
                    break;

                case 1:
                    rows.Add(Row(TransportMethodKey, TransportMethodNorthWest));
                    rows.Add(Row(nameof(AppTemplateData.Supply), data.Supply));
                    rows.Add(Row(nameof(AppTemplateData.Demand), data.Demand));
                    rows.Add(Row(nameof(AppTemplateData.CostRows), SerializeCostRows(data.CostRows)));
                    break;

                case 2:
                    rows.Add(Row(TransportMethodKey, TransportMethodLeastCost));
                    rows.Add(Row(nameof(AppTemplateData.Supply), data.Supply));
                    rows.Add(Row(nameof(AppTemplateData.Demand), data.Demand));
                    rows.Add(Row(nameof(AppTemplateData.CostRows), SerializeCostRows(data.CostRows)));
                    break;

                case 3:
                    rows.Add(Row(nameof(AppTemplateData.Objective), data.Objective));
                    rows.Add(Row(nameof(AppTemplateData.Matrix), data.Matrix));
                    rows.Add(Row(nameof(AppTemplateData.B), data.B));
                    break;
            }

            AppendResultRows(rows, data.ResultText);
            return rows;
        }

        /// <summary>
        /// Записывает содержимое в указанный файл внутри ZIP-архива XLSX.
        /// </summary>
        private static void WriteEntry(ZipArchive archive, string entryPath, string content)
        {
            var entry = archive.CreateEntry(entryPath, CompressionLevel.Optimal);
            using (var writer = new StreamWriter(entry.Open()))
                writer.Write(content);
        }

        /// <summary>
        /// Собирает XML-представление листа Excel из подготовленного списка строк.
        /// </summary>
        private static string BuildWorksheetXml(IEnumerable<string[]> rows)
        {
            var sheetData = new XElement(
                SpreadsheetNamespace + "sheetData",
                rows.Select(row =>
                    new XElement(
                        SpreadsheetNamespace + "row",
                        row.Select(value =>
                            new XElement(
                                SpreadsheetNamespace + "c",
                                new XAttribute("t", "inlineStr"),
                                new XElement(
                                    SpreadsheetNamespace + "is",
                                    new XElement(
                                        SpreadsheetNamespace + "t",
                                        new XAttribute(XNamespace.Xml + "space", "preserve"),
                                        value ?? string.Empty)))))));

            return "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                   new XDocument(new XElement(SpreadsheetNamespace + "worksheet", sheetData))
                       .ToString(SaveOptions.DisableFormatting);
        }

        /// <summary>
        /// Читает лист Excel в словарь `ключ -> значение`,
        /// где каждая строка листа соответствует одной логической записи.
        /// </summary>
        private static Dictionary<string, string> ReadRows(ZipArchive archive)
        {
            XDocument worksheet = LoadXml(archive, "xl/worksheets/sheet1.xml");
            IReadOnlyList<string> sharedStrings = ReadSharedStrings(archive);
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (XElement row in worksheet.Descendants(SpreadsheetNamespace + "row"))
            {
                List<string> cells = row.Elements(SpreadsheetNamespace + "c")
                    .Select(cell => ReadCellValue(cell, sharedStrings))
                    .ToList();

                if (cells.Count == 0 || string.IsNullOrWhiteSpace(cells[0]))
                    continue;

                result[cells[0]] = cells.Count > 1 ? cells[1] : string.Empty;
            }

            return result;
        }

        /// <summary>
        /// Загружает таблицу общих строк, если Excel решил хранить текст не прямо в ячейках.
        /// </summary>
        private static IReadOnlyList<string> ReadSharedStrings(ZipArchive archive)
        {
            var entry = archive.GetEntry("xl/sharedStrings.xml");
            if (entry == null)
                return Array.Empty<string>();

            using (var stream = entry.Open())
            {
                return XDocument.Load(stream)
                    .Descendants(SpreadsheetNamespace + "si")
                    .Select(item => string.Concat(item.Descendants(SpreadsheetNamespace + "t").Select(text => text.Value)))
                    .ToList();
            }
        }

        /// <summary>
        /// Загружает XML-документ из нужной части XLSX-архива.
        /// </summary>
        private static XDocument LoadXml(ZipArchive archive, string entryPath)
        {
            var entry = archive.GetEntry(entryPath);
            if (entry == null)
                throw new InvalidDataException($"В Excel-файле отсутствует часть {entryPath}.");

            using (var stream = entry.Open())
                return XDocument.Load(stream);
        }

        /// <summary>
        /// Читает текстовое содержимое отдельной ячейки Excel независимо от того,
        /// хранится ли оно прямо в ячейке или через таблицу общих строк.
        /// </summary>
        private static string ReadCellValue(XElement cell, IReadOnlyList<string> sharedStrings)
        {
            string type = (string)cell.Attribute("t");

            if (type == "inlineStr")
                return string.Concat(cell.Descendants(SpreadsheetNamespace + "t").Select(element => element.Value));

            string value = cell.Element(SpreadsheetNamespace + "v")?.Value ?? string.Empty;
            if (type == "s" && int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int sharedIndex)
                && sharedIndex >= 0 && sharedIndex < sharedStrings.Count)
            {
                return sharedStrings[sharedIndex];
            }

            return value;
        }

        /// <summary>
        /// Преобразует матрицу стоимостей в компактный текстовый формат для одной ячейки Excel.
        /// </summary>
        private static string SerializeCostRows(IEnumerable<CostRow> rows)
        {
            return string.Join(
                "\n",
                (rows ?? Enumerable.Empty<CostRow>()).Select(row =>
                    string.Join(";", new[]
                    {
                        row.V1.ToString(CultureInfo.InvariantCulture),
                        row.V2.ToString(CultureInfo.InvariantCulture),
                        row.V3.ToString(CultureInfo.InvariantCulture),
                        row.V4.ToString(CultureInfo.InvariantCulture),
                        row.V5.ToString(CultureInfo.InvariantCulture)
                    })));
        }

        /// <summary>
        /// Восстанавливает строки матрицы стоимостей из текстового представления,
        /// сохраненного в Excel-файле.
        /// </summary>
        private static List<CostRow> ParseCostRows(string rawValue)
        {
            return (rawValue ?? string.Empty)
                .Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.RemoveEmptyEntries)
                .Select(line => line.Split(';'))
                .Where(parts => parts.Length >= 5)
                .Select(parts => new CostRow
                {
                    V1 = ParseInt(parts[0]),
                    V2 = ParseInt(parts[1]),
                    V3 = ParseInt(parts[2]),
                    V4 = ParseInt(parts[3]),
                    V5 = ParseInt(parts[4])
                })
                .ToList();
        }

        /// <summary>
        /// Читает целое значение из словаря Excel-данных по указанному ключу.
        /// </summary>
        private static int ReadInt(IReadOnlyDictionary<string, string> values, string key)
        {
            return ParseInt(ReadString(values, key));
        }

        /// <summary>
        /// Читает строковое значение по ключу. Если ключ отсутствует, возвращает пустую строку.
        /// </summary>
        private static string ReadString(IReadOnlyDictionary<string, string> values, string key)
        {
            return values.TryGetValue(key, out string rawValue) ? rawValue ?? string.Empty : string.Empty;
        }

        /// <summary>
        /// Преобразует текст из Excel в целое число, допуская хранение как integer, так и floating-point.
        /// </summary>
        private static int ParseInt(string rawValue)
        {
            if (string.IsNullOrWhiteSpace(rawValue))
                return 0;

            if (int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int result))
                return result;

            if (double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double doubleValue))
                return Convert.ToInt32(Math.Round(doubleValue, MidpointRounding.AwayFromZero));

            throw new InvalidDataException($"Не удалось прочитать число из Excel: \"{rawValue}\".");
        }

        /// <summary>
        /// Создает строку Excel из пары `ключ - значение`.
        /// </summary>
        private static string[] Row(string key, object value)
        {
            return new[] { key, Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty };
        }

        /// <summary>
        /// Добавляет в экспорт результат расчета как набор отдельных строк,
        /// чтобы Excel отображал его раздельно, а не одним слитым блоком текста.
        /// </summary>
        private static void AppendResultRows(ICollection<string[]> rows, string resultText)
        {
            string[] resultLines = SplitResultLines(resultText).ToArray();
            if (resultLines.Length == 0)
                return;

            rows.Add(new[] { string.Empty, string.Empty });

            for (int index = 0; index < resultLines.Length; index++)
                rows.Add(Row(ResultLineKeyPrefix + (index + 1).ToString(CultureInfo.InvariantCulture), resultLines[index]));
        }

        /// <summary>
        /// Разбивает результат по строкам и отбрасывает пустые строки,
        /// чтобы экспорт выглядел компактно и читабельно.
        /// </summary>
        private static IEnumerable<string> SplitResultLines(string resultText)
        {
            return (resultText ?? string.Empty)
                .Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None)
                .Select(line => line.TrimEnd())
                .Where(line => !string.IsNullOrWhiteSpace(line));
        }

        /// <summary>
        /// Собирает текст результата обратно либо из старого поля ResultText,
        /// либо из нового набора Result.1, Result.2, ...
        /// </summary>
        private static string ReadResultText(IReadOnlyDictionary<string, string> values)
        {
            string singleResult = ReadString(values, nameof(AppTemplateData.ResultText));
            if (!string.IsNullOrWhiteSpace(singleResult))
                return singleResult;

            return string.Join(
                Environment.NewLine,
                values
                    .Where(pair => pair.Key.StartsWith(ResultLineKeyPrefix, StringComparison.OrdinalIgnoreCase))
                    .OrderBy(pair => ParseResultLineIndex(pair.Key))
                    .Select(pair => pair.Value));
        }

        /// <summary>
        /// Определяет индекс задачи по экспортированным полям, если явный TaskIndex отсутствует.
        /// </summary>
        private static int InferTaskIndex(IReadOnlyDictionary<string, string> values)
        {
            if (values.ContainsKey(nameof(AppTemplateData.TaskIndex)))
                return ReadInt(values, nameof(AppTemplateData.TaskIndex));

            if (values.ContainsKey(TransportMethodKey))
                return ParseTransportTaskIndex(ReadString(values, TransportMethodKey));

            if (HasAnyKey(values, nameof(AppTemplateData.Supply), nameof(AppTemplateData.Demand), nameof(AppTemplateData.CostRows)))
                return 1;

            if (HasAnyKey(values, nameof(AppTemplateData.Objective), nameof(AppTemplateData.Matrix), nameof(AppTemplateData.B)))
                return 3;

            return 0;
        }

        /// <summary>
        /// Определяет, присутствует ли в наборе данных хотя бы один из указанных ключей.
        /// </summary>
        private static bool HasAnyKey(IReadOnlyDictionary<string, string> values, params string[] keys)
        {
            return keys.Any(values.ContainsKey);
        }

        /// <summary>
        /// Восстанавливает выбранный метод транспортной задачи из строки Excel.
        /// </summary>
        private static int ParseTransportTaskIndex(string rawValue)
        {
            if (string.Equals(rawValue, TransportMethodLeastCost, StringComparison.OrdinalIgnoreCase))
                return 2;

            return 1;
        }

        /// <summary>
        /// Извлекает номер строки результата из ключа вида `Result.N`.
        /// </summary>
        private static int ParseResultLineIndex(string key)
        {
            string rawIndex = key.Substring(ResultLineKeyPrefix.Length);
            return int.TryParse(rawIndex, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index)
                ? index
                : int.MaxValue;
        }
    }
}
