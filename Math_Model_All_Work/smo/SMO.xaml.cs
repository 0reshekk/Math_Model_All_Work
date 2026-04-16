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

namespace Math_Model_All_Work.smo
{
    /// <summary>
    /// Логика взаимодействия для SMO.xaml
    /// </summary>
    public partial class SMO : Window
    {
        private DataTable costTable;
        public SMO()
        {
            InitializeComponent();
            Loaded += (s, e) => cboSMOType.SelectedIndex = 0;
        }

        private void CboSMOType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (QueuePanel == null || cboSMOType.SelectedItem == null)
                return;

            string selected = ((ComboBoxItem)cboSMOType.SelectedItem).Content.ToString();

            QueuePanel.Visibility =
                selected == "СМО с ограниченной очередью"
                ? Visibility.Visible
                : Visibility.Collapsed;
        }

        private void BtnCalculate_Click(object sender, RoutedEventArgs e)
        {
            int n, m;
            double lambda, mu;

            if (!ValidateInput(out n, out lambda, out mu, out m))
                return;

            if (mu == 0)
            {
                double tob;
                if (!double.TryParse(txtTob.Text, out tob) || tob <= 0)
                {
                    MessageBox.Show("Введите корректное tобс > 0");
                    return;
                }
                mu = 1.0 / tob;
            }

            string type = ((ComboBoxItem)cboSMOType.SelectedItem).Content.ToString();

            string result = "";

            if (type == "СМО с отказами")
                result = CalculateLossSystem(n, lambda, mu);
            else if (type == "СМО с ограниченной очередью")
                result = CalculateLimitedQueueSystem(n, m, lambda, mu);
            else if (type == "СМО с неограниченной очередью")
                result = CalculateUnlimitedQueueSystem(n, lambda, mu);

            txtResult.Text = result;
        }

        // =========================
        // СМО С ОТКАЗАМИ
        // =========================
        private string CalculateLossSystem(int n, double lambda, double mu)
        {
            StringBuilder sb = new StringBuilder();

            double rho = lambda / mu;
            double sum = 0;

            sb.AppendLine("СИСТЕМА С ОТКАЗАМИ");
            sb.AppendLine("----------------------------------------");

            for (int k = 0; k <= n; k++)
            {
                sum += Math.Pow(rho, k) / Factorial(k);
            }

            double P0 = 1 / sum;

            double[] P = new double[n + 1];

            for (int k = 0; k <= n; k++)
                P[k] = Math.Pow(rho, k) / Factorial(k) * P0;

            double Pblock = P[n];

            double L = 0;
            for (int k = 1; k <= n; k++)
                L += k * P[k];

            double U = L / n;

            sb.AppendLine($"λ = {lambda:F4}");
            sb.AppendLine($"μ = {mu:F4}");
            sb.AppendLine($"ρ = {rho:F4}");
            sb.AppendLine($"P отказа = {Pblock:F4}");
            sb.AppendLine($"Загрузка = {U:P2}");

            return sb.ToString();
        }

        // =========================
        // СМО С ОГРАНИЧЕННОЙ ОЧЕРЕДЬЮ
        // =========================
        private string CalculateLimitedQueueSystem(int n, int m, double lambda, double mu)
        {
            StringBuilder sb = new StringBuilder();

            double rho = lambda / mu;
            double sum = 0;

            int max = n + m;

            for (int k = 0; k <= max; k++)
                sum += Math.Pow(rho, k) / Factorial(Math.Min(k, n));

            double P0 = 1 / sum;

            double[] P = new double[max + 1];

            for (int k = 0; k <= max; k++)
                P[k] = Math.Pow(rho, k) / Factorial(Math.Min(k, n)) * P0;

            double Pblock = P[max];

            double Lq = 0;

            for (int k = n + 1; k <= max; k++)
                Lq += (k - n) * P[k];

            double Wq = (lambda * (1 - Pblock) > 0)
                ? Lq / (lambda * (1 - Pblock))
                : 0;

            sb.AppendLine("СМО С ОГРАНИЧЕННОЙ ОЧЕРЕДЬЮ");
            sb.AppendLine("----------------------------------------");
            sb.AppendLine("P отказа: " + Pblock.ToString("F4"));
            sb.AppendLine("L очереди: " + Lq.ToString("F4"));
            sb.AppendLine("W ожидания: " + Wq.ToString("F4"));

            return sb.ToString();
        }

        // =========================
        // СМО С НЕОГРАНИЧЕННОЙ ОЧЕРЕДЬЮ
        // =========================
        private string CalculateUnlimitedQueueSystem(int n, double lambda, double mu)
        {
            StringBuilder sb = new StringBuilder();

            double rho = lambda / (n * mu);

            sb.AppendLine("СМО С НЕОГРАНИЧЕННОЙ ОЧЕРЕДЬЮ");
            sb.AppendLine("----------------------------------------");

            if (rho >= 1)
            {
                sb.AppendLine("Система неустойчива (ρ ≥ 1)");
                return sb.ToString();
            }

            double L = rho / (1 - rho);
            double W = 1 / (n * mu - lambda);
            double Wq = W - 1 / mu;

            sb.AppendLine("L: " + L.ToString("F4"));
            sb.AppendLine("W: " + W.ToString("F4"));
            sb.AppendLine("Wq: " + Wq.ToString("F4"));

            return sb.ToString();
        }

        // =========================
        // ВАЛИДАЦИЯ
        // =========================
        private bool ValidateInput(out int n, out double lambda, out double mu, out int m)
        {
            n = 0; lambda = 0; mu = 0; m = 0;

            if (!int.TryParse(txtChannels.Text, out n) || n <= 0)
            {
                MessageBox.Show("n должно быть > 0");
                return false;
            }

            if (!double.TryParse(txtLambda.Text, out lambda) || lambda <= 0)
            {
                MessageBox.Show("λ должно быть > 0");
                return false;
            }

            if (!double.TryParse(txtMu.Text, out mu) || mu < 0)
            {
                MessageBox.Show("μ должно быть ≥ 0");
                return false;
            }

            string type = ((ComboBoxItem)cboSMOType.SelectedItem).Content.ToString();

            if (type == "СМО с ограниченной очередью")
            {
                if (!int.TryParse(txtQueue.Text, out m) || m < 0)
                {
                    MessageBox.Show("m должно быть ≥ 0");
                    return false;
                }
            }

            return true;
        }

        // =========================
        // ОЧИСТКА
        // =========================
        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            txtChannels.Text = "";
            txtLambda.Text = "";
            txtMu.Text = "";
            txtTob.Text = "";
            txtQueue.Text = "";
            txtResult.Text = "";

            cboSMOType.SelectedIndex = 0;
        }

        // =========================
        // ФАКТОРИАЛ
        // =========================
        private double Factorial(int k)
        {
            double r = 1;
            for (int i = 2; i <= k; i++)
                r *= i;
            return r;
        }
    }
}
