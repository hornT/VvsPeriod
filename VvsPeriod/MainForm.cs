using System;
using System.Windows.Forms;
using JetBrains.Annotations;
using NLog;
using PeriodProcessor;

namespace VvsPeriod
{
    public partial class MainForm : Form
    {
        [NotNull]
        private readonly Logger _logger;

        [NotNull]
        private readonly Processor _processor;

        public MainForm()
        {
            InitializeComponent();

            _logger = LogManager.GetCurrentClassLogger() ?? throw new Exception("logger is null");

            _processor = new Processor();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                Filter = "xlsx files (*.xlsx)|*.xlsx",
                RestoreDirectory = true
            };

            if (ofd.ShowDialog() != DialogResult.OK)
                return;

            if (string.IsNullOrEmpty(ofd.FileName) == true)
            {
                _logger.Warn("saveFileDialog.FileName empty");
                return;
            }

            try
            {
                Cursor = Cursors.WaitCursor;
                dataGridView.DataSource = _processor.ProcessFile(ofd.FileName);
            }
            catch (Exception exception)
            {
                _logger.Error(exception);
                MessageBox.Show("Возникла ошибка. Подробности в логе", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Arrow;
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "xlsx files (*.xlsx)|*.xlsx",
                RestoreDirectory = true
            };

            if (saveFileDialog.ShowDialog() != DialogResult.OK)
                return;

            if (string.IsNullOrEmpty(saveFileDialog.FileName) == true)
            {
                _logger.Warn("saveFileDialog.FileName empty");
                return;
            }

            try
            {
                _processor.ExportToExcel(saveFileDialog.FileName);
            }
            catch (Exception exception)
            {
                _logger.Error(exception);
                MessageBox.Show("Возникла ошибка. Подробности в логе", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            MessageBox.Show("Выполнено!", "Выгрузка", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
