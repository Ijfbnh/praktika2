using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace TransportApp
{
    public class MainForm : Form
    {
        // Компоненты
        private DataGridView dgData;
        private TextBox txtSearch;
        private Button btnSearch;
        private Button btnClear;
        private Button btnTask1;
        private Button btnTask2;
        private Button btnTask3;
        private Button btnAdd;
        private Button btnSave;
        private Button btnDelete;
        private Button btnSort;
        private Button btnExit;
        private Button btnRefresh;
        private Button btnSelectDB;
        private Label lblStatus;
        private Label lblCount;
        private Label lblConnection;

        private OleDbConnection connection;
        private DataTable dataTable;
        private string dbPath = "";

        public MainForm()
        {
            InitializeComponent();

            // Проверяем наличие файла по умолчанию
            string defaultPath = Application.StartupPath + @"\Transport.accdb";
            if (File.Exists(defaultPath))
            {
                dbPath = defaultPath;
                ConnectToDatabase();
            }
        }

        private void InitializeComponent()
        {
            this.Text = "Грузооборот по видам транспорта (Access)";
            this.Size = new Size(1050, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(240, 240, 240);
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimumSize = new Size(900, 550);

            TableLayoutPanel mainTableLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 7,
                Padding = new Padding(10),
                BackColor = Color.FromArgb(240, 240, 240)
            };

            mainTableLayout.RowStyles.Clear();
            mainTableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));
            mainTableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            mainTableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 60));
            mainTableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));
            mainTableLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            mainTableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 60));
            mainTableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));

            // Заголовок
            Panel headerPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(44, 62, 80),
                Padding = new Padding(15, 0, 15, 0)
            };

            Label lblTitle = new Label
            {
                Text = "🚆 Изменение структуры грузооборота по видам транспорта (в %)",
                Font = new Font("Segoe UI", 14, FontStyle.Bold),
                ForeColor = Color.White,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
            headerPanel.Controls.Add(lblTitle);
            mainTableLayout.Controls.Add(headerPanel, 0, 0);

            // Статус подключения
            Panel connPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(236, 240, 241),
                Padding = new Padding(10, 2, 10, 2)
            };

            FlowLayoutPanel connFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                BackColor = Color.FromArgb(236, 240, 241)
            };

            lblConnection = new Label
            {
                Text = "База данных не выбрана",
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.FromArgb(52, 73, 94),
                Size = new Size(400, 25),
                TextAlign = ContentAlignment.MiddleLeft
            };
            connFlow.Controls.Add(lblConnection);

            btnSelectDB = new Button
            {
                Text = "Выбрать БД",
                Width = 120,
                Height = 25,
                BackColor = Color.FromArgb(52, 152, 219),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            btnSelectDB.Click += BtnSelectDB_Click;
            connFlow.Controls.Add(btnSelectDB);

            connPanel.Controls.Add(connFlow);
            mainTableLayout.Controls.Add(connPanel, 0, 1);

            // Поиск
            Panel searchPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(10, 8, 10, 8)
            };

            FlowLayoutPanel searchFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                BackColor = Color.White
            };

            Label lblSearch = new Label
            {
                Text = "Поиск по виду транспорта:",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Size = new Size(180, 30),
                TextAlign = ContentAlignment.MiddleLeft
            };
            searchFlow.Controls.Add(lblSearch);

            txtSearch = new TextBox
            {
                Width = 200,
                Height = 30,
                Font = new Font("Segoe UI", 10),
                Margin = new Padding(5, 0, 5, 0)
            };
            searchFlow.Controls.Add(txtSearch);

            btnSearch = new Button
            {
                Text = "Поиск",
                Width = 80,
                Height = 32,
                BackColor = Color.FromArgb(52, 152, 219),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Margin = new Padding(0, 0, 5, 0)
            };
            btnSearch.Click += BtnSearch_Click;
            searchFlow.Controls.Add(btnSearch);

            btnClear = new Button
            {
                Text = "Сбросить выборку",
                Width = 130,
                Height = 32,
                BackColor = Color.FromArgb(149, 165, 166),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            btnClear.Click += BtnClear_Click;
            searchFlow.Controls.Add(btnClear);

            searchPanel.Controls.Add(searchFlow);
            mainTableLayout.Controls.Add(searchPanel, 0, 2);

            // Задания
            Panel taskPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(248, 249, 250),
                Padding = new Padding(10, 5, 10, 5)
            };

            FlowLayoutPanel taskFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                BackColor = Color.FromArgb(248, 249, 250)
            };

            Label lblTasks = new Label
            {
                Text = "Задания:",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Size = new Size(80, 30),
                TextAlign = ContentAlignment.MiddleLeft
            };
            taskFlow.Controls.Add(lblTasks);

            btnTask1 = new Button
            {
                Text = "Задание 1 (Макс. груз 2011)",
                Width = 200,
                Height = 32,
                BackColor = Color.FromArgb(155, 89, 182),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Margin = new Padding(0, 0, 5, 0)
            };
            btnTask1.Click += BtnTask1_Click;
            taskFlow.Controls.Add(btnTask1);

            btnTask2 = new Button
            {
                Text = "Задание 2 (Пасс. > 15%)",
                Width = 200,
                Height = 32,
                BackColor = Color.FromArgb(52, 152, 219),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Margin = new Padding(0, 0, 5, 0)
            };
            btnTask2.Click += BtnTask2_Click;
            taskFlow.Controls.Add(btnTask2);

            btnTask3 = new Button
            {
                Text = "Задание 3 (Пасс. < 40%)",
                Width = 200,
                Height = 32,
                BackColor = Color.FromArgb(46, 204, 113),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Margin = new Padding(0, 0, 5, 0)
            };
            btnTask3.Click += BtnTask3_Click;
            taskFlow.Controls.Add(btnTask3);

            btnSort = new Button
            {
                Text = "Сортировка (А-Я)",
                Width = 150,
                Height = 32,
                BackColor = Color.FromArgb(241, 196, 15),
                ForeColor = Color.Black,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Margin = new Padding(0, 0, 5, 0)
            };
            btnSort.Click += BtnSort_Click;
            taskFlow.Controls.Add(btnSort);

            btnRefresh = new Button
            {
                Text = "Обновить",
                Width = 100,
                Height = 32,
                BackColor = Color.FromArgb(26, 188, 156),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            btnRefresh.Click += BtnRefresh_Click;
            taskFlow.Controls.Add(btnRefresh);

            taskPanel.Controls.Add(taskFlow);
            mainTableLayout.Controls.Add(taskPanel, 0, 3);

            // DataGridView
            Panel gridPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(5)
            };

            dgData = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersVisible = false,
                GridColor = Color.FromArgb(200, 200, 200),
                Font = new Font("Segoe UI", 10),
                RowTemplate = { Height = 32 }
            };

            gridPanel.Controls.Add(dgData);
            mainTableLayout.Controls.Add(gridPanel, 0, 4);

            // Действия
            Panel actionPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(10, 10, 10, 10)
            };

            FlowLayoutPanel actionFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                BackColor = Color.White
            };

            Label lblActions = new Label
            {
                Text = "Действия с таблицами:",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Size = new Size(170, 35),
                TextAlign = ContentAlignment.MiddleLeft
            };
            actionFlow.Controls.Add(lblActions);

            btnAdd = new Button
            {
                Text = "Добавить",
                Width = 100,
                Height = 35,
                BackColor = Color.FromArgb(46, 204, 113),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Margin = new Padding(0, 0, 5, 0)
            };
            btnAdd.Click += BtnAdd_Click;
            actionFlow.Controls.Add(btnAdd);

            btnSave = new Button
            {
                Text = "Сохранить",
                Width = 100,
                Height = 35,
                BackColor = Color.FromArgb(52, 152, 219),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Margin = new Padding(0, 0, 5, 0)
            };
            btnSave.Click += BtnSave_Click;
            actionFlow.Controls.Add(btnSave);

            btnDelete = new Button
            {
                Text = "Удалить",
                Width = 100,
                Height = 35,
                BackColor = Color.FromArgb(231, 76, 60),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Margin = new Padding(0, 0, 5, 0)
            };
            btnDelete.Click += BtnDelete_Click;
            actionFlow.Controls.Add(btnDelete);

            btnExit = new Button
            {
                Text = "Выход",
                Width = 100,
                Height = 35,
                BackColor = Color.FromArgb(127, 140, 141),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            btnExit.Click += BtnExit_Click;
            actionFlow.Controls.Add(btnExit);

            actionPanel.Controls.Add(actionFlow);
            mainTableLayout.Controls.Add(actionPanel, 0, 5);

            // Статус
            Panel statusPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(44, 62, 80),
                Padding = new Padding(15, 0, 15, 0)
            };

            TableLayoutPanel statusTable = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                BackColor = Color.FromArgb(44, 62, 80)
            };

            statusTable.ColumnStyles.Clear();
            statusTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            statusTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));

            lblStatus = new Label
            {
                Text = "Выберите базу данных",
                ForeColor = Color.Yellow,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
            statusTable.Controls.Add(lblStatus, 0, 0);

            lblCount = new Label
            {
                Text = "Записей: 0",
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10),
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleRight
            };
            statusTable.Controls.Add(lblCount, 1, 0);

            statusPanel.Controls.Add(statusTable);
            mainTableLayout.Controls.Add(statusPanel, 0, 6);

            this.Controls.Add(mainTableLayout);
        }

        // ============================================================
        // ПОДКЛЮЧЕНИЕ К БАЗЕ ДАННЫХ
        // ============================================================

        private void BtnSelectDB_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Access Database (*.accdb;*.mdb)|*.accdb;*.mdb|All files (*.*)|*.*";
                openFileDialog.Title = "Выберите файл базы данных Access";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    dbPath = openFileDialog.FileName;
                    ConnectToDatabase();
                }
            }
        }

        private void ConnectToDatabase()
        {
            try
            {
                if (connection != null && connection.State == ConnectionState.Open)
                    connection.Close();

                if (!File.Exists(dbPath))
                {
                    lblConnection.Text = "Файл не найден!";
                    lblStatus.Text = "Ошибка: файл не найден";
                    lblStatus.ForeColor = Color.Red;
                    return;
                }

                string[] providers = {
                    $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};Persist Security Info=False;",
                    $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={dbPath};",
                    $"Provider=Microsoft.ACE.OLEDB.15.0;Data Source={dbPath};",
                    $"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={dbPath};"
                };

                bool connected = false;
                string lastError = "";

                foreach (string connStr in providers)
                {
                    try
                    {
                        connection = new OleDbConnection(connStr);
                        connection.Open();

                        // Проверяем наличие таблицы
                        DataTable schemaTable = connection.GetSchema("Tables");
                        bool tableExists = false;
                        foreach (DataRow row in schemaTable.Rows)
                        {
                            if (row["TABLE_NAME"].ToString() == "Транспорт")
                            {
                                tableExists = true;
                                break;
                            }
                        }

                        if (!tableExists)
                        {
                            connection.Close();
                            throw new Exception("Таблица 'Транспорт' не найдена в базе данных");
                        }

                        connected = true;
                        lblConnection.Text = $"Подключено: {Path.GetFileName(dbPath)}";
                        lblStatus.Text = "Подключение к базе данных установлено";
                        lblStatus.ForeColor = Color.FromArgb(46, 204, 113);
                        LoadData();
                        break;
                    }
                    catch (Exception ex)
                    {
                        lastError = ex.Message;
                        if (connection != null)
                        {
                            connection.Close();
                            connection = null;
                        }
                    }
                }

                if (!connected)
                {
                    MessageBox.Show($"Не удалось подключиться к базе данных.\nОшибка: {lastError}",
                        "Ошибка подключения", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    lblConnection.Text = "Ошибка подключения";
                    lblStatus.Text = "Ошибка подключения к БД";
                    lblStatus.ForeColor = Color.Red;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка подключения к базе данных:\n{ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "Ошибка подключения";
                lblStatus.ForeColor = Color.Red;
            }
        }

        // ============================================================
        // ЗАГРУЗКА ДАННЫХ - ИСПРАВЛЕНО ДЛЯ ВАШИХ ИМЕН ПОЛЕЙ
        // ============================================================

        private void LoadData(string filter = "")
        {
            try
            {
                if (connection == null || connection.State != ConnectionState.Open)
                {
                    if (!string.IsNullOrEmpty(dbPath) && File.Exists(dbPath))
                    {
                        ConnectToDatabase();
                    }
                    else
                    {
                        return;
                    }
                }

                // ВАЖНО: имена полей БЕЗ подчеркиваний, как в вашей таблице
                string query = @"
                    SELECT 
                        [Видтранспорта],
                        [Грузооборот2011],
                        [Грузооборот2013],
                        [Грузооборот2015],
                        [Пассажирооборот2013],
                        [Пассажирооборот2017],
                        [Пассажирооборот2018]
                    FROM [Транспорт]";

                if (!string.IsNullOrEmpty(filter))
                {
                    query += $" WHERE [Видтранспорта] LIKE '%{filter.Replace("'", "''")}%'";
                }

                query += " ORDER BY [Видтранспорта]";

                using (OleDbCommand cmd = new OleDbCommand(query, connection))
                {
                    dataTable = new DataTable();
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(cmd))
                    {
                        adapter.Fill(dataTable);
                    }
                }

                dgData.DataSource = dataTable;

                // Настройка колонок
                if (dgData.Columns.Count > 0)
                {
                    // Переименовываем колонки для отображения
                    dgData.Columns["Видтранспорта"].HeaderText = "Вид транспорта";
                    dgData.Columns["Видтранспорта"].Width = 150;

                    dgData.Columns["Грузооборот2011"].HeaderText = "Грузооборот 2011";
                    dgData.Columns["Грузооборот2011"].DefaultCellStyle.Format = "F1";
                    dgData.Columns["Грузооборот2011"].Width = 120;

                    dgData.Columns["Грузооборот2013"].HeaderText = "Грузооборот 2013";
                    dgData.Columns["Грузооборот2013"].DefaultCellStyle.Format = "F1";
                    dgData.Columns["Грузооборот2013"].Width = 120;

                    dgData.Columns["Грузооборот2015"].HeaderText = "Грузооборот 2015";
                    dgData.Columns["Грузооборот2015"].DefaultCellStyle.Format = "F1";
                    dgData.Columns["Грузооборот2015"].Width = 120;

                    dgData.Columns["Пассажирооборот2013"].HeaderText = "Пассажирооборот 2013";
                    dgData.Columns["Пассажирооборот2013"].DefaultCellStyle.Format = "F1";
                    dgData.Columns["Пассажирооборот2013"].Width = 140;

                    dgData.Columns["Пассажирооборот2017"].HeaderText = "Пассажирооборот 2017";
                    dgData.Columns["Пассажирооборот2017"].DefaultCellStyle.Format = "F1";
                    dgData.Columns["Пассажирооборот2017"].Width = 140;

                    dgData.Columns["Пассажирооборот2018"].HeaderText = "Пассажирооборот 2018";
                    dgData.Columns["Пассажирооборот2018"].DefaultCellStyle.Format = "F1";
                    dgData.Columns["Пассажирооборот2018"].Width = 140;

                    // Выравнивание числовых колонок
                    for (int i = 1; i < dgData.Columns.Count; i++)
                    {
                        dgData.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    }
                }

                lblCount.Text = $"Записей: {dataTable.Rows.Count}";
                lblStatus.Text = $"Загружено {dataTable.Rows.Count} записей";
            }
            catch (Exception ex)
            {
                string errorMsg = $"Ошибка загрузки данных:\n{ex.Message}\n\n";
                errorMsg += "Проверьте структуру таблицы в Access.\n";
                errorMsg += "Имена полей должны точно соответствовать:\n";
                errorMsg += "- Видтранспорта\n";
                errorMsg += "- Грузооборот2011\n";
                errorMsg += "- Грузооборот2013\n";
                errorMsg += "- Грузооборот2015\n";
                errorMsg += "- Пассажирооборот2013\n";
                errorMsg += "- Пассажирооборот2017\n";
                errorMsg += "- Пассажирооборот2018";

                MessageBox.Show(errorMsg, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "Ошибка загрузки данных";
                lblStatus.ForeColor = Color.Red;
            }
        }

        // ============================================================
        // ЗАДАНИЕ 1 - ИСПРАВЛЕНО
        // ============================================================

        private void BtnTask1_Click(object sender, EventArgs e)
        {
            try
            {
                if (connection == null || connection.State != ConnectionState.Open)
                {
                    MessageBox.Show("Нет подключения к базе данных", "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Используем правильные имена полей
                string query = @"
                    SELECT TOP 1 
                        [Видтранспорта], 
                        [Грузооборот2011]
                    FROM [Транспорт]
                    ORDER BY [Грузооборот2011] DESC";

                DataTable result = new DataTable();
                using (OleDbCommand cmd = new OleDbCommand(query, connection))
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(cmd))
                {
                    adapter.Fill(result);
                }

                if (result.Rows.Count > 0)
                {
                    string name = result.Rows[0]["Видтранспорта"].ToString();
                    double value = Convert.ToDouble(result.Rows[0]["Грузооборот2011"]);

                    dgData.DataSource = result;
                    lblStatus.Text = $"Максимальный грузооборот в 2011: {name} ({value:F1}%)";

                    MessageBox.Show($"Задание 1:\n\nВ 2011 году максимальный грузооборот был у:\n" +
                                   $"Вид транспорта: {name}\n" +
                                   $"Грузооборот: {value:F1}%",
                                   "Результат", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка выполнения задания 1:\n{ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ============================================================
        // ЗАДАНИЕ 2 - ИСПРАВЛЕНО
        // ============================================================

        private void BtnTask2_Click(object sender, EventArgs e)
        {
            try
            {
                if (connection == null || connection.State != ConnectionState.Open)
                {
                    MessageBox.Show("Нет подключения к базе данных", "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                string query = @"
                    SELECT 
                        [Видтранспорта],
                        [Пассажирооборот2018],
                        ([Пассажирооборот2018] / (SELECT SUM([Пассажирооборот2018]) FROM [Транспорт]) * 100) AS Процент
                    FROM [Транспорт]
                    WHERE ([Пассажирооборот2018] / (SELECT SUM([Пассажирооборот2018]) FROM [Транспорт]) * 100) > 15
                    ORDER BY [Видтранспорта]";

                DataTable result = new DataTable();
                using (OleDbCommand cmd = new OleDbCommand(query, connection))
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(cmd))
                {
                    adapter.Fill(result);
                }

                if (result.Rows.Count > 0)
                {
                    dgData.DataSource = result;
                    lblStatus.Text = $"Найдено {result.Rows.Count} видов транспорта с пассажирооборотом > 15%";

                    string message = "Задание 2:\n\nВиды транспорта с пассажирооборотом в 2018 > 15%:\n\n";
                    foreach (DataRow row in result.Rows)
                    {
                        message += $"{row["Видтранспорта"]}: {row["Пассажирооборот2018"]}% ({row["Процент"]:F1}% от общего)\n";
                    }

                    MessageBox.Show(message, "Результат", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Нет видов транспорта с пассажирооборотом > 15%",
                        "Результат", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка выполнения задания 2:\n{ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ============================================================
        // ЗАДАНИЕ 3 - ИСПРАВЛЕНО
        // ============================================================

        private void BtnTask3_Click(object sender, EventArgs e)
        {
            try
            {
                if (connection == null || connection.State != ConnectionState.Open)
                {
                    MessageBox.Show("Нет подключения к базе данных", "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                string query = @"
                    SELECT 
                        [Видтранспорта],
                        [Пассажирооборот2013],
                        [Пассажирооборот2017],
                        [Пассажирооборот2018],
                        ([Пассажирооборот2013] + [Пассажирооборот2017] + [Пассажирооборот2018]) / 3 AS Средний_пассажир,
                        (([Пассажирооборот2013] + [Пассажирооборот2017] + [Пассажирооборот2018]) / 3 / 
                            (SELECT AVG(Средний)
                             FROM (
                                 SELECT ([Пассажирооборот2013] + [Пассажирооборот2017] + [Пассажирооборот2018]) / 3 AS Средний
                                 FROM [Транспорт]
                             ) AS Подзапрос) * 100) AS Процент_от_среднего
                    FROM [Транспорт]
                    WHERE (([Пассажирооборот2013] + [Пассажирооборот2017] + [Пассажирооборот2018]) / 3 / 
                            (SELECT AVG(Средний)
                             FROM (
                                 SELECT ([Пассажирооборот2013] + [Пассажирооборот2017] + [Пассажирооборот2018]) / 3 AS Средний
                                 FROM [Транспорт]
                             ) AS Подзапрос) * 100) < 40
                    ORDER BY [Видтранспорта]";

                DataTable result = new DataTable();
                using (OleDbCommand cmd = new OleDbCommand(query, connection))
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(cmd))
                {
                    adapter.Fill(result);
                }

                if (result.Rows.Count > 0)
                {
                    dgData.DataSource = result;
                    lblStatus.Text = $"Найдено {result.Rows.Count} видов транспорта с пассажирооборотом < 40%";

                    string message = "Задание 3:\n\nТранспорт со средним пассажирооборотом менее 40%:\n\n";
                    foreach (DataRow row in result.Rows)
                    {
                        message += $"{row["Видтранспорта"]}: средний {row["Средний_пассажир"]:F1}% ({row["Процент_от_среднего"]:F1}% от общего)\n";
                    }

                    MessageBox.Show(message, "Результат", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Нет видов транспорта со средним пассажирооборотом < 40%",
                        "Результат", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка выполнения задания 3:\n{ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ============================================================
        // ОСТАЛЬНЫЕ МЕТОДЫ
        // ============================================================

        private void BtnSearch_Click(object sender, EventArgs e)
        {
            string searchText = txtSearch.Text.Trim();
            if (string.IsNullOrEmpty(searchText))
            {
                MessageBox.Show("Введите текст для поиска", "Предупреждение",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            LoadData(searchText);
        }

        private void BtnClear_Click(object sender, EventArgs e)
        {
            txtSearch.Clear();
            LoadData();
        }

        private void BtnSort_Click(object sender, EventArgs e)
        {
            LoadData();
            lblStatus.Text = "Записи отсортированы по алфавиту";
        }

        private void BtnRefresh_Click(object sender, EventArgs e)
        {
            LoadData();
            lblStatus.Text = "Данные обновлены";
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (connection == null || connection.State != ConnectionState.Open)
            {
                MessageBox.Show("Нет подключения к базе данных", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var dialog = new AddTransportDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    string query = $@"
                        INSERT INTO [Транспорт] 
                        ([Видтранспорта], [Грузооборот2011], [Грузооборот2013], [Грузооборот2015],
                         [Пассажирооборот2013], [Пассажирооборот2017], [Пассажирооборот2018])
                        VALUES (
                            '{dialog.NewName.Replace("'", "''")}',
                            {dialog.NewCargo2011.ToString(System.Globalization.CultureInfo.InvariantCulture)},
                            {dialog.NewCargo2013.ToString(System.Globalization.CultureInfo.InvariantCulture)},
                            {dialog.NewCargo2015.ToString(System.Globalization.CultureInfo.InvariantCulture)},
                            {dialog.NewPassenger2013.ToString(System.Globalization.CultureInfo.InvariantCulture)},
                            {dialog.NewPassenger2017.ToString(System.Globalization.CultureInfo.InvariantCulture)},
                            {dialog.NewPassenger2018.ToString(System.Globalization.CultureInfo.InvariantCulture)}
                        )";

                    using (OleDbCommand cmd = new OleDbCommand(query, connection))
                    {
                        cmd.ExecuteNonQuery();
                    }

                    LoadData();
                    lblStatus.Text = $"Добавлен вид транспорта: {dialog.NewName}";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка добавления записи:\n{ex.Message}", "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            LoadData();
            lblStatus.Text = "Данные сохранены";
            MessageBox.Show("Данные сохранены", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (dgData.SelectedRows.Count == 0)
            {
                MessageBox.Show("Выберите запись для удаления", "Предупреждение",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string name = dgData.SelectedRows[0].Cells["Видтранспорта"].Value.ToString();

            var result = MessageBox.Show($"Удалить запись о виде транспорта '{name}'?",
                "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                try
                {
                    if (connection == null || connection.State != ConnectionState.Open)
                    {
                        MessageBox.Show("Нет подключения к базе данных", "Ошибка",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    string query = $"DELETE FROM [Транспорт] WHERE [Видтранспорта] = '{name.Replace("'", "''")}'";

                    using (OleDbCommand cmd = new OleDbCommand(query, connection))
                    {
                        cmd.ExecuteNonQuery();
                    }

                    LoadData();
                    lblStatus.Text = $"Удален вид транспорта: {name}";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка удаления записи:\n{ex.Message}", "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnExit_Click(object sender, EventArgs e)
        {
            if (connection != null && connection.State == ConnectionState.Open)
                connection.Close();

            Application.Exit();
        }
    }

    // ============================================================
    // ДИАЛОГ ДОБАВЛЕНИЯ
    // ============================================================
    public class AddTransportDialog : Form
    {
        private TableLayoutPanel mainLayout;
        private TextBox txtName;
        private TextBox txtCargo2011;
        private TextBox txtCargo2013;
        private TextBox txtCargo2015;
        private TextBox txtPassenger2013;
        private TextBox txtPassenger2017;
        private TextBox txtPassenger2018;
        private Button btnOk;
        private Button btnCancel;

        public string NewName { get; private set; }
        public double NewCargo2011 { get; private set; }
        public double NewCargo2013 { get; private set; }
        public double NewCargo2015 { get; private set; }
        public double NewPassenger2013 { get; private set; }
        public double NewPassenger2017 { get; private set; }
        public double NewPassenger2018 { get; private set; }

        public AddTransportDialog()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Добавление вида транспорта";
            this.Size = new Size(450, 350);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.FromArgb(240, 240, 240);

            mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 8,
                Padding = new Padding(15),
                BackColor = Color.FromArgb(240, 240, 240)
            };

            mainLayout.ColumnStyles.Clear();
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            for (int i = 0; i < 7; i++)
            {
                mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 35));
            }
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));

            // Название
            mainLayout.Controls.Add(new Label { Text = "Вид транспорта:", Font = new Font("Segoe UI", 10, FontStyle.Bold), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleRight, Margin = new Padding(5) }, 0, 0);
            txtName = new TextBox { Dock = DockStyle.Fill, Font = new Font("Segoe UI", 10), Margin = new Padding(5) };
            mainLayout.Controls.Add(txtName, 1, 0);

            // Грузооборот
            mainLayout.Controls.Add(new Label { Text = "Грузооборот 2011:", Font = new Font("Segoe UI", 10), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleRight, Margin = new Padding(5) }, 0, 1);
            txtCargo2011 = new TextBox { Dock = DockStyle.Fill, Font = new Font("Segoe UI", 10), Margin = new Padding(5), Text = "0" };
            mainLayout.Controls.Add(txtCargo2011, 1, 1);

            mainLayout.Controls.Add(new Label { Text = "Грузооборот 2013:", Font = new Font("Segoe UI", 10), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleRight, Margin = new Padding(5) }, 0, 2);
            txtCargo2013 = new TextBox { Dock = DockStyle.Fill, Font = new Font("Segoe UI", 10), Margin = new Padding(5), Text = "0" };
            mainLayout.Controls.Add(txtCargo2013, 1, 2);

            mainLayout.Controls.Add(new Label { Text = "Грузооборот 2015:", Font = new Font("Segoe UI", 10), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleRight, Margin = new Padding(5) }, 0, 3);
            txtCargo2015 = new TextBox { Dock = DockStyle.Fill, Font = new Font("Segoe UI", 10), Margin = new Padding(5), Text = "0" };
            mainLayout.Controls.Add(txtCargo2015, 1, 3);

            // Пассажирооборот
            mainLayout.Controls.Add(new Label { Text = "Пассажирооборот 2013:", Font = new Font("Segoe UI", 10), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleRight, Margin = new Padding(5) }, 0, 4);
            txtPassenger2013 = new TextBox { Dock = DockStyle.Fill, Font = new Font("Segoe UI", 10), Margin = new Padding(5), Text = "0" };
            mainLayout.Controls.Add(txtPassenger2013, 1, 4);

            mainLayout.Controls.Add(new Label { Text = "Пассажирооборот 2017:", Font = new Font("Segoe UI", 10), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleRight, Margin = new Padding(5) }, 0, 5);
            txtPassenger2017 = new TextBox { Dock = DockStyle.Fill, Font = new Font("Segoe UI", 10), Margin = new Padding(5), Text = "0" };
            mainLayout.Controls.Add(txtPassenger2017, 1, 5);

            mainLayout.Controls.Add(new Label { Text = "Пассажирооборот 2018:", Font = new Font("Segoe UI", 10), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleRight, Margin = new Padding(5) }, 0, 6);
            txtPassenger2018 = new TextBox { Dock = DockStyle.Fill, Font = new Font("Segoe UI", 10), Margin = new Padding(5), Text = "0" };
            mainLayout.Controls.Add(txtPassenger2018, 1, 6);

            // Кнопки
            FlowLayoutPanel buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                BackColor = Color.FromArgb(240, 240, 240),
                Padding = new Padding(5)
            };

            btnOk = new Button
            {
                Text = "OK",
                Width = 80,
                Height = 32,
                BackColor = Color.FromArgb(46, 204, 113),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Margin = new Padding(5)
            };
            btnOk.Click += BtnOk_Click;
            buttonPanel.Controls.Add(btnOk);

            btnCancel = new Button
            {
                Text = "Отмена",
                Width = 80,
                Height = 32,
                BackColor = Color.FromArgb(149, 165, 166),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Margin = new Padding(5)
            };
            btnCancel.Click += BtnCancel_Click;
            buttonPanel.Controls.Add(btnCancel);

            mainLayout.Controls.Add(buttonPanel, 1, 7);

            this.Controls.Add(mainLayout);
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            try
            {
                NewName = txtName.Text.Trim();

                if (string.IsNullOrEmpty(NewName))
                {
                    MessageBox.Show("Введите вид транспорта", "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                NewCargo2011 = double.Parse(txtCargo2011.Text);
                NewCargo2013 = double.Parse(txtCargo2013.Text);
                NewCargo2015 = double.Parse(txtCargo2015.Text);
                NewPassenger2013 = double.Parse(txtPassenger2013.Text);
                NewPassenger2017 = double.Parse(txtPassenger2017.Text);
                NewPassenger2018 = double.Parse(txtPassenger2018.Text);

                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (FormatException)
            {
                MessageBox.Show("Введите корректные числовые значения",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}