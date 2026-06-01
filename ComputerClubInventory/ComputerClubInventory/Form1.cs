using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

namespace ComputerClubInventory
{
    public partial class Form1 : Form
    {
        private int currentUserId;
        private int currentRoleId;
        private string currentUserName;
        private string currentRoleName;
        private string currentReportTitle = "Отчёт";
        private string printText = "";
        private int printPosition;
        private TabPage usersPage;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            usersPage = tabUsers;
            dtpReportFrom.Value = DateTime.Today.AddMonths(-1);
            dtpReportTo.Value = DateTime.Today;
            LoadRequestFilter();

        }

        private void LoginTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btnLogin.PerformClick();
            }
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            string login = txtLogin.Text.Trim();
            string password = txtPassword.Text;
            if (login.Length == 0 || password.Length == 0)
            {
                lblLoginMessage.Text = "Введите логин и пароль";
                return;
            }

            DataTable table = Database.Query(@"
SELECT u.id, u.full_name, u.role_id, r.name AS role_name
FROM users u
INNER JOIN roles r ON r.id = u.role_id
WHERE u.login = @login AND u.password_hash = @password_hash AND u.is_active = 1",
                P("@login", login),
                P("@password_hash", HashPassword(password)));

            if (table.Rows.Count == 0)
            {
                lblLoginMessage.Text = "Неверный логин или пароль";
                return;
            }

            currentUserId = Convert.ToInt32(table.Rows[0]["id"]);
            currentRoleId = Convert.ToInt32(table.Rows[0]["role_id"]);
            currentUserName = Convert.ToString(table.Rows[0]["full_name"]);
            currentRoleName = Convert.ToString(table.Rows[0]["role_name"]);
            lblCurrentUser.Text = "Пользователь: " + currentUserName + " | Роль: " + currentRoleName;
            lblLoginMessage.Text = "";
            loginPanel.Visible = false;
            mainPanel.Visible = true;
            ApplyPermissions();
            LoadAllData();
        }

        private void btnLogout_Click(object sender, EventArgs e)
        {
            currentUserId = 0;
            currentRoleId = 0;
            currentUserName = "";
            currentRoleName = "";
            mainPanel.Visible = false;
            loginPanel.Visible = true;
            txtPassword.Text = "";
            txtPassword.Focus();
        }

        private void ApplyPermissions()
        {
            bool owner = currentRoleId == 1;
            bool admin = currentRoleId == 2;
            bool tech = currentRoleId == 3;

            btnAddComputer.Enabled = owner || admin;
            btnEditComputer.Enabled = owner || admin;
            btnDeleteComputer.Enabled = owner;
            btnAddEquipment.Enabled = owner || admin;
            btnEditEquipment.Enabled = owner || admin;
            btnDeleteEquipment.Enabled = owner;
            btnAddFailure.Enabled = owner || admin;
            btnEditFailure.Enabled = owner || admin;
            btnDeleteFailure.Enabled = owner;
            btnAddRequest.Enabled = owner || admin;
            btnEditRequest.Enabled = owner || admin || tech;
            btnDeleteRequest.Enabled = owner;
            btnAddHistory.Enabled = owner || admin || tech;
            btnEditHistory.Enabled = owner || admin || tech;
            btnDeleteHistory.Enabled = owner;
            btnAddUser.Enabled = owner;
            btnEditUser.Enabled = owner;
            btnDeleteUser.Enabled = owner;

            if (owner)
            {
                if (!tabControl.TabPages.Contains(usersPage))
                {
                    tabControl.TabPages.Add(usersPage);
                }
            }
            else
            {
                if (tabControl.TabPages.Contains(usersPage))
                {
                    tabControl.TabPages.Remove(usersPage);
                }
            }
        }

        private void LoadAllData()
        {
            LoadComputers();
            LoadEquipment();
            LoadFailures();
            LoadRequestFilter();
            LoadRequests();
            LoadHistory();
            LoadUsers();
            LoadEquipmentReport();
        }

        private SQLiteParameter P(string name, object value)
        {
            if (value == null)
            {
                value = DBNull.Value;
            }
            return new SQLiteParameter(name, value);
        }

        private string HashPassword(string password)
        {
            using (SHA256 sha = SHA256.Create())
            {
                byte[] bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(password));
                StringBuilder builder = new StringBuilder();
                for (int i = 0; i < bytes.Length; i++)
                {
                    builder.Append(bytes[i].ToString("x2"));
                }
                return builder.ToString();
            }
        }

        private int SelectedId(DataGridView grid)
        {
            if (grid.CurrentRow == null || grid.CurrentRow.Cells["id"].Value == DBNull.Value)
            {
                return 0;
            }
            return Convert.ToInt32(grid.CurrentRow.Cells["id"].Value);
        }

        private void HideColumns(DataGridView grid, params string[] names)
        {
            foreach (string name in names)
            {
                if (grid.Columns.Contains(name))
                {
                    grid.Columns[name].Visible = false;
                }
            }
        }

        private string LikeText(string value)
        {
            return "%" + value.Trim() + "%";
        }

        private string DateSql(DateTime value)
        {
            return value.ToString("yyyy-MM-dd");
        }

        private string DateTimeSql(DateTime value)
        {
            return value.ToString("yyyy-MM-dd HH:mm:ss");
        }

        private object ComboValue(ComboBox combo)
        {
            if (combo.SelectedValue == null || combo.SelectedValue == DBNull.Value)
            {
                return DBNull.Value;
            }
            return combo.SelectedValue;
        }

        private DataTable Lookup(string sql)
        {
            return Database.Query(sql);
        }

        private ComboBox CreateCombo(DataTable table)
        {
            ComboBox combo = new ComboBox();
            combo.DropDownStyle = ComboBoxStyle.DropDownList;
            combo.DataSource = table;
            combo.DisplayMember = "name";
            combo.ValueMember = "id";
            combo.Width = 310;
            return combo;
        }

        private TextBox CreateText(string value)
        {
            TextBox textBox = new TextBox();
            textBox.Width = 310;
            textBox.Text = value;
            return textBox;
        }

        private TextBox CreateMultilineText(string value)
        {
            TextBox textBox = new TextBox();
            textBox.Width = 310;
            textBox.Height = 70;
            textBox.Multiline = true;
            textBox.ScrollBars = ScrollBars.Vertical;
            textBox.Text = value;
            return textBox;
        }

        private DateTimePicker CreateDate(DateTime value)
        {
            DateTimePicker picker = new DateTimePicker();
            picker.Format = DateTimePickerFormat.Short;
            picker.Width = 160;
            picker.Value = value;
            return picker;
        }

        private DateTimePicker CreateNullableDate(DateTime? value)
        {
            DateTimePicker picker = new DateTimePicker();
            picker.Format = DateTimePickerFormat.Short;
            picker.Width = 160;
            picker.ShowCheckBox = true;
            if (value.HasValue)
            {
                picker.Value = value.Value;
                picker.Checked = true;
            }
            else
            {
                picker.Value = DateTime.Today;
                picker.Checked = false;
            }
            return picker;
        }

        private Form CreateEditForm(string title, int height)
        {
            Form form = new Form();
            form.Text = title;
            form.StartPosition = FormStartPosition.CenterParent;
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.MaximizeBox = false;
            form.MinimizeBox = false;
            form.Width = 530;
            form.Height = height;
            return form;
        }

        private TableLayoutPanel CreateTable(Form form, int rows)
        {
            TableLayoutPanel table = new TableLayoutPanel();
            table.Dock = DockStyle.Fill;
            table.Padding = new Padding(12);
            table.ColumnCount = 2;
            table.RowCount = rows + 1;
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            form.Controls.Add(table);
            return table;
        }

        private void AddRow(TableLayoutPanel table, int row, string labelText, Control control)
        {
            Label label = new Label();
            label.Text = labelText;
            label.TextAlign = ContentAlignment.MiddleRight;
            label.Dock = DockStyle.Fill;
            table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            table.Controls.Add(label, 0, row);
            table.Controls.Add(control, 1, row);
        }

        private FlowLayoutPanel AddButtons(TableLayoutPanel table, int row, Form form)
        {
            FlowLayoutPanel panel = new FlowLayoutPanel();
            panel.FlowDirection = FlowDirection.RightToLeft;
            panel.Dock = DockStyle.Fill;
            Button ok = new Button();
            ok.Text = "Сохранить";
            ok.Width = 100;
            ok.DialogResult = DialogResult.OK;
            Button cancel = new Button();
            cancel.Text = "Отмена";
            cancel.Width = 100;
            cancel.DialogResult = DialogResult.Cancel;
            panel.Controls.Add(ok);
            panel.Controls.Add(cancel);
            table.Controls.Add(panel, 1, row);
            form.AcceptButton = ok;
            form.CancelButton = cancel;
            return panel;
        }

        private bool Empty(string text)
        {
            return text == null || text.Trim().Length == 0;
        }

        private DateTime? ToDate(object value)
        {
            if (value == null || value == DBNull.Value || Convert.ToString(value).Length == 0)
            {
                return null;
            }
            DateTime result;
            if (DateTime.TryParse(Convert.ToString(value), out result))
            {
                return result;
            }
            return null;
        }

        private void SelectCombo(ComboBox combo, object value)
        {
            if (combo.Items.Count == 0)
            {
                combo.SelectedIndex = -1;
                return;
            }

            if (value == null || value == DBNull.Value)
            {
                combo.SelectedIndex = -1;
                return;
            }

            int id;
            if (int.TryParse(value.ToString(), out id))
            {
                combo.SelectedValue = id;

                if (combo.SelectedIndex < 0 && combo.Items.Count > 0)
                {
                    combo.SelectedIndex = 0;
                }
            }
            else
            {
                combo.SelectedIndex = -1;
            }
        }

        private void DataGrid_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            DataGridView grid = sender as DataGridView;
            if (grid == null || !grid.Columns.Contains("color"))
            {
                return;
            }

            object value = grid.Rows[e.RowIndex].Cells["color"].Value;
            if (value == null || value == DBNull.Value)
            {
                return;
            }

            try
            {
                grid.Rows[e.RowIndex].DefaultCellStyle.BackColor = ColorTranslator.FromHtml(Convert.ToString(value));
            }
            catch
            {
            }
        }

        private void LoadComputers()
        {
            DataTable table = Database.Query(@"
SELECT c.id,
       c.number AS 'Номер',
       c.place AS 'Место',
       s.name AS 'Состояние',
       s.color AS color,
       c.installation_date AS 'Дата установки',
       c.specifications AS 'Характеристики',
       c.note AS 'Примечание'
FROM computers c
INNER JOIN computer_statuses s ON s.id = c.status_id
WHERE c.number LIKE @search OR c.place LIKE @search OR c.specifications LIKE @search OR s.name LIKE @search
ORDER BY c.number", P("@search", LikeText(txtComputerSearch.Text)));
            dgvComputers.DataSource = table;
            HideColumns(dgvComputers, "id", "color");
        }

        private void LoadEquipment()
        {
            DataTable table = Database.Query(@"
SELECT e.id,
       e.inventory_number AS 'Инв. номер',
       e.name AS 'Наименование',
       t.name AS 'Тип',
       e.model AS 'Модель',
       e.serial_number AS 'Серийный номер',
       IFNULL(c.number || ' / ' || c.place, '') AS 'Компьютер',
       e.location AS 'Размещение',
       s.name AS 'Состояние',
       s.color AS color,
       e.installation_date AS 'Дата установки',
       e.responsible_person AS 'Ответственный',
       e.cost AS 'Стоимость'
FROM equipment e
INNER JOIN equipment_types t ON t.id = e.type_id
INNER JOIN equipment_statuses s ON s.id = e.status_id
LEFT JOIN computers c ON c.id = e.computer_id
WHERE e.inventory_number LIKE @search OR e.name LIKE @search OR e.model LIKE @search OR e.serial_number LIKE @search OR t.name LIKE @search OR e.location LIKE @search OR s.name LIKE @search
ORDER BY e.inventory_number", P("@search", LikeText(txtEquipmentSearch.Text)));
            dgvEquipment.DataSource = table;
            HideColumns(dgvEquipment, "id", "color");
        }

        private void LoadFailures()
        {
            DataTable table = Database.Query(@"
SELECT f.id,
       f.detected_date AS 'Дата',
       e.inventory_number AS 'Инв. номер',
       e.name AS 'Оборудование',
       t.name AS 'Тип неисправности',
       f.criticality AS 'Критичность',
       CASE WHEN f.is_closed = 1 THEN 'Закрыта' ELSE 'Открыта' END AS 'Состояние',
       u.full_name AS 'Выявил',
       f.source AS 'Источник',
       f.description AS 'Описание'
FROM failures f
INNER JOIN equipment e ON e.id = f.equipment_id
INNER JOIN failure_types t ON t.id = f.failure_type_id
INNER JOIN users u ON u.id = f.detected_by_user_id
LEFT JOIN computers c ON c.id = f.computer_id
WHERE e.inventory_number LIKE @search OR e.name LIKE @search OR t.name LIKE @search OR f.description LIKE @search OR f.criticality LIKE @search OR u.full_name LIKE @search
ORDER BY f.detected_date DESC, f.id DESC", P("@search", LikeText(txtFailureSearch.Text)));
            dgvFailures.DataSource = table;
            HideColumns(dgvFailures, "id");
        }

        private void LoadRequestFilter()
        {
            DataTable table = Database.Query(@"
SELECT 0 AS id, 'Все' AS name
UNION ALL
SELECT id, name FROM request_statuses ORDER BY id");
            cmbRequestFilter.DataSource = table;
            cmbRequestFilter.DisplayMember = "name";
            cmbRequestFilter.ValueMember = "id";
        }

        private void LoadRequests()
        {
            int status = 0;
            if (cmbRequestFilter.SelectedValue != null && cmbRequestFilter.SelectedValue != DBNull.Value)
            {
                int.TryParse(Convert.ToString(cmbRequestFilter.SelectedValue), out status);
            }
            string statusSql = status == 0 ? "" : " AND r.status_id = @status";
            DataTable table = Database.Query(@"
SELECT r.id,
       r.created_date AS 'Дата создания',
       e.inventory_number AS 'Инв. номер',
       e.name AS 'Оборудование',
       rs.name AS 'Статус',
       IFNULL(u.full_name, '') AS 'Ответственный',
       f.description AS 'Неисправность',
       r.work_description AS 'Выполненные работы',
       r.result AS 'Результат',
       IFNULL(r.closed_date, '') AS 'Дата закрытия'
FROM requests r
INNER JOIN failures f ON f.id = r.failure_id
INNER JOIN equipment e ON e.id = r.equipment_id
INNER JOIN request_statuses rs ON rs.id = r.status_id
LEFT JOIN users u ON u.id = r.responsible_user_id
WHERE (e.inventory_number LIKE @search OR e.name LIKE @search OR rs.name LIKE @search OR IFNULL(u.full_name, '') LIKE @search OR f.description LIKE @search OR r.work_description LIKE @search OR r.result LIKE @search)" + statusSql + @"
ORDER BY r.created_date DESC, r.id DESC", P("@search", LikeText(txtRequestSearch.Text)), P("@status", status));
            dgvRequests.DataSource = table;
            HideColumns(dgvRequests, "id");
        }

        private void LoadHistory()
        {
            DataTable table = Database.Query(@"
SELECT h.id,
       h.service_date AS 'Дата',
       e.inventory_number AS 'Инв. номер',
       e.name AS 'Оборудование',
       h.action AS 'Действие',
       IFNULL(u.full_name, '') AS 'Сотрудник',
       COALESCE(CAST(r.id AS TEXT), '') AS 'Заявка',
       h.result AS 'Результат'
FROM service_history h
INNER JOIN equipment e ON e.id = h.equipment_id
LEFT JOIN users u ON u.id = h.employee_user_id
LEFT JOIN requests r ON r.id = h.request_id
WHERE e.inventory_number LIKE @search 
   OR e.name LIKE @search 
   OR h.action LIKE @search 
   OR h.result LIKE @search 
   OR IFNULL(u.full_name, '') LIKE @search
   OR COALESCE(CAST(r.id AS TEXT), '') LIKE @search
ORDER BY h.service_date DESC, h.id DESC", P("@search", LikeText(txtHistorySearch.Text)));

            dgvHistory.DataSource = table;
            HideColumns(dgvHistory, "id");
        }

        private void LoadUsers()
        {
            if (currentRoleId != 1)
            {
                return;
            }

            DataTable table = Database.Query(@"
SELECT u.id,
       u.login AS 'Логин',
       u.full_name AS 'ФИО',
       r.name AS 'Роль',
       CASE WHEN u.is_active = 1 THEN 'Активен' ELSE 'Отключён' END AS 'Состояние',
       u.created_at AS 'Создан'
FROM users u
INNER JOIN roles r ON r.id = u.role_id
WHERE u.login LIKE @search OR u.full_name LIKE @search OR r.name LIKE @search
ORDER BY u.id", P("@search", LikeText(txtUserSearch.Text)));
            dgvUsers.DataSource = table;
            HideColumns(dgvUsers, "id");
        }

        private void txtComputerSearch_TextChanged(object sender, EventArgs e)
        {
            LoadComputers();
        }

        private void txtEquipmentSearch_TextChanged(object sender, EventArgs e)
        {
            LoadEquipment();
        }

        private void txtFailureSearch_TextChanged(object sender, EventArgs e)
        {
            LoadFailures();
        }

        private void txtRequestSearch_TextChanged(object sender, EventArgs e)
        {
            LoadRequests();
        }

        private void txtHistorySearch_TextChanged(object sender, EventArgs e)
        {
            LoadHistory();
        }

        private void txtUserSearch_TextChanged(object sender, EventArgs e)
        {
            LoadUsers();
        }

        private void cmbRequestFilter_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (mainPanel.Visible)
            {
                LoadRequests();
            }
        }

        private void btnRefreshComputers_Click(object sender, EventArgs e)
        {
            LoadComputers();
        }

        private void btnRefreshEquipment_Click(object sender, EventArgs e)
        {
            LoadEquipment();
        }

        private void btnRefreshFailures_Click(object sender, EventArgs e)
        {
            LoadFailures();
        }

        private void btnRefreshRequests_Click(object sender, EventArgs e)
        {
            LoadRequests();
        }

        private void btnRefreshHistory_Click(object sender, EventArgs e)
        {
            LoadHistory();
        }

        private void btnRefreshUsers_Click(object sender, EventArgs e)
        {
            LoadUsers();
        }

        private void btnAddComputer_Click(object sender, EventArgs e)
        {
            ShowComputerDialog(0);
        }

        private void btnEditComputer_Click(object sender, EventArgs e)
        {
            ShowComputerDialog(SelectedId(dgvComputers));
        }

        private void dgvComputers_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && btnEditComputer.Enabled)
            {
                ShowComputerDialog(SelectedId(dgvComputers));
            }
        }

        private void btnDeleteComputer_Click(object sender, EventArgs e)
        {
            int id = SelectedId(dgvComputers);
            if (id == 0)
            {
                return;
            }
            if (MessageBox.Show("Удалить выбранный компьютер?", "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                return;
            }
            try
            {
                Database.Execute("DELETE FROM computers WHERE id = @id", P("@id", id));
                LoadAllData();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Нельзя удалить компьютер, пока к нему привязано оборудование или записи.\n" + ex.Message);
            }
        }

        private void ShowComputerDialog(int id)
        {
            DataRow row = null;
            if (id > 0)
            {
                DataTable data = Database.Query("SELECT * FROM computers WHERE id = @id", P("@id", id));
                if (data.Rows.Count == 0)
                {
                    return;
                }
                row = data.Rows[0];
            }

            Form form = CreateEditForm(id == 0 ? "Добавление компьютера" : "Редактирование компьютера", 420);
            TableLayoutPanel table = CreateTable(form, 7);
            TextBox txtNumber = CreateText(row == null ? "" : Convert.ToString(row["number"]));
            TextBox txtPlace = CreateText(row == null ? "" : Convert.ToString(row["place"]));
            TextBox txtSpecs = CreateMultilineText(row == null ? "" : Convert.ToString(row["specifications"]));
            ComboBox cmbStatus = CreateCombo(Lookup("SELECT id, name FROM computer_statuses ORDER BY id"));
            DateTimePicker dtpInstall = CreateDate(row == null ? DateTime.Today : Convert.ToDateTime(row["installation_date"]));
            TextBox txtNote = CreateMultilineText(row == null ? "" : Convert.ToString(row["note"]));

            AddRow(table, 0, "Номер", txtNumber);
            AddRow(table, 1, "Место", txtPlace);
            AddRow(table, 2, "Характеристики", txtSpecs);
            AddRow(table, 3, "Состояние", cmbStatus);
            AddRow(table, 4, "Дата установки", dtpInstall);
            AddRow(table, 5, "Примечание", txtNote);
            AddButtons(table, 6, form);

            if (row != null)
            {
                SelectCombo(cmbStatus, row["status_id"]);
            }

            if (form.ShowDialog(this) != DialogResult.OK)
            {
                return;
            }
            if (Empty(txtNumber.Text) || Empty(txtPlace.Text) || Empty(txtSpecs.Text))
            {
                MessageBox.Show("Заполните номер, место и характеристики");
                return;
            }

            if (id == 0)
            {
                Database.Execute(@"
INSERT INTO computers(number, place, specifications, status_id, installation_date, note)
VALUES(@number, @place, @specifications, @status_id, @installation_date, @note)",
                    P("@number", txtNumber.Text.Trim()),
                    P("@place", txtPlace.Text.Trim()),
                    P("@specifications", txtSpecs.Text.Trim()),
                    P("@status_id", ComboValue(cmbStatus)),
                    P("@installation_date", DateSql(dtpInstall.Value)),
                    P("@note", txtNote.Text.Trim()));
            }
            else
            {
                Database.Execute(@"
UPDATE computers
SET number = @number, place = @place, specifications = @specifications, status_id = @status_id, installation_date = @installation_date, note = @note
WHERE id = @id",
                    P("@number", txtNumber.Text.Trim()),
                    P("@place", txtPlace.Text.Trim()),
                    P("@specifications", txtSpecs.Text.Trim()),
                    P("@status_id", ComboValue(cmbStatus)),
                    P("@installation_date", DateSql(dtpInstall.Value)),
                    P("@note", txtNote.Text.Trim()),
                    P("@id", id));
            }
            LoadAllData();
        }

        private void btnAddEquipment_Click(object sender, EventArgs e)
        {
            ShowEquipmentDialog(0);
        }

        private void btnEditEquipment_Click(object sender, EventArgs e)
        {
            ShowEquipmentDialog(SelectedId(dgvEquipment));
        }

        private void dgvEquipment_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && btnEditEquipment.Enabled)
            {
                ShowEquipmentDialog(SelectedId(dgvEquipment));
            }
        }

        private void btnDeleteEquipment_Click(object sender, EventArgs e)
        {
            int id = SelectedId(dgvEquipment);
            if (id == 0)
            {
                return;
            }
            if (MessageBox.Show("Удалить выбранное оборудование?", "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                return;
            }
            try
            {
                Database.Execute("DELETE FROM equipment WHERE id = @id", P("@id", id));
                LoadAllData();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось удалить оборудование.\n" + ex.Message);
            }
        }

        private void ShowEquipmentDialog(int id)
        {
            DataRow row = null;
            if (id > 0)
            {
                DataTable data = Database.Query("SELECT * FROM equipment WHERE id = @id", P("@id", id));
                if (data.Rows.Count == 0)
                {
                    return;
                }
                row = data.Rows[0];
            }

            Form form = CreateEditForm(id == 0 ? "Добавление оборудования" : "Редактирование оборудования", 690);
            TableLayoutPanel table = CreateTable(form, 15);
            TextBox txtInventory = CreateText(row == null ? "" : Convert.ToString(row["inventory_number"]));
            TextBox txtName = CreateText(row == null ? "" : Convert.ToString(row["name"]));
            TextBox txtModel = CreateText(row == null ? "" : Convert.ToString(row["model"]));
            TextBox txtSerial = CreateText(row == null ? "" : Convert.ToString(row["serial_number"]));
            ComboBox cmbType = CreateCombo(Lookup("SELECT id, name FROM equipment_types ORDER BY name"));
            ComboBox cmbComputer = CreateCombo(Lookup("SELECT NULL AS id, 'Не привязано' AS name UNION ALL SELECT id, number || ' / ' || place AS name FROM computers ORDER BY name"));
            DateTimePicker dtpPurchase = CreateNullableDate(row == null ? null : ToDate(row["purchase_date"]));
            DateTimePicker dtpInstall = CreateDate(row == null ? DateTime.Today : Convert.ToDateTime(row["installation_date"]));
            TextBox txtCost = CreateText(row == null ? "0" : Convert.ToString(row["cost"], CultureInfo.InvariantCulture));
            TextBox txtResponsible = CreateText(row == null ? "" : Convert.ToString(row["responsible_person"]));
            TextBox txtLocation = CreateText(row == null ? "" : Convert.ToString(row["location"]));
            TextBox txtSpecs = CreateMultilineText(row == null ? "" : Convert.ToString(row["specifications"]));
            ComboBox cmbStatus = CreateCombo(Lookup("SELECT id, name FROM equipment_statuses ORDER BY id"));
            TextBox txtNote = CreateMultilineText(row == null ? "" : Convert.ToString(row["note"]));

            AddRow(table, 0, "Инвентарный номер", txtInventory);
            AddRow(table, 1, "Наименование", txtName);
            AddRow(table, 2, "Модель", txtModel);
            AddRow(table, 3, "Серийный номер", txtSerial);
            AddRow(table, 4, "Тип", cmbType);
            AddRow(table, 5, "Компьютер", cmbComputer);
            AddRow(table, 6, "Дата покупки", dtpPurchase);
            AddRow(table, 7, "Дата установки", dtpInstall);
            AddRow(table, 8, "Стоимость", txtCost);
            AddRow(table, 9, "Ответственный", txtResponsible);
            AddRow(table, 10, "Размещение", txtLocation);
            AddRow(table, 11, "Характеристики", txtSpecs);
            AddRow(table, 12, "Состояние", cmbStatus);
            AddRow(table, 13, "Примечание", txtNote);
            AddButtons(table, 14, form);

            if (row != null)
            {
                SelectCombo(cmbType, row["type_id"]);
                SelectCombo(cmbComputer, row["computer_id"]);
                SelectCombo(cmbStatus, row["status_id"]);
            }

            if (form.ShowDialog(this) != DialogResult.OK)
            {
                return;
            }
            if (Empty(txtInventory.Text) || Empty(txtName.Text) || Empty(txtModel.Text) || Empty(txtSerial.Text) || Empty(txtResponsible.Text) || Empty(txtLocation.Text) || Empty(txtSpecs.Text))
            {
                MessageBox.Show("Заполните обязательные поля оборудования");
                return;
            }

            decimal cost;
            if (!decimal.TryParse(txtCost.Text.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out cost))
            {
                MessageBox.Show("Стоимость введена неверно");
                return;
            }

            object purchaseDate = dtpPurchase.Checked ? (object)DateSql(dtpPurchase.Value) : DBNull.Value;
            if (id == 0)
            {
                Database.Execute(@"
INSERT INTO equipment(inventory_number, name, model, serial_number, type_id, computer_id, purchase_date, installation_date, cost, responsible_person, location, specifications, status_id, note)
VALUES(@inventory_number, @name, @model, @serial_number, @type_id, @computer_id, @purchase_date, @installation_date, @cost, @responsible_person, @location, @specifications, @status_id, @note)",
                    P("@inventory_number", txtInventory.Text.Trim()),
                    P("@name", txtName.Text.Trim()),
                    P("@model", txtModel.Text.Trim()),
                    P("@serial_number", txtSerial.Text.Trim()),
                    P("@type_id", ComboValue(cmbType)),
                    P("@computer_id", ComboValue(cmbComputer)),
                    P("@purchase_date", purchaseDate),
                    P("@installation_date", DateSql(dtpInstall.Value)),
                    P("@cost", cost),
                    P("@responsible_person", txtResponsible.Text.Trim()),
                    P("@location", txtLocation.Text.Trim()),
                    P("@specifications", txtSpecs.Text.Trim()),
                    P("@status_id", ComboValue(cmbStatus)),
                    P("@note", txtNote.Text.Trim()));
            }
            else
            {
                Database.Execute(@"
UPDATE equipment
SET inventory_number = @inventory_number, name = @name, model = @model, serial_number = @serial_number, type_id = @type_id, computer_id = @computer_id, purchase_date = @purchase_date, installation_date = @installation_date, cost = @cost, responsible_person = @responsible_person, location = @location, specifications = @specifications, status_id = @status_id, note = @note
WHERE id = @id",
                    P("@inventory_number", txtInventory.Text.Trim()),
                    P("@name", txtName.Text.Trim()),
                    P("@model", txtModel.Text.Trim()),
                    P("@serial_number", txtSerial.Text.Trim()),
                    P("@type_id", ComboValue(cmbType)),
                    P("@computer_id", ComboValue(cmbComputer)),
                    P("@purchase_date", purchaseDate),
                    P("@installation_date", DateSql(dtpInstall.Value)),
                    P("@cost", cost),
                    P("@responsible_person", txtResponsible.Text.Trim()),
                    P("@location", txtLocation.Text.Trim()),
                    P("@specifications", txtSpecs.Text.Trim()),
                    P("@status_id", ComboValue(cmbStatus)),
                    P("@note", txtNote.Text.Trim()),
                    P("@id", id));
            }
            RecalculateComputerStatuses();
            LoadAllData();
        }

        private void btnAddFailure_Click(object sender, EventArgs e)
        {
            ShowFailureDialog(0);
        }

        private void btnEditFailure_Click(object sender, EventArgs e)
        {
            ShowFailureDialog(SelectedId(dgvFailures));
        }

        private void dgvFailures_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && btnEditFailure.Enabled)
            {
                ShowFailureDialog(SelectedId(dgvFailures));
            }
        }

        private void btnDeleteFailure_Click(object sender, EventArgs e)
        {
            int id = SelectedId(dgvFailures);
            if (id == 0)
            {
                return;
            }
            if (MessageBox.Show("Удалить выбранную неисправность и связанные заявки?", "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                return;
            }
            Database.Execute("DELETE FROM failures WHERE id = @id", P("@id", id));
            RecalculateComputerStatuses();
            LoadAllData();
        }

        private void ShowFailureDialog(int id)
        {
            DataRow row = null;
            if (id > 0)
            {
                DataTable data = Database.Query("SELECT * FROM failures WHERE id = @id", P("@id", id));
                if (data.Rows.Count == 0)
                {
                    return;
                }
                row = data.Rows[0];
            }

            Form form = CreateEditForm(id == 0 ? "Регистрация неисправности" : "Редактирование неисправности", 520);
            TableLayoutPanel table = CreateTable(form, 10);
            ComboBox cmbEquipment = CreateCombo(Lookup("SELECT id, inventory_number || ' / ' || name || ' / ' || model AS name FROM equipment ORDER BY inventory_number"));
            ComboBox cmbComputer = CreateCombo(Lookup("SELECT NULL AS id, 'Не указано' AS name UNION ALL SELECT id, number || ' / ' || place AS name FROM computers ORDER BY name"));
            ComboBox cmbType = CreateCombo(Lookup("SELECT id, name FROM failure_types ORDER BY name"));
            DateTimePicker dtpDate = CreateDate(row == null ? DateTime.Today : Convert.ToDateTime(row["detected_date"]));
            ComboBox cmbCriticality = new ComboBox();
            cmbCriticality.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbCriticality.Width = 310;
            cmbCriticality.Items.AddRange(new object[] { "Низкая", "Средняя", "Высокая", "Критическая" });
            cmbCriticality.SelectedIndex = 1;
            TextBox txtSource = CreateText(row == null ? "Администратор клуба" : Convert.ToString(row["source"]));
            TextBox txtDescription = CreateMultilineText(row == null ? "" : Convert.ToString(row["description"]));
            CheckBox chkClosed = new CheckBox();
            chkClosed.Text = "Закрыта";
            chkClosed.Checked = row != null && Convert.ToInt32(row["is_closed"]) == 1;

            AddRow(table, 0, "Оборудование", cmbEquipment);
            AddRow(table, 1, "Компьютер", cmbComputer);
            AddRow(table, 2, "Тип", cmbType);
            AddRow(table, 3, "Дата обнаружения", dtpDate);
            AddRow(table, 4, "Критичность", cmbCriticality);
            AddRow(table, 5, "Источник", txtSource);
            AddRow(table, 6, "Описание", txtDescription);
            AddRow(table, 7, "Состояние", chkClosed);

            Label info = new Label();
            info.Text = id == 0 ? "Будет автоматически создана заявка на ремонт" : "Изменение заявки выполняется на вкладке Заявки";
            
            info.Dock = DockStyle.Fill;
            info.TextAlign = ContentAlignment.MiddleLeft;
            table.Controls.Add(info, 1, 8);
            AddButtons(table, 9, form);

            if (row != null)
            {
                SelectCombo(cmbEquipment, row["equipment_id"]);
                SelectCombo(cmbComputer, row["computer_id"]);
                SelectCombo(cmbType, row["failure_type_id"]);
                cmbCriticality.SelectedItem = Convert.ToString(row["criticality"]);
            }

            if (form.ShowDialog(this) != DialogResult.OK)
            {
                return;
            }
            if (Empty(txtDescription.Text) || Empty(txtSource.Text))
            {
                MessageBox.Show("Заполните описание и источник неисправности");
                return;
            }

            if (id == 0)
            {
                using (SQLiteConnection connection = Database.OpenConnection())
                using (SQLiteTransaction transaction = connection.BeginTransaction())
                {
                    SQLiteCommand command = connection.CreateCommand();
                    command.Transaction = transaction;
                    command.CommandText = @"
INSERT INTO failures(equipment_id, computer_id, failure_type_id, detected_date, description, criticality, detected_by_user_id, source, is_closed)
VALUES(@equipment_id, @computer_id, @failure_type_id, @detected_date, @description, @criticality, @detected_by_user_id, @source, 0);";
                    command.Parameters.Add(P("@equipment_id", ComboValue(cmbEquipment)));
                    command.Parameters.Add(P("@computer_id", ComboValue(cmbComputer)));
                    command.Parameters.Add(P("@failure_type_id", ComboValue(cmbType)));
                    command.Parameters.Add(P("@detected_date", DateSql(dtpDate.Value)));
                    command.Parameters.Add(P("@description", txtDescription.Text.Trim()));
                    command.Parameters.Add(P("@criticality", Convert.ToString(cmbCriticality.SelectedItem)));
                    command.Parameters.Add(P("@detected_by_user_id", currentUserId));
                    command.Parameters.Add(P("@source", txtSource.Text.Trim()));
                    command.ExecuteNonQuery();

                    command.Parameters.Clear();
                    command.CommandText = "SELECT last_insert_rowid()";
                    long failureId = (long)command.ExecuteScalar();

                    command.CommandText = "UPDATE equipment SET status_id = 2 WHERE id = @equipment_id";
                    command.Parameters.Add(P("@equipment_id", ComboValue(cmbEquipment)));
                    command.ExecuteNonQuery();

                    if (ComboValue(cmbComputer) != DBNull.Value)
                    {
                        command.Parameters.Clear();
                        command.CommandText = "UPDATE computers SET status_id = 2 WHERE id = @computer_id";
                        command.Parameters.Add(P("@computer_id", ComboValue(cmbComputer)));
                        command.ExecuteNonQuery();
                    }

                    command.Parameters.Clear();
                    command.CommandText = @"
INSERT INTO requests(failure_id, equipment_id, created_date, status_id, responsible_user_id, work_description, result)
VALUES(@failure_id, @equipment_id, @created_date, 1, NULL, '', '')";
                    command.Parameters.Add(P("@failure_id", failureId));
                    command.Parameters.Add(P("@equipment_id", ComboValue(cmbEquipment)));
                    command.Parameters.Add(P("@created_date", DateTimeSql(DateTime.Now)));
                    command.ExecuteNonQuery();
                    transaction.Commit();
                }
            }
            else
            {
                Database.Execute(@"
UPDATE failures
SET equipment_id = @equipment_id, computer_id = @computer_id, failure_type_id = @failure_type_id, detected_date = @detected_date, description = @description, criticality = @criticality, source = @source, is_closed = @is_closed
WHERE id = @id",
                    P("@equipment_id", ComboValue(cmbEquipment)),
                    P("@computer_id", ComboValue(cmbComputer)),
                    P("@failure_type_id", ComboValue(cmbType)),
                    P("@detected_date", DateSql(dtpDate.Value)),
                    P("@description", txtDescription.Text.Trim()),
                    P("@criticality", Convert.ToString(cmbCriticality.SelectedItem)),
                    P("@source", txtSource.Text.Trim()),
                    P("@is_closed", chkClosed.Checked ? 1 : 0),
                    P("@id", id));
                if (!chkClosed.Checked)
                {
                    Database.Execute("UPDATE equipment SET status_id = 2 WHERE id = @equipment_id", P("@equipment_id", ComboValue(cmbEquipment)));
                }
            }
            RecalculateComputerStatuses();
            LoadAllData();
        }

        private void btnAddRequest_Click(object sender, EventArgs e)
        {
            ShowRequestDialog(0);
        }

        private void btnEditRequest_Click(object sender, EventArgs e)
        {
            ShowRequestDialog(SelectedId(dgvRequests));
        }

        private void dgvRequests_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && btnEditRequest.Enabled)
            {
                ShowRequestDialog(SelectedId(dgvRequests));
            }
        }

        private void btnDeleteRequest_Click(object sender, EventArgs e)
        {
            int id = SelectedId(dgvRequests);
            if (id == 0)
            {
                return;
            }
            if (MessageBox.Show("Удалить выбранную заявку?", "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                return;
            }
            Database.Execute("DELETE FROM requests WHERE id = @id", P("@id", id));
            LoadAllData();
        }

        private void ShowRequestDialog(int id)
        {
            DataRow row = null;
            if (id > 0)
            {
                DataTable data = Database.Query("SELECT * FROM requests WHERE id = @id", P("@id", id));
                if (data.Rows.Count == 0)
                {
                    return;
                }
                row = data.Rows[0];
            }

            Form form = CreateEditForm(id == 0 ? "Создание заявки" : "Редактирование заявки", 570);
            TableLayoutPanel table = CreateTable(form, 10);
            ComboBox cmbFailure = CreateCombo(Lookup(@"
SELECT f.id, e.inventory_number || ' / ' || substr(f.description, 1, 70) AS name
FROM failures f
INNER JOIN equipment e ON e.id = f.equipment_id
ORDER BY f.detected_date DESC, f.id DESC"));
            ComboBox cmbEquipment = CreateCombo(Lookup("SELECT id, inventory_number || ' / ' || name || ' / ' || model AS name FROM equipment ORDER BY inventory_number"));
            ComboBox cmbStatus = CreateCombo(Lookup("SELECT id, name FROM request_statuses ORDER BY id"));
            ComboBox cmbResponsible = CreateCombo(Lookup("SELECT NULL AS id, 'Не назначен' AS name UNION ALL SELECT id, full_name AS name FROM users WHERE is_active = 1 ORDER BY name"));
            DateTimePicker dtpCreated = CreateDate(row == null ? DateTime.Today : Convert.ToDateTime(row["created_date"]));
            TextBox txtWork = CreateMultilineText(row == null ? "" : Convert.ToString(row["work_description"]));
            TextBox txtResult = CreateMultilineText(row == null ? "" : Convert.ToString(row["result"]));
            DateTimePicker dtpClosed = CreateNullableDate(row == null ? null : ToDate(row["closed_date"]));

            AddRow(table, 0, "Неисправность", cmbFailure);
            AddRow(table, 1, "Оборудование", cmbEquipment);
            AddRow(table, 2, "Дата создания", dtpCreated);
            AddRow(table, 3, "Статус", cmbStatus);
            AddRow(table, 4, "Ответственный", cmbResponsible);
            AddRow(table, 5, "Выполненные работы", txtWork);
            AddRow(table, 6, "Результат", txtResult);
            AddRow(table, 7, "Дата закрытия", dtpClosed);
            Label info = new Label();
            info.Text = "При статусе Завершена запись автоматически попадёт в историю обслуживания";
            info.Dock = DockStyle.Fill;
            table.Controls.Add(info, 1, 8);
            AddButtons(table, 9, form);

            cmbFailure.SelectedIndexChanged += delegate
            {
                if (id == 0 && cmbFailure.SelectedValue != null && cmbFailure.SelectedValue != DBNull.Value)
                {
                    object eq = Database.Scalar("SELECT equipment_id FROM failures WHERE id = @id", P("@id", cmbFailure.SelectedValue));
                    if (eq != null && eq != DBNull.Value)
                    {
                        cmbEquipment.SelectedValue = Convert.ToInt32(eq);
                    }
                }
            };

            if (row != null)
            {
                SelectCombo(cmbFailure, row["failure_id"]);
                SelectCombo(cmbEquipment, row["equipment_id"]);
                SelectCombo(cmbStatus, row["status_id"]);
                SelectCombo(cmbResponsible, row["responsible_user_id"]);
            }

            if (form.ShowDialog(this) != DialogResult.OK)
            {
                return;
            }

            int statusId = Convert.ToInt32(ComboValue(cmbStatus));
            object closedDate = DBNull.Value;
            if (statusId == 4 || statusId == 5)
            {
                closedDate = dtpClosed.Checked ? (object)DateSql(dtpClosed.Value) : DateSql(DateTime.Today);
            }

            if (id == 0)
            {
                Database.Execute(@"
INSERT INTO requests(failure_id, equipment_id, created_date, status_id, responsible_user_id, work_description, result, closed_date)
VALUES(@failure_id, @equipment_id, @created_date, @status_id, @responsible_user_id, @work_description, @result, @closed_date)",
                    P("@failure_id", ComboValue(cmbFailure)),
                    P("@equipment_id", ComboValue(cmbEquipment)),
                    P("@created_date", DateTimeSql(dtpCreated.Value)),
                    P("@status_id", statusId),
                    P("@responsible_user_id", ComboValue(cmbResponsible)),
                    P("@work_description", txtWork.Text.Trim()),
                    P("@result", txtResult.Text.Trim()),
                    P("@closed_date", closedDate));
                object newId = Database.Scalar("SELECT MAX(id) FROM requests");
                id = Convert.ToInt32(newId);
            }
            else
            {
                Database.Execute(@"
UPDATE requests
SET failure_id = @failure_id, equipment_id = @equipment_id, created_date = @created_date, status_id = @status_id, responsible_user_id = @responsible_user_id, work_description = @work_description, result = @result, closed_date = @closed_date
WHERE id = @id",
                    P("@failure_id", ComboValue(cmbFailure)),
                    P("@equipment_id", ComboValue(cmbEquipment)),
                    P("@created_date", DateTimeSql(dtpCreated.Value)),
                    P("@status_id", statusId),
                    P("@responsible_user_id", ComboValue(cmbResponsible)),
                    P("@work_description", txtWork.Text.Trim()),
                    P("@result", txtResult.Text.Trim()),
                    P("@closed_date", closedDate),
                    P("@id", id));
            }

            UpdateAfterRequest(id);
            RecalculateComputerStatuses();
            LoadAllData();
        }

        private void UpdateAfterRequest(int requestId)
        {
            DataTable data = Database.Query(@"
SELECT r.id, r.failure_id, r.equipment_id, r.status_id, r.responsible_user_id, r.work_description, r.result, e.computer_id
FROM requests r
INNER JOIN equipment e ON e.id = r.equipment_id
WHERE r.id = @id", P("@id", requestId));
            if (data.Rows.Count == 0)
            {
                return;
            }

            DataRow row = data.Rows[0];
            int statusId = Convert.ToInt32(row["status_id"]);
            int equipmentId = Convert.ToInt32(row["equipment_id"]);
            int failureId = Convert.ToInt32(row["failure_id"]);

            if (statusId == 4 || statusId == 5)
            {
                Database.Execute("UPDATE failures SET is_closed = 1 WHERE id = @id", P("@id", failureId));
                if (statusId == 4)
                {
                    Database.Execute("UPDATE equipment SET status_id = 1 WHERE id = @id", P("@id", equipmentId));
                }
                string action = statusId == 4 ? "Ремонт завершён" : "Заявка отменена";
                object count = Database.Scalar("SELECT COUNT(*) FROM service_history WHERE request_id = @request_id AND action = @action", P("@request_id", requestId), P("@action", action));
                if (Convert.ToInt32(count) == 0)
                {
                    string result = Convert.ToString(row["result"]);
                    string work = Convert.ToString(row["work_description"]);
                    if (result.Length == 0)
                    {
                        result = work;
                    }
                    Database.Execute(@"
INSERT INTO service_history(equipment_id, request_id, service_date, action, employee_user_id, result)
VALUES(@equipment_id, @request_id, @service_date, @action, @employee_user_id, @result)",
                        P("@equipment_id", equipmentId),
                        P("@request_id", requestId),
                        P("@service_date", DateTimeSql(DateTime.Now)),
                        P("@action", action),
                        P("@employee_user_id", row["responsible_user_id"] == DBNull.Value ? (object)currentUserId : row["responsible_user_id"]),
                        P("@result", result));
                }
            }
            else
            {
                Database.Execute("UPDATE failures SET is_closed = 0 WHERE id = @id", P("@id", failureId));
                Database.Execute("UPDATE equipment SET status_id = 2 WHERE id = @id", P("@id", equipmentId));
            }
        }

        private void btnAddHistory_Click(object sender, EventArgs e)
        {
            ShowHistoryDialog(0);
        }

        private void btnEditHistory_Click(object sender, EventArgs e)
        {
            ShowHistoryDialog(SelectedId(dgvHistory));
        }

        private void dgvHistory_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && btnEditHistory.Enabled)
            {
                ShowHistoryDialog(SelectedId(dgvHistory));
            }
        }

        private void btnDeleteHistory_Click(object sender, EventArgs e)
        {
            int id = SelectedId(dgvHistory);
            if (id == 0)
            {
                return;
            }
            if (MessageBox.Show("Удалить запись истории?", "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                return;
            }
            Database.Execute("DELETE FROM service_history WHERE id = @id", P("@id", id));
            LoadHistory();
        }

        private void ShowHistoryDialog(int id)
        {
            DataRow row = null;
            if (id > 0)
            {
                DataTable data = Database.Query("SELECT * FROM service_history WHERE id = @id", P("@id", id));
                if (data.Rows.Count == 0)
                {
                    return;
                }
                row = data.Rows[0];
            }

            Form form = CreateEditForm(id == 0 ? "Добавление записи истории" : "Редактирование записи истории", 480);
            TableLayoutPanel table = CreateTable(form, 8);
            ComboBox cmbEquipment = CreateCombo(Lookup("SELECT id, inventory_number || ' / ' || name || ' / ' || model AS name FROM equipment ORDER BY inventory_number"));
            ComboBox cmbRequest = CreateCombo(Lookup(@"
SELECT NULL AS id, 'Без заявки' AS name
UNION ALL
SELECT r.id, '#' || r.id || ' / ' || e.inventory_number || ' / ' || rs.name AS name
FROM requests r
INNER JOIN equipment e ON e.id = r.equipment_id
INNER JOIN request_statuses rs ON rs.id = r.status_id
ORDER BY name"));
            DateTimePicker dtpService = CreateDate(row == null ? DateTime.Today : Convert.ToDateTime(row["service_date"]));
            TextBox txtAction = CreateText(row == null ? "Плановое обслуживание" : Convert.ToString(row["action"]));
            ComboBox cmbEmployee = CreateCombo(Lookup("SELECT NULL AS id, 'Не указан' AS name UNION ALL SELECT id, full_name AS name FROM users WHERE is_active = 1 ORDER BY name"));
            TextBox txtResult = CreateMultilineText(row == null ? "" : Convert.ToString(row["result"]));

            AddRow(table, 0, "Оборудование", cmbEquipment);
            AddRow(table, 1, "Заявка", cmbRequest);
            AddRow(table, 2, "Дата", dtpService);
            AddRow(table, 3, "Действие", txtAction);
            AddRow(table, 4, "Сотрудник", cmbEmployee);
            AddRow(table, 5, "Результат", txtResult);
            Label info = new Label();
            info.Text = "Запись используется для просмотра истории обслуживания оборудования";
            info.Dock = DockStyle.Fill;
            table.Controls.Add(info, 1, 6);
            AddButtons(table, 7, form);

            if (row != null)
            {
                SelectCombo(cmbEquipment, row["equipment_id"]);
                SelectCombo(cmbRequest, row["request_id"]);
                SelectCombo(cmbEmployee, row["employee_user_id"]);
            }
            else
            {
                SelectCombo(cmbEmployee, currentUserId);
            }

            if (form.ShowDialog(this) != DialogResult.OK)
            {
                return;
            }
            if (Empty(txtAction.Text) || Empty(txtResult.Text))
            {
                MessageBox.Show("Заполните действие и результат");
                return;
            }

            if (id == 0)
            {
                Database.Execute(@"
INSERT INTO service_history(equipment_id, request_id, service_date, action, employee_user_id, result)
VALUES(@equipment_id, @request_id, @service_date, @action, @employee_user_id, @result)",
                    P("@equipment_id", ComboValue(cmbEquipment)),
                    P("@request_id", ComboValue(cmbRequest)),
                    P("@service_date", DateTimeSql(dtpService.Value)),
                    P("@action", txtAction.Text.Trim()),
                    P("@employee_user_id", ComboValue(cmbEmployee)),
                    P("@result", txtResult.Text.Trim()));
            }
            else
            {
                Database.Execute(@"
UPDATE service_history
SET equipment_id = @equipment_id, request_id = @request_id, service_date = @service_date, action = @action, employee_user_id = @employee_user_id, result = @result
WHERE id = @id",
                    P("@equipment_id", ComboValue(cmbEquipment)),
                    P("@request_id", ComboValue(cmbRequest)),
                    P("@service_date", DateTimeSql(dtpService.Value)),
                    P("@action", txtAction.Text.Trim()),
                    P("@employee_user_id", ComboValue(cmbEmployee)),
                    P("@result", txtResult.Text.Trim()),
                    P("@id", id));
            }
            LoadHistory();
        }

        private void btnAddUser_Click(object sender, EventArgs e)
        {
            ShowUserDialog(0);
        }

        private void btnEditUser_Click(object sender, EventArgs e)
        {
            ShowUserDialog(SelectedId(dgvUsers));
        }

        private void dgvUsers_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && btnEditUser.Enabled)
            {
                ShowUserDialog(SelectedId(dgvUsers));
            }
        }

        private void btnDeleteUser_Click(object sender, EventArgs e)
        {
            int id = SelectedId(dgvUsers);
            if (id == 0)
            {
                return;
            }
            if (id == currentUserId)
            {
                MessageBox.Show("Нельзя отключить текущего пользователя");
                return;
            }
            if (MessageBox.Show("Отключить выбранного пользователя?", "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                return;
            }
            Database.Execute("UPDATE users SET is_active = 0 WHERE id = @id", P("@id", id));
            LoadUsers();
        }

        private void ShowUserDialog(int id)
        {
            DataRow row = null;
            if (id > 0)
            {
                DataTable data = Database.Query("SELECT * FROM users WHERE id = @id", P("@id", id));
                if (data.Rows.Count == 0)
                {
                    return;
                }
                row = data.Rows[0];
            }

            Form form = CreateEditForm(id == 0 ? "Добавление пользователя" : "Редактирование пользователя", 390);
            TableLayoutPanel table = CreateTable(form, 7);
            TextBox txtUserLogin = CreateText(row == null ? "" : Convert.ToString(row["login"]));
            TextBox txtFullName = CreateText(row == null ? "" : Convert.ToString(row["full_name"]));
            TextBox txtUserPassword = CreateText("");
            txtUserPassword.PasswordChar = '*';
            ComboBox cmbRole = CreateCombo(Lookup("SELECT id, name FROM roles ORDER BY id"));
            CheckBox chkActive = new CheckBox();
            chkActive.Text = "Активен";
            chkActive.Checked = row == null || Convert.ToInt32(row["is_active"]) == 1;
            Label passwordInfo = new Label();
            passwordInfo.Text = id == 0 ? "Пароль обязателен" : "Оставьте пустым, если пароль менять не нужно";
            passwordInfo.Dock = DockStyle.Fill;
            passwordInfo.TextAlign = ContentAlignment.MiddleLeft;

            AddRow(table, 0, "Логин", txtUserLogin);
            AddRow(table, 1, "ФИО", txtFullName);
            AddRow(table, 2, "Пароль", txtUserPassword);
            table.Controls.Add(passwordInfo, 1, 3);
            AddRow(table, 4, "Роль", cmbRole);
            AddRow(table, 5, "Состояние", chkActive);
            AddButtons(table, 6, form);

            if (row != null)
            {
                SelectCombo(cmbRole, row["role_id"]);
            }

            if (form.ShowDialog(this) != DialogResult.OK)
            {
                return;
            }
            if (Empty(txtUserLogin.Text) || Empty(txtFullName.Text) || (id == 0 && Empty(txtUserPassword.Text)))
            {
                MessageBox.Show("Заполните логин, ФИО и пароль");
                return;
            }

            if (id == 0)
            {
                Database.Execute(@"
INSERT INTO users(login, password_hash, full_name, role_id, is_active)
VALUES(@login, @password_hash, @full_name, @role_id, @is_active)",
                    P("@login", txtUserLogin.Text.Trim()),
                    P("@password_hash", HashPassword(txtUserPassword.Text)),
                    P("@full_name", txtFullName.Text.Trim()),
                    P("@role_id", ComboValue(cmbRole)),
                    P("@is_active", chkActive.Checked ? 1 : 0));
            }
            else
            {
                if (Empty(txtUserPassword.Text))
                {
                    Database.Execute(@"
UPDATE users
SET login = @login, full_name = @full_name, role_id = @role_id, is_active = @is_active
WHERE id = @id",
                        P("@login", txtUserLogin.Text.Trim()),
                        P("@full_name", txtFullName.Text.Trim()),
                        P("@role_id", ComboValue(cmbRole)),
                        P("@is_active", chkActive.Checked ? 1 : 0),
                        P("@id", id));
                }
                else
                {
                    Database.Execute(@"
UPDATE users
SET login = @login, password_hash = @password_hash, full_name = @full_name, role_id = @role_id, is_active = @is_active
WHERE id = @id",
                        P("@login", txtUserLogin.Text.Trim()),
                        P("@password_hash", HashPassword(txtUserPassword.Text)),
                        P("@full_name", txtFullName.Text.Trim()),
                        P("@role_id", ComboValue(cmbRole)),
                        P("@is_active", chkActive.Checked ? 1 : 0),
                        P("@id", id));
                }
            }
            LoadUsers();
        }

        private void btnFailuresReport_Click(object sender, EventArgs e)
        {
            LoadFailuresReport();
        }

        private void btnEquipmentReport_Click(object sender, EventArgs e)
        {
            LoadEquipmentReport();
        }

        private void btnCurrentRequestsReport_Click(object sender, EventArgs e)
        {
            LoadCurrentRequestsReport();
        }

        private void LoadFailuresReport()
        {
            currentReportTitle = "Отчёт по неисправностям за период";
            DataTable table = Database.Query(@"
SELECT f.detected_date AS 'Дата обнаружения',
       e.inventory_number AS 'Инв. номер',
       e.name AS 'Оборудование',
       ft.name AS 'Тип неисправности',
       f.criticality AS 'Критичность',
       CASE WHEN f.is_closed = 1 THEN 'Закрыта' ELSE 'Открыта' END AS 'Состояние неисправности',
       IFNULL(rs.name, '') AS 'Статус заявки',
       IFNULL(u.full_name, '') AS 'Ответственный',
       f.description AS 'Описание'
FROM failures f
INNER JOIN equipment e ON e.id = f.equipment_id
INNER JOIN failure_types ft ON ft.id = f.failure_type_id
LEFT JOIN requests r ON r.failure_id = f.id
LEFT JOIN request_statuses rs ON rs.id = r.status_id
LEFT JOIN users u ON u.id = r.responsible_user_id
WHERE date(f.detected_date) BETWEEN date(@date_from) AND date(@date_to)
ORDER BY f.detected_date DESC", P("@date_from", DateSql(dtpReportFrom.Value)), P("@date_to", DateSql(dtpReportTo.Value)));
            dgvReports.DataSource = table;
        }

        private void LoadEquipmentReport()
        {
            currentReportTitle = "Отчёт по списку оборудования";
            DataTable table = Database.Query(@"
SELECT e.inventory_number AS 'Инв. номер',
       e.name AS 'Наименование',
       t.name AS 'Тип',
       e.model AS 'Модель',
       e.serial_number AS 'Серийный номер',
       IFNULL(c.number || ' / ' || c.place, '') AS 'Компьютер',
       e.location AS 'Размещение',
       s.name AS 'Состояние',
       e.installation_date AS 'Дата установки',
       e.cost AS 'Стоимость',
       e.responsible_person AS 'Ответственный'
FROM equipment e
INNER JOIN equipment_types t ON t.id = e.type_id
INNER JOIN equipment_statuses s ON s.id = e.status_id
LEFT JOIN computers c ON c.id = e.computer_id
ORDER BY t.name, e.inventory_number");
            dgvReports.DataSource = table;
        }

        private void LoadCurrentRequestsReport()
        {
            currentReportTitle = "Отчёт по текущим заявкам";
            DataTable table = Database.Query(@"
SELECT r.created_date AS 'Дата создания',
       e.inventory_number AS 'Инв. номер',
       e.name AS 'Оборудование',
       rs.name AS 'Статус',
       IFNULL(u.full_name, '') AS 'Ответственный',
       f.description AS 'Неисправность',
       r.work_description AS 'Выполненные работы'
FROM requests r
INNER JOIN failures f ON f.id = r.failure_id
INNER JOIN equipment e ON e.id = r.equipment_id
INNER JOIN request_statuses rs ON rs.id = r.status_id
LEFT JOIN users u ON u.id = r.responsible_user_id
WHERE r.status_id NOT IN (4, 5)
ORDER BY r.created_date DESC");
            dgvReports.DataSource = table;
        }

        private void btnExportReport_Click(object sender, EventArgs e)
        {
            if (dgvReports.Rows.Count == 0)
            {
                MessageBox.Show("Нет данных для экспорта");
                return;
            }

            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "Excel (*.xls)|*.xls|CSV (*.csv)|*.csv";
            dialog.FileName = currentReportTitle.Replace(' ', '_') + ".xls";
            if (dialog.ShowDialog(this) != DialogResult.OK)
            {
                return;
            }

            if (Path.GetExtension(dialog.FileName).ToLower() == ".csv")
            {
                WriteCsv(dialog.FileName, dgvReports);
            }
            else
            {
                WriteHtmlExcel(dialog.FileName, dgvReports);
            }
            MessageBox.Show("Отчёт сохранён");
        }

        private void WriteCsv(string path, DataGridView grid)
        {
            StringBuilder builder = new StringBuilder();
            List<string> headers = new List<string>();
            foreach (DataGridViewColumn column in grid.Columns)
            {
                if (column.Visible)
                {
                    headers.Add(EscapeCsv(column.HeaderText));
                }
            }
            builder.AppendLine(string.Join(";", headers.ToArray()));
            foreach (DataGridViewRow row in grid.Rows)
            {
                if (row.IsNewRow)
                {
                    continue;
                }
                List<string> values = new List<string>();
                foreach (DataGridViewColumn column in grid.Columns)
                {
                    if (column.Visible)
                    {
                        values.Add(EscapeCsv(Convert.ToString(row.Cells[column.Index].Value)));
                    }
                }
                builder.AppendLine(string.Join(";", values.ToArray()));
            }
            File.WriteAllText(path, builder.ToString(), Encoding.UTF8);
        }

        private string EscapeCsv(string value)
        {
            if (value == null)
            {
                value = "";
            }
            value = value.Replace("\r", " ").Replace("\n", " ").Replace("\"", "\"\"");
            return "\"" + value + "\"";
        }

        private void WriteHtmlExcel(string path, DataGridView grid)
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("<html><head><meta charset='utf-8'></head><body>");
            builder.AppendLine("<h3>" + Html(currentReportTitle) + "</h3>");
            builder.AppendLine("<table border='1'>");
            builder.AppendLine("<tr>");
            foreach (DataGridViewColumn column in grid.Columns)
            {
                if (column.Visible)
                {
                    builder.AppendLine("<th>" + Html(column.HeaderText) + "</th>");
                }
            }
            builder.AppendLine("</tr>");
            foreach (DataGridViewRow row in grid.Rows)
            {
                if (row.IsNewRow)
                {
                    continue;
                }
                builder.AppendLine("<tr>");
                foreach (DataGridViewColumn column in grid.Columns)
                {
                    if (column.Visible)
                    {
                        builder.AppendLine("<td>" + Html(Convert.ToString(row.Cells[column.Index].Value)) + "</td>");
                    }
                }
                builder.AppendLine("</tr>");
            }
            builder.AppendLine("</table></body></html>");
            File.WriteAllText(path, builder.ToString(), Encoding.UTF8);
        }

        private string Html(string value)
        {
            if (value == null)
            {
                return "";
            }
            return value.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
        }

        private void btnPrintReport_Click(object sender, EventArgs e)
        {
            if (dgvReports.Rows.Count == 0)
            {
                MessageBox.Show("Нет данных для печати");
                return;
            }

            printText = BuildPrintText(dgvReports);
            printPosition = 0;

            PrintDocument document = new PrintDocument();
            document.DefaultPageSettings.Landscape = true;
            document.DefaultPageSettings.Margins = new Margins(40, 40, 40, 40);
            document.PrintPage += PrintDocument_PrintPage;

            PrintPreviewDialog preview = new PrintPreviewDialog();
            preview.Document = document;
            preview.Width = 1000;
            preview.Height = 700;
            preview.ShowDialog(this);
        }
        private string BuildPrintText(DataGridView grid)
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine(currentReportTitle);
            builder.AppendLine("Дата формирования: " + DateTime.Now.ToString("dd.MM.yyyy HH:mm"));
            builder.AppendLine(new string('-', 120));
            foreach (DataGridViewColumn column in grid.Columns)
            {
                if (column.Visible)
                {
                    builder.Append(column.HeaderText);
                    builder.Append(" | ");
                }
            }
            builder.AppendLine();
            builder.AppendLine(new string('-', 120));
            foreach (DataGridViewRow row in grid.Rows)
            {
                if (row.IsNewRow)
                {
                    continue;
                }
                foreach (DataGridViewColumn column in grid.Columns)
                {
                    if (column.Visible)
                    {
                        string value = Convert.ToString(row.Cells[column.Index].Value);
                        value = value.Replace("\r", " ").Replace("\n", " ");
                        if (value.Length > 45)
                        {
                            value = value.Substring(0, 45);
                        }
                        builder.Append(value);
                        builder.Append(" | ");
                    }
                }
                builder.AppendLine();
            }
            return builder.ToString();
        }

        private void PrintDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            Font font = new Font("Arial", 9);
            Brush brush = Brushes.Black;

            float x = e.MarginBounds.Left;
            float y = e.MarginBounds.Top;
            float width = e.MarginBounds.Width;
            float lineHeight = font.GetHeight(e.Graphics) + 4;

            string[] lines = printText.Replace("\r\n", "\n").Split('\n');

            while (printPosition < lines.Length)
            {
                string line = lines[printPosition];

                List<string> wrappedLines = WrapText(e.Graphics, line, font, width);

                foreach (string wrappedLine in wrappedLines)
                {
                    if (y + lineHeight > e.MarginBounds.Bottom)
                    {
                        e.HasMorePages = true;
                        return;
                    }

                    e.Graphics.DrawString(wrappedLine, font, brush, x, y);
                    y += lineHeight;
                }

                printPosition++;
            }

            e.HasMorePages = false;
            printPosition = 0;
        }
        private List<string> WrapText(Graphics graphics, string text, Font font, float maxWidth)
        {
            List<string> result = new List<string>();

            if (string.IsNullOrWhiteSpace(text))
            {
                result.Add("");
                return result;
            }

            string[] words = text.Split(' ');
            string currentLine = "";

            foreach (string word in words)
            {
                string testLine;

                if (currentLine.Length == 0)
                {
                    testLine = word;
                }
                else
                {
                    testLine = currentLine + " " + word;
                }

                SizeF size = graphics.MeasureString(testLine, font);

                if (size.Width <= maxWidth)
                {
                    currentLine = testLine;
                }
                else
                {
                    if (currentLine.Length > 0)
                    {
                        result.Add(currentLine);
                        currentLine = word;
                    }
                    else
                    {
                        result.Add(word);
                        currentLine = "";
                    }
                }
            }

            if (currentLine.Length > 0)
            {
                result.Add(currentLine);
            }

            return result;
        }
        private void RecalculateComputerStatuses()
        {
            DataTable computers = Database.Query("SELECT id FROM computers");
            foreach (DataRow row in computers.Rows)
            {
                int id = Convert.ToInt32(row["id"]);
                object broken = Database.Scalar("SELECT COUNT(*) FROM equipment WHERE computer_id = @id AND status_id = 3", P("@id", id));
                object issue = Database.Scalar("SELECT COUNT(*) FROM equipment WHERE computer_id = @id AND status_id = 2", P("@id", id));
                object failures = Database.Scalar("SELECT COUNT(*) FROM failures WHERE computer_id = @id AND is_closed = 0", P("@id", id));
                int status = 1;
                if (Convert.ToInt32(broken) > 0)
                {
                    status = 3;
                }
                else if (Convert.ToInt32(issue) > 0 || Convert.ToInt32(failures) > 0)
                {
                    status = 2;
                }
                Database.Execute("UPDATE computers SET status_id = @status WHERE id = @id", P("@status", status), P("@id", id));
            }
        }
    }
}
