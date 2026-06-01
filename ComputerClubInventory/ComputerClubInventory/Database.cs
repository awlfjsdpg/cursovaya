using System;
using System.Data;
using System.Data.SQLite;
using System.IO;

namespace ComputerClubInventory
{
    public static class Database
    {
        public static readonly string DbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "club_inventory.db");
        public static readonly string ConnectionString = "Data Source=" + DbPath + ";Version=3;Foreign Keys=True;";

        public static void Initialize()
        {
            bool needCreate = !File.Exists(DbPath);
            if (needCreate)
            {
                SQLiteConnection.CreateFile(DbPath);
            }

            using (SQLiteConnection connection = OpenConnection())
            using (SQLiteCommand command = connection.CreateCommand())
            {
                command.CommandText = SchemaSql;
                command.ExecuteNonQuery();
            }
        }

        public static SQLiteConnection OpenConnection()
        {
            SQLiteConnection connection = new SQLiteConnection(ConnectionString);
            connection.Open();
            using (SQLiteCommand command = connection.CreateCommand())
            {
                command.CommandText = "PRAGMA foreign_keys = ON;";
                command.ExecuteNonQuery();
            }
            return connection;
        }

        public static DataTable Query(string sql, params SQLiteParameter[] parameters)
        {
            using (SQLiteConnection connection = OpenConnection())
            using (SQLiteCommand command = new SQLiteCommand(sql, connection))
            {
                if (parameters != null && parameters.Length > 0)
                {
                    command.Parameters.AddRange(parameters);
                }
                using (SQLiteDataAdapter adapter = new SQLiteDataAdapter(command))
                {
                    DataTable table = new DataTable();
                    adapter.Fill(table);
                    return table;
                }
            }
        }

        public static int Execute(string sql, params SQLiteParameter[] parameters)
        {
            using (SQLiteConnection connection = OpenConnection())
            using (SQLiteCommand command = new SQLiteCommand(sql, connection))
            {
                if (parameters != null && parameters.Length > 0)
                {
                    command.Parameters.AddRange(parameters);
                }
                return command.ExecuteNonQuery();
            }
        }

        public static object Scalar(string sql, params SQLiteParameter[] parameters)
        {
            using (SQLiteConnection connection = OpenConnection())
            using (SQLiteCommand command = new SQLiteCommand(sql, connection))
            {
                if (parameters != null && parameters.Length > 0)
                {
                    command.Parameters.AddRange(parameters);
                }
                return command.ExecuteScalar();
            }
        }

        public static string SchemaSql = @"
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS roles (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    login TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    full_name TEXT NOT NULL,
    role_id INTEGER NOT NULL,
    is_active INTEGER NOT NULL DEFAULT 1,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (role_id) REFERENCES roles(id)
);

CREATE TABLE IF NOT EXISTS computer_statuses (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE,
    color TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS computers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    number TEXT NOT NULL UNIQUE,
    place TEXT NOT NULL,
    specifications TEXT NOT NULL,
    status_id INTEGER NOT NULL,
    installation_date TEXT NOT NULL,
    note TEXT,
    FOREIGN KEY (status_id) REFERENCES computer_statuses(id)
);

CREATE TABLE IF NOT EXISTS equipment_types (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS equipment_statuses (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE,
    color TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS equipment (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    inventory_number TEXT NOT NULL UNIQUE,
    name TEXT NOT NULL,
    model TEXT NOT NULL,
    serial_number TEXT NOT NULL UNIQUE,
    type_id INTEGER NOT NULL,
    computer_id INTEGER,
    purchase_date TEXT,
    installation_date TEXT NOT NULL,
    cost REAL NOT NULL DEFAULT 0,
    responsible_person TEXT NOT NULL,
    location TEXT NOT NULL,
    specifications TEXT NOT NULL,
    status_id INTEGER NOT NULL,
    note TEXT,
    FOREIGN KEY (type_id) REFERENCES equipment_types(id),
    FOREIGN KEY (computer_id) REFERENCES computers(id) ON DELETE SET NULL,
    FOREIGN KEY (status_id) REFERENCES equipment_statuses(id)
);

CREATE TABLE IF NOT EXISTS failure_types (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS failures (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    equipment_id INTEGER NOT NULL,
    computer_id INTEGER,
    failure_type_id INTEGER NOT NULL,
    detected_date TEXT NOT NULL,
    description TEXT NOT NULL,
    criticality TEXT NOT NULL,
    detected_by_user_id INTEGER NOT NULL,
    source TEXT NOT NULL,
    is_closed INTEGER NOT NULL DEFAULT 0,
    FOREIGN KEY (equipment_id) REFERENCES equipment(id) ON DELETE CASCADE,
    FOREIGN KEY (computer_id) REFERENCES computers(id) ON DELETE SET NULL,
    FOREIGN KEY (failure_type_id) REFERENCES failure_types(id),
    FOREIGN KEY (detected_by_user_id) REFERENCES users(id)
);

CREATE TABLE IF NOT EXISTS request_statuses (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS requests (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    failure_id INTEGER NOT NULL,
    equipment_id INTEGER NOT NULL,
    created_date TEXT NOT NULL,
    status_id INTEGER NOT NULL,
    responsible_user_id INTEGER,
    work_description TEXT NOT NULL DEFAULT '',
    result TEXT NOT NULL DEFAULT '',
    closed_date TEXT,
    FOREIGN KEY (failure_id) REFERENCES failures(id) ON DELETE CASCADE,
    FOREIGN KEY (equipment_id) REFERENCES equipment(id) ON DELETE CASCADE,
    FOREIGN KEY (status_id) REFERENCES request_statuses(id),
    FOREIGN KEY (responsible_user_id) REFERENCES users(id) ON DELETE SET NULL
);

CREATE TABLE IF NOT EXISTS service_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    equipment_id INTEGER NOT NULL,
    request_id INTEGER,
    service_date TEXT NOT NULL,
    action TEXT NOT NULL,
    employee_user_id INTEGER,
    result TEXT NOT NULL,
    FOREIGN KEY (equipment_id) REFERENCES equipment(id) ON DELETE CASCADE,
    FOREIGN KEY (request_id) REFERENCES requests(id) ON DELETE SET NULL,
    FOREIGN KEY (employee_user_id) REFERENCES users(id) ON DELETE SET NULL
);

CREATE INDEX IF NOT EXISTS idx_computers_number ON computers(number);
CREATE INDEX IF NOT EXISTS idx_equipment_inventory ON equipment(inventory_number);
CREATE INDEX IF NOT EXISTS idx_equipment_type ON equipment(type_id);
CREATE INDEX IF NOT EXISTS idx_equipment_status ON equipment(status_id);
CREATE INDEX IF NOT EXISTS idx_failures_date ON failures(detected_date);
CREATE INDEX IF NOT EXISTS idx_failures_equipment ON failures(equipment_id);
CREATE INDEX IF NOT EXISTS idx_requests_status ON requests(status_id);
CREATE INDEX IF NOT EXISTS idx_requests_date ON requests(created_date);
CREATE INDEX IF NOT EXISTS idx_history_equipment ON service_history(equipment_id);

INSERT OR IGNORE INTO roles(id, name) VALUES
(1, 'Руководитель'),
(2, 'Администратор'),
(3, 'Технический специалист');

INSERT OR IGNORE INTO users(id, login, password_hash, full_name, role_id, is_active) VALUES
(1, 'owner', '43a0d17178a9d26c9e0fe9a74b0b45e38d32f27aed887a008a54bf6e033bf7b9', 'Ответственный сотрудник', 1, 1),
(2, 'admin', '240be518fabd2724ddb6f04eeb1da5967448d7e831c08c8fa822809f74c720a9', 'Администратор клуба', 2, 1),
(3, 'tech', '3ac40463b419a7de590185c7121f0bfbe411d6168699e8014f521b050b1d6653', 'Технический специалист', 3, 1);

INSERT OR IGNORE INTO computer_statuses(id, name, color) VALUES
(1, 'Исправен', '#C8E6C9'),
(2, 'Есть неисправность', '#FFF9C4'),
(3, 'Неработоспособен', '#FFCDD2');

INSERT OR IGNORE INTO equipment_statuses(id, name, color) VALUES
(1, 'Исправен', '#C8E6C9'),
(2, 'Есть неисправность', '#FFF9C4'),
(3, 'Неработоспособен', '#FFCDD2');

INSERT OR IGNORE INTO request_statuses(id, name) VALUES
(1, 'Создана'),
(2, 'В работе'),
(3, 'Ожидает комплектующие'),
(4, 'Завершена'),
(5, 'Отменена');

INSERT OR IGNORE INTO equipment_types(id, name) VALUES
(1, 'Системный блок'),
(2, 'Монитор'),
(3, 'Клавиатура'),
(4, 'Мышь'),
(5, 'Гарнитура'),
(6, 'Сетевое оборудование'),
(7, 'ИБП'),
(8, 'Периферийное устройство');

INSERT OR IGNORE INTO failure_types(id, name) VALUES
(1, 'Перегрев оборудования'),
(2, 'Выход из строя периферии'),
(3, 'Повреждение кабеля'),
(4, 'Неисправность видеокарты'),
(5, 'Сбой накопителя'),
(6, 'Проблема системы охлаждения'),
(7, 'Программный сбой'),
(8, 'Другое');
";
    }
}
