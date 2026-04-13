using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.IO;
using Group4337.Новая_папка;

namespace Group4337.Новая_папка1
{
    public static class DatabaseHelper
    {
        private static string DbPath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "lab3_db.sqlite");
        private static string ConnectionString => $"Data Source={DbPath};";

        public static void Initialize()
        {
            using var conn = new SqliteConnection(ConnectionString);
            conn.Open();
            string sql = @"CREATE TABLE IF NOT EXISTS Clients (
                             Id INTEGER PRIMARY KEY AUTOINCREMENT,
                             ClientCode TEXT NOT NULL,
                             FullName TEXT NOT NULL,
                             Email TEXT NOT NULL,
                             Street TEXT NOT NULL)";
            using var cmd = new SqliteCommand(sql, conn);
            cmd.ExecuteNonQuery();
        }

        public static void SaveClients(List<Client> clients)
        {
            using var conn = new SqliteConnection(ConnectionString);
            conn.Open();
            using var tran = conn.BeginTransaction();

            foreach (var c in clients)
            {
                string sql = "INSERT INTO Clients (ClientCode, FullName, Email, Street) VALUES (@cc, @fn, @em, @st)";

                using var cmd = new SqliteCommand(sql, conn, tran);

                cmd.Parameters.AddWithValue("@cc", c.ClientCode ?? "");
                cmd.Parameters.AddWithValue("@fn", c.FullName ?? "");
                cmd.Parameters.AddWithValue("@em", c.Email ?? "");
                cmd.Parameters.AddWithValue("@st", c.Street ?? "");

                cmd.ExecuteNonQuery();
            }

            tran.Commit();
        }

        public static List<Client> GetAllClients()
        {
            var result = new List<Client>();
            using var conn = new SqliteConnection(ConnectionString);
            conn.Open();
            using var cmd = new SqliteCommand("SELECT ClientCode, FullName, Email, Street FROM Clients", conn);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                result.Add(new Client
                {
                    ClientCode = reader.GetString(0),
                    FullName = reader.GetString(1),
                    Email = reader.GetString(2),
                    Street = reader.GetString(3)
                });
            }
            return result;
        }

        public static void ClearTable()
        {
            using var conn = new SqliteConnection(ConnectionString);
            conn.Open();
            using var cmd = new SqliteCommand("DELETE FROM Clients", conn);
            cmd.ExecuteNonQuery();
        }
    }
}