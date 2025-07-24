//******************************************************************************
// $Author: Kevin Fortune
// $Date: 2025/05/23 
// $Name:  MailFlow Utility$
// $Revision: 1.0 - Initial Conception. Exporting of select tickets to EML files to be imported to an Email Client
// $Revision: 1.5 - Renamed to Super Utility, allow for soft delete of tickets
// $Revision: 2.0 - Allow for hard delete of tickets
// $Revision: 3.0 - Allow for automation via command-line arguments
// $Revision: 3.5 - Add feature to count enabled agents
// $Revision: 3.6 - Add MailFlow Totals: Agent count, and Ticket counts by status
// $Revision: 4.0 - Add registration key verification
//******************************************************************************

using System;
using System.IO;
using System.Data.SqlClient;
using System.Net.Mail;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Net.Mime;

namespace MailFlowUtility
{
    class Program
    {
        static void Main(string[] args)
        {
            string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MailFlowUtility.log");
            void Log(string message)
            {
                string entry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";
                Console.WriteLine(entry);
                File.AppendAllText(logPath, entry + Environment.NewLine);
            }

            Log("=== Starting MailFlow Utility ===\n");

            if (!VerifyRegistrationKey())
            {
                Log("Registration failed. Please enter a valid registration key.");
                Console.WriteLine("Invalid registration key. Application will now exit.");
                return;
            }

            string connectionString = "";
            SqlConnection conn = null;
            bool connected = false;

            string server = null;
            string database = null;
            string authMode = null;
            string user = null;
            string pass = null;

            if (args.Length > 0)
            {
                foreach (var arg in args)
                {
                    var parts = arg.Split('=');
                    if (parts.Length == 2)
                    {
                        string key = parts[0].ToLower();
                        string value = parts[1];
                        if (key == "--server") server = value;
                        else if (key == "--db") database = value;
                        else if (key == "--auth") authMode = value;
                        else if (key == "--user") user = value;
                        else if (key == "--pass") pass = value;
                    }
                }
            }

            while (!connected)
            {
                try
                {
                    if (string.IsNullOrEmpty(server))
                    {
                        Console.Write("Enter SQL Server name (e.g., localhost\\SQLEXPRESS): ");
                        server = Console.ReadLine();
                    }
                    if (string.IsNullOrEmpty(database))
                    {
                        Console.Write("Enter Database name: ");
                        database = Console.ReadLine();
                    }
                    if (string.IsNullOrEmpty(authMode))
                    {
                        Console.WriteLine("Select Authentication Mode:");
                        Console.WriteLine("1 - Windows Integrated Security");
                        Console.WriteLine("2 - SQL Server Authentication");
                        Console.Write("Choice: ");
                        authMode = Console.ReadLine();
                    }

                    if (authMode == "1" || authMode.ToLower() == "windows")
                    {
                        connectionString = $"Server={server};Database={database};Integrated Security=True;";
                    }
                    else if (authMode == "2" || authMode.ToLower() == "sql")
                    {
                        if (string.IsNullOrEmpty(user))
                        {
                            Console.Write("SQL Username: ");
                            user = Console.ReadLine();
                        }
                        if (string.IsNullOrEmpty(pass))
                        {
                            Console.Write("SQL Password: ");
                            pass = ReadPassword();
                        }
                        connectionString = $"Server={server};Database={database};User ID={user};Password={pass};";
                    }
                    else
                    {
                        Log("Invalid authentication option. Exiting.");
                        return;
                    }

                    connectionString += ";MultipleActiveResultSets=True";
                    conn = new SqlConnection(connectionString);
                    Log("Connecting to database...");
                    conn.Open();
                    connected = true;
                    Log("Successfully logged in.");
                }
                catch (Exception ex)
                {
                    Log("Connection failed: " + ex.Message);
                    if (args.Length > 0)
                    {
                        Log("Exiting due to command-line mode failure.");
                        return;
                    }
                    Console.WriteLine("Would you like to retry? (y/n): ");
                    if (Console.ReadLine().Trim().ToLower() != "y") return;
                }
            }

            while (true)
            {
                Console.WriteLine("\nMailFlow Utility Menu:");
                Console.WriteLine("1 - MailFlow Totals (Agent and Ticket counts)");
                Console.WriteLine("2 - Export Messages to EML");
                Console.WriteLine("3 - Soft Delete Tickets");
                Console.WriteLine("4 - Hard Delete Tickets");
                Console.WriteLine("0 - Exit");
                Console.Write("Select an option: ");
                var input = Console.ReadLine();

                switch (input)
                {
                    case "0":
                        return;
                    case "1":
                        ExecuteMailFlowTotals(conn, Log);
                        break;
                    case "2":
                        Console.WriteLine("Export to EML not yet implemented in this snippet.");
                        break;
                    case "3":
                        Console.WriteLine("Soft delete not yet implemented in this snippet.");
                        break;
                    case "4":
                        Console.WriteLine("Hard delete not yet implemented in this snippet.");
                        break;
                    default:
                        Console.WriteLine("Invalid option.");
                        break;
                }
            }
        }

        private static bool VerifyRegistrationKey()
        {
            const string storedHash = "A25F7F1D8846C99F4C2B89A5DC95B64D6C9824E8161F9FCEFF69DBF8619D2E40"; // SHA256 hash of valid key
            const string salt = "MySecretSalt";

            Console.Write("Enter registration key: ");
            string inputKey = Console.ReadLine().Trim();

            using (SHA256 sha = SHA256.Create())
            {
                byte[] hashBytes = sha.ComputeHash(Encoding.UTF8.GetBytes(inputKey + salt));
                string inputHash = BitConverter.ToString(hashBytes).Replace("-", "");
                return inputHash.Equals(storedHash, StringComparison.OrdinalIgnoreCase);
            }
        }

        private static void ExecuteMailFlowTotals(SqlConnection conn, Action<string> Log)
        {
            try
            {
                Log("=== MailFlow Totals ===");

                Log("Counting enabled agents...");
                using (SqlCommand cmd = new SqlCommand("SELECT COUNT(*) FROM Agents WHERE IsEnabled = 1", conn))
                {
                    int count = (int)cmd.ExecuteScalar();
                    Log("Total enabled agents: " + count);
                }

                Log("Counting tickets by status...");
                string query = @"SELECT TicketStateID, COUNT(*) AS Count FROM Tickets WHERE TicketStateID IN (1, 2, 3, 6) GROUP BY TicketStateID ORDER BY TicketStateID";
                using (SqlCommand cmd = new SqlCommand(query, conn))
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        int state = reader.GetInt32(0);
                        int count = reader.GetInt32(1);
                        string stateName = "";
                        switch (state)
                        {
                            case 1:
                                stateName = "Closed";
                                break;
                            case 2:
                                stateName = "Open";
                                break;
                            case 3:
                                stateName = "On-Hold";
                                break;
                            case 6:
                                stateName = "Marked for Deletion";
                                break;
                            default:
                                stateName = "Unknown State (" + state + ")";
                                break;
                        }

                        Log("Tickets (" + stateName + "): " + count);
                    }
                }
            }
            catch (Exception ex)
            {
                Log("Error while retrieving MailFlow totals: " + ex.Message);
            }
        }

        private static string ReadPassword()
        {
            string pass = "";
            ConsoleKeyInfo key;
            do
            {
                key = Console.ReadKey(true);
                if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
                {
                    pass += key.KeyChar;
                    Console.Write("*");
                }
                else if (key.Key == ConsoleKey.Backspace && pass.Length > 0)
                {
                    pass = pass.Substring(0, pass.Length - 1);
                    Console.Write("\b \b");
                }
            } while (key.Key != ConsoleKey.Enter);
            Console.WriteLine();
            return pass;
        }
    }
}
