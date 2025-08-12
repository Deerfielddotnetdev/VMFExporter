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
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace MailFlowEmlExporter
{
    internal static class Program
    {
        /// <summary>
        /// Entry point. Shows the UnlockForm (unless --dev-bypass is supplied), then runs the exporter.
        /// - Presents the unlock UI for licensing/gating.
        /// - Parses CLI args and hands control to Exporter.Run().
        /// - Guards against unhandled exceptions and shows a friendly error dialog if possible.
        /// </summary>
        [STAThread]
        private static int Main(string[] args)
        {
            try
            {
                var devBypass = HasArg(args, "--dev-bypass");

                // --- Unlock gate ---
                // If not bypassed, display the UnlockForm as a modal dialog.
                // The exporter only runs when the dialog returns OK (user unlocked successfully).
                if (!devBypass)
                {
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);

                    using (var unlock = new UnlockForm())
                    {
                        var result = unlock.ShowDialog();
                        if (result != DialogResult.OK)
                        {
                            // User closed the form or did not unlock; exit gracefully.
                            return 1;
                        }
                    }
                }

                // --- Parse options + run ---
                // Converts CLI args into an ExportOptions object and starts the export.
                var options = Exporter.ParseArgs(args);
                return Exporter.Run(options);
            }
            catch (Exception ex)
            {
                // Final safety net. Try to show a dialog; if we’re headless (no UI), just return failure.
                try { MessageBox.Show(ex.ToString(), "MailFlow EML Exporter - Fatal Error", MessageBoxButtons.OK, MessageBoxIcon.Error); } catch { /* ignore */ }
                return 1;
            }
        }

        /// <summary>
        /// Convenience: check if a flag exists in the args (e.g., --dev-bypass).
        /// </summary>
        private static bool HasArg(string[] args, string key) =>
            Array.FindIndex(args, a => string.Equals(a, key, StringComparison.OrdinalIgnoreCase)) >= 0;
    }

    /// <summary>
    /// Strongly-typed container for CLI options (connection string, export directory, and optional date range).
    /// </summary>
    internal sealed class ExportOptions
    {
        public string ConnectionString { get; set; } =
            "Server=YOUR_SERVER;Database=YOUR_DB;Trusted_Connection=True;";

        public string ExportDir { get; set; } =
            Path.Combine(Environment.CurrentDirectory, "exports");

        public DateTime? From { get; set; }
        public DateTime? To { get; set; }
    }

    internal static class Exporter
    {
        /// <summary>
        /// Orchestrates the full export:
        /// - Ensures output folder exists.
        /// - Connects to SQL.
        /// - Iterates Inbound and Outbound tables with optional date filters.
        /// - For each message:
        ///   * Builds a MailMessage (addresses, subject, body, headers)
        ///   * Attaches files from the attachment tables
        ///   * Generates a safe, unique .eml filename
        ///   * Uses SmtpClient’s pickup-directory to emit a .eml, then moves it to the export folder
        /// - Logs progress and warnings; continues past per-message errors.
        /// </summary>
        public static int Run(ExportOptions options)
        {
            // --- Ensure output directory exists ---
            Directory.CreateDirectory(options.ExportDir);

            Console.WriteLine($"Export dir : {options.ExportDir}");
            Console.WriteLine($"Date range : {(options.From.HasValue ? options.From.Value.ToString("u") : "Any")} -> {(options.To.HasValue ? options.To.Value.ToString("u") : "Any")}");
            Console.WriteLine();

            // --- Open SQL connection ---
            using var conn = new SqlConnection(options.ConnectionString);
            conn.Open();

            // --- Process both directions with the same logic ---
            foreach (var meta in new[]
            {
                new { Direction = "Inbound",  MsgTable = "InboundMessages",  AttTable = "InboundMessageAttachments" },
                new { Direction = "Outbound", MsgTable = "OutboundMessages", AttTable = "OutboundMessageAttachments" }
            })
            {
                Console.WriteLine($"=== Processing {meta.Direction} ===");

                // --- Build WHERE clause for optional date filtering ---
                var where = new List<string>();
                if (options.From.HasValue) where.Add("EmailDateTime >= @From");
                if (options.To.HasValue)   where.Add("EmailDateTime <  @To");

                // --- Compose SQL for messages (ordered for deterministic output) ---
                var messageQuery = $@"
SELECT ID, EmailFrom, EmailPrimaryTo, EmailTo, EmailCc, EmailBcc, EmailDateTime, Subject, Body
FROM {meta.MsgTable}
{(where.Count > 0 ? "WHERE " + string.Join(" AND ", where) : "")}
ORDER BY EmailDateTime ASC, ID ASC";

                using var cmd = new SqlCommand(messageQuery, conn);
                if (options.From.HasValue) cmd.Parameters.AddWithValue("@From", options.From.Value);
                if (options.To.HasValue)   cmd.Parameters.AddWithValue("@To", options.To.Value);

                using var reader = cmd.ExecuteReader();
                var count = 0;

                // --- Iterate messages ---
                while (reader.Read())
                {
                    try
                    {
                        // --- Extract fields from the current row ---
                        var id      = reader.GetInt32(reader.GetOrdinal("ID"));
                        var from    = reader["EmailFrom"]?.ToString() ?? "";
                        var pto     = reader["EmailPrimaryTo"]?.ToString() ?? "";
                        var to      = reader["EmailTo"]?.ToString() ?? "";
                        var cc      = reader["EmailCc"]?.ToString() ?? "";
                        var bcc     = reader["EmailBcc"]?.ToString() ?? "";
                        var subject = reader["Subject"]?.ToString() ?? "";
                        var body    = reader["Body"]?.ToString() ?? "";
                        var date    = reader.GetDateTime(reader.GetOrdinal("EmailDateTime"));

                        // --- Build MailMessage with proper encodings and headers ---
                        using var msg = new MailMessage();

                        // From (tolerates "Name <addr@x.com>" formats or raw addresses)
                        if (!string.IsNullOrWhiteSpace(from))
                            msg.From = SafeAddress(from);

                        // To/CC/BCC
                        // Note: we prioritize EmailPrimaryTo then append EmailTo, and split on comma/semicolon.
                        AddAddresses(msg.To, pto);
                        AddAddresses(msg.To, to);
                        AddAddresses(msg.CC, cc);
                        AddAddresses(msg.Bcc, bcc);

                        // Subject + encodings
                        msg.Subject = subject ?? "";
                        msg.SubjectEncoding  = Encoding.UTF8;
                        msg.BodyEncoding     = Encoding.UTF8;
                        msg.HeadersEncoding  = Encoding.UTF8;

                        // Body (detect HTML vs. plain text via lightweight heuristic)
                        var isHtml = LooksLikeHtml(body);
                        msg.IsBodyHtml = isHtml;
                        msg.Body       = body ?? "";

                        // Force RFC 2822 Date header (UTC "r" format)
                        msg.Headers.Remove("Date");
                        msg.Headers.Add("Date", date.ToUniversalTime().ToString("r"));

                        // --- Load and attach files for this message ---
                        foreach (var att in GetAttachments(conn, meta.AttTable, id))
                        {
                            if (File.Exists(att.Path))
                            {
                                var attachment = new Attachment(att.Path, MediaTypeNames.Application.Octet)
                                {
                                    Name = att.FileName
                                };
                                msg.Attachments.Add(attachment);
                            }
                            else
                            {
                                // Missing files aren’t fatal; we log and continue.
                                Console.WriteLine($"[warn] Missing attachment file: {att.Path}");
                            }
                        }

                        // --- Build a safe, unique output filename ---
                        var safeBase = BuildSafeFileBase(id, subject, date, meta.Direction);
                        var emlPath  = UniquePath(Path.Combine(options.ExportDir, safeBase + ".eml"));

                        // --- Emit .eml via pickup directory (isolated per message to avoid races) ---
                        var pickup = Path.Combine(Path.GetTempPath(), "eml_pickup_" + Guid.NewGuid().ToString("N"));
                        Directory.CreateDirectory(pickup);

                        using (var client = new SmtpClient())
                        {
                            client.DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory;
                            client.PickupDirectoryLocation = pickup;
                            client.Send(msg);
                        }

                        // --- Move the generated .eml into the export directory ---
                        var generated = Directory.EnumerateFiles(pickup, "*.eml")
                                                 .OrderByDescending(File.GetCreationTimeUtc)
                                                 .FirstOrDefault();
                        if (generated == null)
                            throw new InvalidOperationException("No .eml generated by pickup directory.");

                        File.Move(generated, emlPath);
                        Directory.Delete(pickup, true);
                        count++;

                        // --- Periodic progress log ---
                        if (count % 50 == 0) Console.WriteLine($"... {count} messages exported");
                    }
                    catch (Exception exMsg)
                    {
                        // Per-message error: log and continue to next message.
                        Console.WriteLine($"[error] Message export failed: {exMsg.Message}");
                    }
                }

                // --- Direction summary ---
                Console.WriteLine($"Exported {count} {meta.Direction.ToLowerInvariant()} message(s).");
            }

            Console.WriteLine("\nExport complete.");
            return 0;
        }

        /// <summary>
        /// Parses CLI args into an ExportOptions instance.
        /// Supports:
        ///   --conn "SQL connection string"
        ///   --out  "output folder"
        ///   --from YYYY-MM-DD (inclusive)
        ///   --to   YYYY-MM-DD (exclusive)
        /// </summary>
        public static ExportOptions ParseArgs(string[] args)
        {
            var opt = new ExportOptions
            {
                ConnectionString = GetArg(args, "--conn") ?? "Server=YOUR_SERVER;Database=YOUR_DB;Trusted_Connection=True;",
                ExportDir        = GetArg(args, "--out")  ?? Path.Combine(Environment.CurrentDirectory, "exports"),
                From             = TryParseDate(GetArg(args, "--from")),
                To               = TryParseDate(GetArg(args, "--to"))
            };
            return opt;
        }

        /// <summary>
        /// Returns the value immediately following a flag (e.g., for "--out C:\path", returns "C:\path").
        /// Returns null if not present.
        /// </summary>
        private static string GetArg(string[] args, string key)
        {
            var i = Array.FindIndex(args, a => string.Equals(a, key, StringComparison.OrdinalIgnoreCase));
            if (i >= 0 && i + 1 < args.Length) return args[i + 1];
            return null;
        }

        /// <summary>
        /// Tries to parse a date string (flexible parser). Returns null on failure.
        /// </summary>
        private static DateTime? TryParseDate(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;
            return DateTime.TryParse(s, out var dt) ? dt : (DateTime?)null;
        }

        /// <summary>
        /// Builds a MailAddress from raw input.
        /// Accepts "Name &lt;email@x.com&gt;" or "email@x.com"; falls back to scraping angle brackets if needed.
        /// </summary>
        private static MailAddress SafeAddress(string raw)
        {
            try { return new MailAddress(raw.Trim()); }
            catch { return new MailAddress(ScrubAngleAddress(raw)); }
        }

        /// <summary>
        /// If the address is in "Name &lt;addr&gt;" format, returns only "addr"; otherwise returns the trimmed input.
        /// </summary>
        private static string ScrubAngleAddress(string raw)
        {
            var m = Regex.Match(raw, @"<([^>]+)>");
            return m.Success ? m.Groups[1].Value.Trim() : raw.Trim();
        }

        /// <summary>
        /// Adds multiple addresses into a MailAddressCollection.
        /// Splits on commas and semicolons; ignores malformed addresses instead of throwing.
        /// </summary>
        private static void AddAddresses(MailAddressCollection collection, string list)
        {
            if (string.IsNullOrWhiteSpace(list)) return;
            var parts = list.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var p in parts)
            {
                var s = p.Trim();
                if (s.Length == 0) continue;
                try { collection.Add(new MailAddress(s)); }
                catch { /* ignore malformed */ }
            }
        }

        /// <summary>
        /// Quick heuristic to guess whether a body is HTML.
        /// Looks for common HTML markers/tags.
        /// </summary>
        private static bool LooksLikeHtml(string body)
        {
            if (string.IsNullOrEmpty(body)) return false;
            var b = body.TrimStart();
            return b.StartsWith("<!DOCTYPE", StringComparison.OrdinalIgnoreCase)
                || b.StartsWith("<html", StringComparison.OrdinalIgnoreCase)
                || Regex.IsMatch(body, @"</\s*(html|body|p|div|span|table)\s*>", RegexOptions.IgnoreCase);
        }

        /// <summary>
        /// Builds a readable, mostly safe filename base:
        /// "[yyyyMMdd_HHmmss]_[Direction]_[ID]_[Subject...]"
        /// - Replaces invalid path characters with underscores
        /// - Collapses whitespace
        /// - Truncates to ~140 chars to avoid path issues
        /// </summary>
        private static string BuildSafeFileBase(int id, string subject, DateTime date, string direction)
        {
            var ts = date.ToString("yyyyMMdd_HHmmss");
            var baseName = $"{ts}_{direction}_{id}_{subject ?? ""}";

            var invalid = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
            var cleaned = new string(baseName.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray());
            cleaned = Regex.Replace(cleaned, @"\s+", " ").Trim();

            if (cleaned.Length > 140) cleaned = cleaned.Substring(0, 140).Trim();
            return cleaned.Length == 0 ? $"{ts}_{direction}_{id}" : cleaned;
        }

        /// <summary>
        /// If the target file exists, appends " (n)" before the extension until a free path is found.
        /// Throws if it can’t find a unique name after many attempts (practically unreachable).
        /// </summary>
        private static string UniquePath(string path)
        {
            if (!File.Exists(path)) return path;
            var dir  = Path.GetDirectoryName(path)!;
            var name = Path.GetFileNameWithoutExtension(path);
            var ext  = Path.GetExtension(path);
            for (int i = 1; i < 10_000; i++)
            {
                var candidate = Path.Combine(dir, $"{name} ({i}){ext}");
                if (!File.Exists(candidate)) return candidate;
            }
            throw new IOException("Could not create a unique filename.");
        }

        /// <summary>
        /// Enumerates (FileName, Path) for attachments linked to a given message ID.
        /// Joins the direction-specific link table (Inbound/OutboundMessageAttachments) to the shared Attachments table.
        /// </summary>
        private static IEnumerable<(string FileName, string Path)> GetAttachments(SqlConnection conn, string attachmentTable, int messageId)
        {
            const string sql = @"
SELECT A.FileName, A.AttachmentLocation
FROM {0} MA
JOIN Attachments A ON MA.AttachmentID = A.ID
WHERE MA.MessageID = @MsgID";

            using var cmd = new SqlCommand(string.Format(sql, attachmentTable), conn);
            cmd.Parameters.AddWithValue("@MsgID", messageId);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                yield return (
                    r["FileName"]?.ToString() ?? "attachment",
                    r["AttachmentLocation"]?.ToString() ?? ""
                );
            }
        }
    }
}
