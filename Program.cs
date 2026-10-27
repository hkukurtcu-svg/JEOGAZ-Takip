using System;
using System.IO;
using System.Linq;
using System.Drawing;
using System.Windows.Forms;
using ClosedXML.Excel;
using Microsoft.Win32;

internal static class Program
{
    [STAThread]
    static void Main()
    {
        ApplicationConfiguration.Initialize();
        string? excelPath = GetExcelPath();
        if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
        {
            MessageBox.Show("Excel takip dosyası seçilmedi. Program kapatılıyor.", "JEOGAZ Takip",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        using var context = new ReminderContext(excelPath);
        Application.Run(context);
    }

    static string? GetExcelPath()
    {
        string settings = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "JEOGAZTakip");
        Directory.CreateDirectory(settings);
        string config = Path.Combine(settings, "excelpath.txt");

        if (File.Exists(config))
        {
            string saved = File.ReadAllText(config).Trim();
            if (File.Exists(saved)) return saved;
        }

        using var dialog = new OpenFileDialog
        {
            Title = "JEOGAZ araç takip Excel dosyasını seçin",
            Filter = "Excel dosyaları (*.xlsx)|*.xlsx",
            CheckFileExists = true
        };
        if (dialog.ShowDialog() != DialogResult.OK) return null;
        File.WriteAllText(config, dialog.FileName);
        return dialog.FileName;
    }
}

internal sealed class ReminderContext : ApplicationContext
{
    private readonly NotifyIcon tray;
    private readonly System.Windows.Forms.Timer timer;
    private readonly string excelPath;
    private DateTime lastDailyNotice = DateTime.MinValue;

    public ReminderContext(string path)
    {
        excelPath = path;
        var menu = new ContextMenuStrip();
        menu.Items.Add("Şimdi kontrol et", null, (_, _) => CheckAndNotify(true));
        menu.Items.Add("Excel dosyasını aç", null, (_, _) => OpenExcel());
        menu.Items.Add("Çıkış", null, (_, _) => ExitThread());

        tray = new NotifyIcon
        {
            Icon = SystemIcons.Information,
            Text = "JEOGAZ Araç Takip",
            ContextMenuStrip = menu,
            Visible = true
        };
        tray.DoubleClick += (_, _) => CheckAndNotify(true);

        timer = new System.Windows.Forms.Timer { Interval = 60 * 60 * 1000 };
        timer.Tick += (_, _) => CheckAndNotify(false);
        timer.Start();

        CheckAndNotify(true);
        ShowBalloon("JEOGAZ Takip çalışıyor", "Araç muayene ve T9 tarihleri takip ediliyor.");
    }

    private void CheckAndNotify(bool manual)
    {
        try
        {
            var items = ReadAlerts();
            if (items.Count == 0)
            {
                if (manual) ShowBalloon("JEOGAZ Takip", "Önümüzdeki 30 gün içinde süresi dolacak kayıt bulunamadı.");
                return;
            }

            bool daily = DateTime.Now.Date != lastDailyNotice.Date;
            if (manual || daily)
            {
                string message = string.Join("\n", items.Take(8));
                if (items.Count > 8) message += $"\n... ve {items.Count - 8} kayıt daha.";
                ShowBalloon("JEOGAZ Araç Takip Uyarısı", message);
                lastDailyNotice = DateTime.Now;
            }
        }
        catch (Exception ex)
        {
            if (manual) MessageBox.Show("Excel dosyası okunamadı:\n" + ex.Message, "JEOGAZ Takip",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private System.Collections.Generic.List<string> ReadAlerts()
    {
        var alerts = new System.Collections.Generic.List<(int Days, string Text)>();
        DateTime today = DateTime.Today;

        using var workbook = new XLWorkbook(excelPath);
        var ws = workbook.Worksheet("Sayfa1");

        for (int r = 6; r <= ws.LastRowUsed().RowNumber(); r++)
        {
            string plate = ws.Cell(r, 3).GetString().Trim();
            if (string.IsNullOrWhiteSpace(plate)) continue;
            string brand = ws.Cell(r, 4).GetString().Trim();

            AddDateAlert(alerts, plate, brand, "Muayene", ws.Cell(r, 5), today, 20);
            AddDateAlert(alerts, plate, brand, "T9/Sigorta", ws.Cell(r, 8), today, 30);
        }

        return alerts.OrderBy(x => x.Days).Select(x => x.Text).ToList();
    }

    private static void AddDateAlert(
        System.Collections.Generic.List<(int Days, string Text)> alerts,
        string plate, string brand, string label, IXLCell cell, DateTime today, int threshold)
    {
        if (!cell.TryGetValue<DateTime>(out DateTime expiry) || expiry.Year < 2000) return;
        int days = (expiry.Date - today).Days;
        if (days < 0)
            alerts.Add((days, $"SÜRESİ GEÇTİ: {plate} ({brand}) - {label}: {Math.Abs(days)} gün önce"));
        else if (days == 0)
            alerts.Add((0, $"BUGÜN BİTİYOR: {plate} ({brand}) - {label}"));
        else if (days <= threshold)
            alerts.Add((days, $"Yaklaşıyor: {plate} ({brand}) - {label}: {days} gün kaldı"));
    }

    private void OpenExcel()
    {
        try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(excelPath) { UseShellExecute = true }); }
        catch (Exception ex) { MessageBox.Show(ex.Message, "JEOGAZ Takip"); }
    }

    private void ShowBalloon(string title, string message)
    {
        tray.BalloonTipTitle = title;
        tray.BalloonTipText = message.Length > 900 ? message[..900] : message;
        tray.BalloonTipIcon = ToolTipIcon.Warning;
        tray.ShowBalloonTip(10000);
    }

    protected override void ExitThreadCore()
    {
        timer.Stop();
        timer.Dispose();
        tray.Visible = false;
        tray.Dispose();
        base.ExitThreadCore();
    }
}
