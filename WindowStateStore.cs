namespace Tabel;

using System.Text.Json;

internal static class WindowStateStore
{
    private sealed record SavedWindow(int X, int Y, int Width, int Height, bool Maximized);

    private static readonly object SyncRoot = new();
    private static readonly string SettingsPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "TabelSkud",
        "window-state.json");

    public static void Attach(Form form, string key)
    {
        form.Load += (_, _) => Restore(form, key);
        form.FormClosing += (_, _) => Save(form, key);
    }

    private static void Restore(Form form, string key)
    {
        SavedWindow? saved;

        lock (SyncRoot)
        {
            Dictionary<string, SavedWindow> states = ReadStates();
            states.TryGetValue(key, out saved);
        }

        if (saved == null || saved.Width <= 0 || saved.Height <= 0)
            return;

        var requested = new Rectangle(saved.X, saved.Y, saved.Width, saved.Height);
        Rectangle area = Screen.FromRectangle(requested).WorkingArea;
        if (area.Width <= 0 || area.Height <= 0)
            return;

        int minimumWidth = Math.Min(Math.Max(240, form.MinimumSize.Width), area.Width);
        int minimumHeight = Math.Min(Math.Max(180, form.MinimumSize.Height), area.Height);
        int width = Math.Clamp(saved.Width, minimumWidth, area.Width);
        int height = Math.Clamp(saved.Height, minimumHeight, area.Height);
        int x = Math.Clamp(saved.X, area.Left, area.Right - width);
        int y = Math.Clamp(saved.Y, area.Top, area.Bottom - height);

        form.StartPosition = FormStartPosition.Manual;
        form.Bounds = new Rectangle(x, y, width, height);
        form.WindowState = saved.Maximized ? FormWindowState.Maximized : FormWindowState.Normal;
    }

    private static void Save(Form form, string key)
    {
        if (form.WindowState == FormWindowState.Minimized)
            return;

        Rectangle bounds = form.WindowState == FormWindowState.Normal ? form.Bounds : form.RestoreBounds;
        if (bounds.Width <= 0 || bounds.Height <= 0)
            return;

        var saved = new SavedWindow(
            bounds.X,
            bounds.Y,
            bounds.Width,
            bounds.Height,
            form.WindowState == FormWindowState.Maximized);

        lock (SyncRoot)
        {
            try
            {
                Dictionary<string, SavedWindow> states = ReadStates();
                states[key] = saved;
                string? directory = Path.GetDirectoryName(SettingsPath);
                if (!string.IsNullOrWhiteSpace(directory))
                    Directory.CreateDirectory(directory);

                File.WriteAllText(
                    SettingsPath,
                    JsonSerializer.Serialize(states, new JsonSerializerOptions { WriteIndented = true }));
            }
            catch
            {
                // Ошибка сохранения положения окна не должна мешать основной работе.
            }
        }
    }

    private static Dictionary<string, SavedWindow> ReadStates()
    {
        try
        {
            if (!File.Exists(SettingsPath))
                return new Dictionary<string, SavedWindow>(StringComparer.Ordinal);

            return JsonSerializer.Deserialize<Dictionary<string, SavedWindow>>(File.ReadAllText(SettingsPath))
                ?? new Dictionary<string, SavedWindow>(StringComparer.Ordinal);
        }
        catch
        {
            return new Dictionary<string, SavedWindow>(StringComparer.Ordinal);
        }
    }
}
