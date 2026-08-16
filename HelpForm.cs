namespace Tabel;

using System.Net;
using System.Text;
using System.Text.RegularExpressions;

internal sealed class HelpForm : Form
{
    public HelpForm()
    {
        Text = "Справка — Табель СКУД";
        StartPosition = FormStartPosition.CenterParent;
        MinimumSize = new Size(640, 480);
        Size = new Size(900, 700);
        KeyPreview = true;

        var tabs = new TabControl
        {
            Dock = DockStyle.Fill,
            Font = new Font("Segoe UI", 10.5F, FontStyle.Bold),
            Padding = new Point(18, 6)
        };

        tabs.TabPages.Add(CreateDocumentPage(
            "Справка",
            LoadDocumentHtml("Справка.md", "Табель СКУД", "Справка недоступна")));
        tabs.TabPages.Add(CreateDocumentPage(
            "Тех. задание",
            LoadDocumentHtml("ТЗ.md", "Техническое задание", "ТЗ недоступно")));

        Controls.Add(tabs);
        WindowStateStore.Attach(this, "Help");
        KeyDown += (_, e) =>
        {
            if (e.KeyCode == Keys.Escape)
                Close();
        };
    }

    private static TabPage CreateDocumentPage(string title, string html)
    {
        var page = new TabPage(title)
        {
            BackColor = Color.White,
            Padding = Padding.Empty
        };

        var browser = new WebBrowser
        {
            Dock = DockStyle.Fill,
            AllowWebBrowserDrop = false,
            IsWebBrowserContextMenuEnabled = true,
            ScriptErrorsSuppressed = true,
            WebBrowserShortcutsEnabled = true,
            DocumentText = html
        };

        page.Controls.Add(browser);
        return page;
    }

    private static string LoadDocumentHtml(string fileName, string navigationTitle, string errorTitle)
    {
        string path = Path.Combine(AppContext.BaseDirectory, "docs", fileName);

        try
        {
            string markdown = File.ReadAllText(path, Encoding.UTF8);
            return MarkdownToHtml(markdown, navigationTitle);
        }
        catch (Exception ex)
        {
            return WrapHtml(
                $"<h1>{WebUtility.HtmlEncode(errorTitle)}</h1><p>{WebUtility.HtmlEncode(ex.Message)}</p>",
                string.Empty,
                navigationTitle);
        }
    }

    private static string MarkdownToHtml(string markdown, string navigationTitle)
    {
        var body = new StringBuilder();
        var navigation = new StringBuilder();
        bool inUnorderedList = false;
        bool inOrderedList = false;
        bool inCodeBlock = false;
        int sectionNumber = 0;

        void CloseLists()
        {
            if (inUnorderedList)
            {
                body.AppendLine("</ul>");
                inUnorderedList = false;
            }

            if (inOrderedList)
            {
                body.AppendLine("</ol>");
                inOrderedList = false;
            }
        }

        foreach (string rawLine in markdown.Replace("\r\n", "\n").Split('\n'))
        {
            string line = rawLine.TrimEnd();

            if (line.TrimStart().StartsWith("```", StringComparison.Ordinal))
            {
                CloseLists();
                body.AppendLine(inCodeBlock ? "</code></pre>" : "<pre><code>");
                inCodeBlock = !inCodeBlock;
                continue;
            }

            if (inCodeBlock)
            {
                body.AppendLine(WebUtility.HtmlEncode(line));
                continue;
            }

            if (string.IsNullOrWhiteSpace(line))
            {
                CloseLists();
                continue;
            }

            Match heading = Regex.Match(line, @"^(#{1,6})\s+(.+)$");
            if (heading.Success)
            {
                CloseLists();
                int level = heading.Groups[1].Value.Length;
                string headingText = heading.Groups[2].Value;
                string id = level == 1 ? "top" : $"section-{++sectionNumber}";
                body.AppendLine($"<h{level} id=\"{id}\">{FormatInline(headingText)}</h{level}>");

                if (level is 2 or 3)
                {
                    string cssClass = level == 2 ? "nav-section" : "nav-subsection";
                    navigation.AppendLine(
                        $"<a class=\"{cssClass}\" href=\"#{id}\">{WebUtility.HtmlEncode(headingText)}</a>");
                }
                continue;
            }

            Match unordered = Regex.Match(line, @"^\s*[-*]\s+(.+)$");
            if (unordered.Success)
            {
                if (inOrderedList)
                {
                    body.AppendLine("</ol>");
                    inOrderedList = false;
                }

                if (!inUnorderedList)
                {
                    body.AppendLine("<ul>");
                    inUnorderedList = true;
                }

                body.AppendLine($"<li>{FormatInline(unordered.Groups[1].Value)}</li>");
                continue;
            }

            Match ordered = Regex.Match(line, @"^\s*\d+[.)]\s+(.+)$");
            if (ordered.Success)
            {
                if (inUnorderedList)
                {
                    body.AppendLine("</ul>");
                    inUnorderedList = false;
                }

                if (!inOrderedList)
                {
                    body.AppendLine("<ol>");
                    inOrderedList = true;
                }

                body.AppendLine($"<li>{FormatInline(ordered.Groups[1].Value)}</li>");
                continue;
            }

            CloseLists();
            body.AppendLine($"<p>{FormatInline(line.Trim())}</p>");
        }

        CloseLists();
        if (inCodeBlock)
            body.AppendLine("</code></pre>");

        return WrapHtml(body.ToString(), navigation.ToString(), navigationTitle);
    }

    private static string FormatInline(string value)
    {
        string result = WebUtility.HtmlEncode(value);
        result = Regex.Replace(
            result,
            @"\[(#[0-9a-fA-F]{6})\]",
            "<span class=\"color-swatch\" style=\"background-color:$1\"></span>");
        result = Regex.Replace(result, @"`([^`]+)`", "<code>$1</code>");
        result = Regex.Replace(result, @"\*\*([^*]+)\*\*", "<strong>$1</strong>");
        return result;
    }

    private static string WrapHtml(string body, string navigation, string navigationTitle)
    {
        return """
            <!doctype html>
            <html lang="ru">
            <head>
              <meta charset="utf-8">
              <meta http-equiv="X-UA-Compatible" content="IE=edge">
              <style>
                * { box-sizing: border-box; }
                html { scroll-behavior: smooth; }
                body { margin: 0; font-family: "Segoe UI", Arial, sans-serif; color: #243447; background: #f5f8fa; line-height: 1.55; }
                nav { position: fixed; z-index: 2; left: 0; top: 0; bottom: 0; width: 320px; padding: 24px 16px 40px; overflow-y: auto; color: white; background: #183f53; border-right: 1px solid #0e3041; }
                .nav-title { display: block; margin: 0 8px 20px; padding-bottom: 16px; color: white; font-size: 25px; font-weight: 700; text-decoration: none; border-bottom: 2px solid #00a0dc; }
                nav a { display: block; color: #eaf6fb; text-decoration: none; border-radius: 3px; }
                nav a:hover, nav a:focus { color: white; background: #0d79a8; }
                .nav-section { margin-top: 5px; padding: 10px 11px; font-size: 17px; line-height: 1.35; font-weight: 600; }
                .nav-subsection { padding: 7px 10px 7px 27px; color: #bdd9e5; font-size: 15px; line-height: 1.35; }
                .content { margin-left: 320px; padding: 28px 36px 60px; }
                main { max-width: 980px; margin: 0 auto; padding: 30px 44px 56px; background: transparent; border: 0; box-shadow: none; }
                h1 { margin: 0 0 24px; padding-bottom: 14px; color: #0078b8; font-size: 38px; border-bottom: 3px solid #00a0dc; }
                h2 { margin: 42px 0 14px; padding-top: 6px; color: #006da5; font-size: 29px; border-top: 1px solid #e5edf1; }
                h3 { margin: 26px 0 10px; color: #204f68; font-size: 23px; }
                h4 { margin: 22px 0 8px; color: #315f73; font-size: 20px; }
                h5 { margin: 18px 0 7px; color: #3d6577; font-size: 18px; }
                h6 { margin: 16px 0 6px; color: #496d7d; font-size: 17px; }
                p { margin: 8px 0 12px; }
                ul, ol { margin: 8px 0 16px 26px; padding: 0; }
                li { margin: 5px 0; }
                strong { color: #122b3a; }
                code { padding: 2px 5px; color: #8b1d56; background: #eef3f6; border: 1px solid #dbe4e9; font-family: Consolas, monospace; }
                pre { padding: 14px; overflow: auto; background: #eef3f6; border-left: 4px solid #00a0dc; }
                pre code { padding: 0; border: 0; }
                .color-swatch { display: inline-block; width: 22px; height: 16px; margin: 0 8px -3px 0; border: 1px solid #80919a; border-radius: 2px; }
                @media (max-width: 760px) {
                  nav { width: 250px; }
                  .nav-title { font-size: 22px; }
                  .nav-section { font-size: 15px; }
                  .nav-subsection { font-size: 14px; }
                  .content { margin-left: 250px; padding: 16px; }
                  main { padding: 24px; }
                }
              </style>
            </head>
            <body><nav><a class="nav-title" href="#top">
            """ + WebUtility.HtmlEncode(navigationTitle) + "</a>" + navigation +
            "</nav><div class=\"content\"><main>" + body + "</main></div></body></html>";
    }
}
