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

        var browser = new WebBrowser
        {
            Dock = DockStyle.Fill,
            AllowWebBrowserDrop = false,
            IsWebBrowserContextMenuEnabled = true,
            ScriptErrorsSuppressed = true,
            WebBrowserShortcutsEnabled = true,
            DocumentText = LoadHelpHtml()
        };

        Controls.Add(browser);
        KeyDown += (_, e) =>
        {
            if (e.KeyCode == Keys.Escape)
                Close();
        };
    }

    private static string LoadHelpHtml()
    {
        string helpPath = Path.Combine(AppContext.BaseDirectory, "docs", "Справка.md");

        try
        {
            string markdown = File.ReadAllText(helpPath, Encoding.UTF8);
            return MarkdownToHtml(markdown);
        }
        catch (Exception ex)
        {
            return WrapHtml($"<h1>Справка недоступна</h1><p>{WebUtility.HtmlEncode(ex.Message)}</p>");
        }
    }

    private static string MarkdownToHtml(string markdown)
    {
        var body = new StringBuilder();
        bool inUnorderedList = false;
        bool inOrderedList = false;
        bool inCodeBlock = false;

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

            Match heading = Regex.Match(line, @"^(#{1,4})\s+(.+)$");
            if (heading.Success)
            {
                CloseLists();
                int level = heading.Groups[1].Value.Length;
                body.AppendLine($"<h{level}>{FormatInline(heading.Groups[2].Value)}</h{level}>");
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

        return WrapHtml(body.ToString());
    }

    private static string FormatInline(string value)
    {
        string result = WebUtility.HtmlEncode(value);
        result = Regex.Replace(result, @"`([^`]+)`", "<code>$1</code>");
        result = Regex.Replace(result, @"\*\*([^*]+)\*\*", "<strong>$1</strong>");
        return result;
    }

    private static string WrapHtml(string body)
    {
        return """
            <!doctype html>
            <html lang="ru">
            <head>
              <meta charset="utf-8">
              <meta http-equiv="X-UA-Compatible" content="IE=edge">
              <style>
                body { margin: 0; padding: 28px 36px 60px; font-family: "Segoe UI", Arial, sans-serif; color: #243447; background: #f5f8fa; line-height: 1.55; }
                main { max-width: 920px; margin: 0 auto; padding: 28px 42px 48px; background: white; border: 1px solid #dce5eb; }
                h1 { margin: 0 0 24px; padding-bottom: 14px; color: #0078b8; font-size: 32px; border-bottom: 3px solid #00a0dc; }
                h2 { margin: 34px 0 12px; color: #006da5; font-size: 23px; }
                h3 { margin: 24px 0 8px; color: #204f68; font-size: 18px; }
                p { margin: 8px 0 12px; }
                ul, ol { margin: 8px 0 16px 26px; padding: 0; }
                li { margin: 5px 0; }
                strong { color: #122b3a; }
                code { padding: 2px 5px; color: #8b1d56; background: #eef3f6; border: 1px solid #dbe4e9; font-family: Consolas, monospace; }
                pre { padding: 14px; overflow: auto; background: #eef3f6; border-left: 4px solid #00a0dc; }
                pre code { padding: 0; border: 0; }
              </style>
            </head>
            <body><main>
            """ + body + "</main></body></html>";
    }
}
