namespace Tabel;

using System.Net;
using System.Text;

internal sealed class ProtocolReportForm : Form
{
    public ProtocolReportForm(
        string title,
        string summary,
        string resultCaption,
        string resultPath,
        string protocolPath)
    {
        Text = title;
        StartPosition = FormStartPosition.CenterParent;
        MinimumSize = new Size(700, 500);
        Size = new Size(980, 760);
        KeyPreview = true;

        var browser = new WebBrowser
        {
            Dock = DockStyle.Fill,
            AllowWebBrowserDrop = false,
            IsWebBrowserContextMenuEnabled = true,
            ScriptErrorsSuppressed = true,
            WebBrowserShortcutsEnabled = true,
            DocumentText = BuildHtml(title, summary, resultCaption, resultPath, protocolPath)
        };

        Controls.Add(browser);
        WindowStateStore.Attach(this, "ProtocolReport");
        KeyDown += (_, e) =>
        {
            if (e.KeyCode == Keys.Escape)
                Close();
        };
    }

    private static string BuildHtml(
        string title,
        string summary,
        string resultCaption,
        string resultPath,
        string protocolPath)
    {
        string protocol;
        try
        {
            protocol = File.ReadAllText(protocolPath, Encoding.UTF8);
        }
        catch (Exception ex)
        {
            protocol = "# Протокол недоступен\n" + ex.Message;
        }

        return """
            <!doctype html>
            <html lang="ru">
            <head>
              <meta charset="utf-8">
              <meta http-equiv="X-UA-Compatible" content="IE=edge">
              <style>
                * { box-sizing: border-box; }
                body { margin: 0; padding: 28px 34px 60px; font-family: "Segoe UI", Arial, sans-serif; color: #263b47; background: #edf3f6; line-height: 1.45; }
                main { max-width: 1120px; margin: 0 auto; }
                .hero { padding: 26px 30px; color: white; background: #0099d8; border-radius: 4px; }
                .eyebrow { margin-bottom: 5px; font-size: 14px; opacity: .9; }
                h1 { margin: 0; font-size: 30px; }
                .summary { margin: 18px 0; padding: 22px 24px; background: white; border: 1px solid #d6e1e7; }
                .summary h2, .protocol h2 { margin: 0 0 14px; color: #0879ad; font-size: 22px; }
                .lead { margin: 0 0 16px; font-size: 17px; }
                .metrics { margin: -6px; }
                .metric { display: inline-block; min-width: 190px; margin: 6px; padding: 12px 15px; vertical-align: top; background: #f2f7fa; border-left: 4px solid #00a0dc; }
                .metric-label { display: block; color: #607784; font-size: 13px; }
                .metric-value { display: block; margin-top: 3px; color: #142f3d; font-size: 19px; font-weight: 700; }
                .file { margin-top: 17px; padding: 12px 15px; background: #f7f9fa; border: 1px solid #dce5e9; }
                .file-label { color: #607784; font-size: 13px; }
                .path { margin-top: 4px; font-family: Consolas, monospace; overflow-wrap: break-word; word-wrap: break-word; }
                .protocol { padding: 22px 24px 34px; background: white; border: 1px solid #d6e1e7; }
                .protocol-section { margin: 20px 0 28px; }
                .protocol-section h3 { margin: 0 0 10px; padding-bottom: 7px; color: #205b76; font-size: 18px; border-bottom: 2px solid #d9e9f0; }
                .record { margin: 7px 0; padding: 10px 12px; background: #f8fafb; border-left: 3px solid #a9c5d2; overflow-wrap: break-word; word-wrap: break-word; }
                .cell { display: inline-block; margin: 2px 7px 2px 0; padding: 3px 7px; vertical-align: top; background: white; border: 1px solid #dce5e9; }
                .empty { color: #71838c; font-style: italic; }
                .protocol-path { margin-top: 28px; color: #657a85; font-size: 13px; overflow-wrap: break-word; word-wrap: break-word; }
              </style>
            </head>
            <body><main>
            """
            + $"<div class=\"hero\"><div class=\"eyebrow\">Операция завершена</div><h1>{Encode(title)}</h1></div>"
            + "<section class=\"summary\"><h2>Краткий итог</h2>"
            + BuildSummaryHtml(summary)
            + $"<div class=\"file\"><div class=\"file-label\">{Encode(resultCaption)}</div><div class=\"path\">{Encode(resultPath)}</div></div></section>"
            + "<section class=\"protocol\"><h2>Полный протокол</h2>"
            + BuildProtocolHtml(protocol)
            + $"<div class=\"protocol-path\">Файл протокола: {Encode(protocolPath)}</div></section>"
            + "</main></body></html>";
    }

    private static string BuildSummaryHtml(string summary)
    {
        int fileIndex = summary.IndexOf(" Файл:", StringComparison.OrdinalIgnoreCase);
        if (fileIndex >= 0)
            summary = summary[..fileIndex];

        var lead = new StringBuilder();
        var metrics = new StringBuilder();

        foreach (string part in summary.Split('.', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            int separator = part.IndexOf(':');
            if (separator <= 0 || separator == part.Length - 1)
            {
                lead.Append($"<p class=\"lead\">{Encode(part)}.</p>");
                continue;
            }

            string label = part[..separator].Trim();
            string value = part[(separator + 1)..].Trim();
            metrics.Append("<div class=\"metric\">")
                .Append($"<span class=\"metric-label\">{Encode(label)}</span>")
                .Append($"<span class=\"metric-value\">{Encode(value)}</span></div>");
        }

        if (metrics.Length > 0)
            lead.Append("<div class=\"metrics\">").Append(metrics).Append("</div>");

        return lead.ToString();
    }

    private static string BuildProtocolHtml(string protocol)
    {
        var html = new StringBuilder();
        bool sectionOpen = false;

        foreach (string sourceLine in protocol.Replace("\r\n", "\n").Split('\n'))
        {
            string line = sourceLine.Trim();
            if (line.Length == 0)
                continue;

            if (line.StartsWith("# ", StringComparison.Ordinal))
            {
                if (sectionOpen)
                    html.Append("</div>");

                html.Append($"<div class=\"protocol-section\"><h3>{Encode(line[2..].Trim())}</h3>");
                sectionOpen = true;
                continue;
            }

            string[] fields = line.Split(';');
            html.Append("<div class=\"record\">");
            if (fields.Length > 1)
            {
                foreach (string field in fields)
                {
                    if (!string.IsNullOrWhiteSpace(field))
                        html.Append($"<span class=\"cell\">{Encode(field.Trim())}</span>");
                }
            }
            else
            {
                string cssClass = line.Equals("Нет записей.", StringComparison.OrdinalIgnoreCase) ? "empty" : "";
                html.Append($"<span class=\"{cssClass}\">{Encode(line)}</span>");
            }

            html.Append("</div>");
        }

        if (sectionOpen)
            html.Append("</div>");

        return html.ToString();
    }

    private static string Encode(string value) => WebUtility.HtmlEncode(value);
}
