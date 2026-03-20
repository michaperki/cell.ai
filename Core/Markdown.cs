using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace SpreadsheetApp.Core
{
    public static class Markdown
    {
        // Convert a subset of Markdown to HTML for in-app viewing.
        // Supports: headings, paragraphs, code blocks, inline code, hr, lists, blockquotes, bold/italic, links, images.
        public static string ToHtml(string md)
        {
            if (string.IsNullOrEmpty(md)) return string.Empty;
            var lines = SplitLines(md);
            var sb = new StringBuilder();
            bool inCode = false;
            string? codeLang = null;
            var para = new List<string>();
            bool inUl = false, inOl = false, inBlockquote = false;

            void FlushPara()
            {
                if (para.Count > 0)
                {
                    var text = string.Join(" ", para);
                    sb.Append("<p>").Append(Inline(text)).Append("</p>\n");
                    para.Clear();
                }
            }

            void CloseLists()
            {
                if (inUl) { sb.Append("</ul>\n"); inUl = false; }
                if (inOl) { sb.Append("</ol>\n"); inOl = false; }
            }

            void CloseBlockquote()
            {
                if (inBlockquote) { sb.Append("</blockquote>\n"); inBlockquote = false; }
            }

            for (int i = 0; i < lines.Count; i++)
            {
                string line = lines[i];

                // Code block fences ```lang
                var fenceMatch = Regex.Match(line, @"^```(.*)$");
                if (fenceMatch.Success)
                {
                    if (!inCode)
                    {
                        FlushPara();
                        CloseLists();
                        CloseBlockquote();
                        inCode = true;
                        codeLang = fenceMatch.Groups[1].Value.Trim();
                        sb.Append("<pre><code");
                        if (!string.IsNullOrEmpty(codeLang)) sb.Append(" class=\"language-").Append(Html(codeLang)).Append("\"");
                        sb.Append(">");
                    }
                    else
                    {
                        sb.Append("</code></pre>\n");
                        inCode = false; codeLang = null;
                    }
                    continue;
                }

                if (inCode)
                {
                    sb.Append(Html(line)).Append("\n");
                    continue;
                }

                // Blank line => paragraph/list/blockquote boundaries
                if (string.IsNullOrWhiteSpace(line))
                {
                    FlushPara();
                    CloseLists();
                    CloseBlockquote();
                    continue;
                }

                // Blockquote
                if (Regex.IsMatch(line, @"^> "))
                {
                    FlushPara();
                    CloseLists();
                    if (!inBlockquote) { sb.Append("<blockquote>\n"); inBlockquote = true; }
                    sb.Append(Inline(line.Substring(2))).Append("<br/>");
                    sb.Append('\n');
                    continue;
                }

                // Horizontal rule
                if (Regex.IsMatch(line.Trim(), @"^(\n|\r|\s)*(\*\s*\*\s*\*|-\s*-\s*-|_\s*_\s*_)$"))
                {
                    FlushPara();
                    CloseLists();
                    CloseBlockquote();
                    sb.Append("<hr/>\n");
                    continue;
                }

                // Headings #..######
                var h = Regex.Match(line, @"^(#{1,6})\s+(.*)$");
                if (h.Success)
                {
                    FlushPara();
                    CloseLists();
                    CloseBlockquote();
                    int level = Math.Min(6, h.Groups[1].Value.Length);
                    string text = h.Groups[2].Value.Trim();
                    sb.Append('<').Append('h').Append(level).Append('>')
                      .Append(Inline(text))
                      .Append("</h").Append(level).Append(">\n");
                    continue;
                }

                // Lists
                var ul = Regex.Match(line, @"^\t*\s*([*+-])\s+(.*)$");
                if (ul.Success)
                {
                    FlushPara();
                    CloseBlockquote();
                    if (!inUl) { CloseLists(); sb.Append("<ul>\n"); inUl = true; }
                    sb.Append("<li>").Append(Inline(ul.Groups[2].Value.Trim())).Append("</li>\n");
                    continue;
                }
                var ol = Regex.Match(line, @"^\t*\s*(\d+)\.\s+(.*)$");
                if (ol.Success)
                {
                    FlushPara();
                    CloseBlockquote();
                    if (!inOl) { CloseLists(); sb.Append("<ol>\n"); inOl = true; }
                    sb.Append("<li>").Append(Inline(ol.Groups[2].Value.Trim())).Append("</li>\n");
                    continue;
                }

                // Default: paragraph continuation
                para.Add(line.Trim());
            }

            if (inCode) sb.Append("</code></pre>\n");
            FlushPara();
            CloseLists();
            CloseBlockquote();

            return sb.ToString();
        }

        public static string WrapHtml(string title, string body)
        {
            var sb = new StringBuilder();
            sb.Append("<!DOCTYPE html><html><head><meta charset=\"utf-8\"/>");
            sb.Append("<meta http-equiv=\"X-UA-Compatible\" content=\"IE=edge\"/>");
            sb.Append("<title>").Append(Html(string.IsNullOrEmpty(title) ? "Document" : title)).Append("</title>");
            sb.Append("<style>");
            sb.Append(Styles);
            sb.Append("</style></head><body>");
            sb.Append(body);
            sb.Append("</body></html>");
            return sb.ToString();
        }

        private static string Styles => @"
body { font-family: -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif; color: #222; background: #fff; padding: 18px; line-height: 1.5; }
h1 { font-size: 1.8em; margin: 0.6em 0 0.4em; }
h2 { font-size: 1.5em; margin: 0.8em 0 0.4em; }
h3 { font-size: 1.25em; margin: 0.8em 0 0.3em; }
hr { border: none; border-top: 1px solid #e0e0e0; margin: 1em 0; }
p { margin: 0.5em 0; }
ul, ol { margin: 0.4em 0 0.6em 1.6em; }
code { background: #f6f8fa; border: 1px solid #eaecef; padding: 0 4px; border-radius: 4px; font-family: Consolas, Menlo, Monaco, monospace; }
pre { background: #f6f8fa; border: 1px solid #eaecef; padding: 10px; overflow: auto; border-radius: 6px; }
blockquote { margin: 0.5em 0; padding: 0.4em 0.8em; background: #f9f9f9; border-left: 4px solid #d0d0d0; color: #444; }
a { color: #0366d6; text-decoration: none; }
a:hover { text-decoration: underline; }
img { max-width: 100%; }
";

        private static string Inline(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            // Escape first
            string s = Html(text);
            // Images ![alt](url)
            s = Regex.Replace(s, @"!\[([^\]]*)\]\(([^\)]+)\)", m => $"<img alt=\"{Html(m.Groups[1].Value)}\" src=\"{Html(m.Groups[2].Value)}\"/>");
            // Links [text](url)
            s = Regex.Replace(s, @"\[([^\]]+)\]\(([^\)]+)\)", m => $"<a href=\"{Html(m.Groups[2].Value)}\">{m.Groups[1].Value}</a>");
            // Bold **text** or __text__
            s = Regex.Replace(s, @"(\*\*|__)(.+?)\1", m => $"<strong>{m.Groups[2].Value}</strong>");
            // Italic *text* or _text_
            s = Regex.Replace(s, @"(\*|_)([^*_].+?)\1", m => $"<em>{m.Groups[2].Value}</em>");
            // Inline code `code`
            s = Regex.Replace(s, @"`([^`]+)`", m => $"<code>{m.Groups[1].Value}</code>");
            return s;
        }

        private static string Html(string s)
        {
            return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
        }

        private static List<string> SplitLines(string s)
        {
            var list = new List<string>();
            var sb = new StringBuilder();
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '\r')
                {
                    if (i + 1 < s.Length && s[i + 1] == '\n') i++;
                    list.Add(sb.ToString()); sb.Clear();
                }
                else if (c == '\n')
                {
                    list.Add(sb.ToString()); sb.Clear();
                }
                else sb.Append(c);
            }
            list.Add(sb.ToString());
            return list;
        }
    }
}

