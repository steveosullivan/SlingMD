using System;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Dependency-free HTML -> Markdown converter used to turn an Outlook email's HTMLBody into
    /// readable Markdown for Obsidian: real headings, clickable [text](url) links, and embedded
    /// ![](url) images, instead of the flat plain-text dump you get from MailItem.Body.
    ///
    /// This is a pragmatic converter, not a full HTML parser. It handles the tags that matter for
    /// email/newsletter readability (a, img, headings, p/br, lists, bold/italic, blockquote, hr)
    /// and strips the rest. Newsletter HTML (e.g. Substack) converts cleanly; unusual markup
    /// degrades to plain text rather than breaking.
    /// </summary>
    internal static class HtmlToMarkdownConverter
    {
        private const RegexOptions Opts = RegexOptions.IgnoreCase | RegexOptions.Singleline;

        /// <summary>
        /// Converts <paramref name="html"/> (an Outlook HTMLBody) to Markdown.
        /// If the HTML is empty or conversion yields nothing useful, falls back to
        /// <paramref name="plainTextFallback"/> (typically MailItem.Body).
        /// </summary>
        public static string Convert(string html, string plainTextFallback)
        {
            if (string.IsNullOrWhiteSpace(html))
            {
                return plainTextFallback ?? string.Empty;
            }

            string s = html;

            try
            {
                // 1. Drop everything that is not visible content.
                s = Regex.Replace(s, @"<!--.*?-->", string.Empty, Opts);
                s = Regex.Replace(s, @"<head\b.*?</head>", string.Empty, Opts);
                s = Regex.Replace(s, @"<style\b.*?</style>", string.Empty, Opts);
                s = Regex.Replace(s, @"<script\b.*?</script>", string.Empty, Opts);

                // 2. Prefer the <body> content if present.
                Match body = Regex.Match(s, @"<body\b[^>]*>(.*?)</body>", Opts);
                if (body.Success)
                {
                    s = body.Groups[1].Value;
                }

                // 3. Line breaks and horizontal rules.
                s = Regex.Replace(s, @"<br\s*/?>", "\n", Opts);
                s = Regex.Replace(s, @"<hr\s*/?>", "\n\n---\n\n", Opts);

                // 4. Headings -> #.. (strip any inline tags inside the heading text).
                for (int level = 1; level <= 6; level++)
                {
                    string prefix = new string('#', level) + " ";
                    s = Regex.Replace(
                        s,
                        @"<h" + level + @"\b[^>]*>(.*?)</h" + level + ">",
                        m => "\n\n" + prefix + StripTags(m.Groups[1].Value).Trim() + "\n\n",
                        Opts);
                }

                // 5. Images -> ![alt](src). Skip tracking pixels (1x1 or Substack "p.gif").
                s = Regex.Replace(s, @"<img\b[^>]*>", m => ConvertImage(m.Value), Opts);

                // 6. Links -> [text](href). Done AFTER images so a linked image becomes
                //    [![alt](img)](href), which Obsidian renders as a clickable image.
                s = Regex.Replace(
                    s,
                    @"<a\b[^>]*?href\s*=\s*(?:""([^""]*)""|'([^']*)')[^>]*>(.*?)</a>",
                    m =>
                    {
                        string href = !string.IsNullOrEmpty(m.Groups[1].Value) ? m.Groups[1].Value : m.Groups[2].Value;
                        string text = StripTags(m.Groups[3].Value).Trim();
                        href = href.Trim();
                        if (string.IsNullOrEmpty(href)) return text;
                        if (string.IsNullOrEmpty(text)) text = href;
                        return "[" + text + "](" + href + ")";
                    },
                    Opts);

                // 7. Bold / italic.
                s = Regex.Replace(s, @"<(?:strong|b)\b[^>]*>(.*?)</(?:strong|b)>", m => "**" + StripTags(m.Groups[1].Value).Trim() + "**", Opts);
                s = Regex.Replace(s, @"<(?:em|i)\b[^>]*>(.*?)</(?:em|i)>", m => "*" + StripTags(m.Groups[1].Value).Trim() + "*", Opts);

                // 8. List items -> "- item".
                s = Regex.Replace(s, @"<li\b[^>]*>(.*?)</li>", m => "\n- " + StripTags(m.Groups[1].Value).Trim(), Opts);

                // 9. Blockquotes -> "> " prefixed lines.
                s = Regex.Replace(s, @"<blockquote\b[^>]*>(.*?)</blockquote>", m => QuoteLines(StripTags(m.Groups[1].Value)), Opts);

                // 10. Block-level closers become paragraph breaks.
                s = Regex.Replace(s, @"</(p|div|tr|table|ul|ol|section|article|header|footer)>", "\n\n", Opts);

                // 11. Strip every remaining tag.
                s = StripTags(s);

                // 12. Decode HTML entities (&amp; &nbsp; &#8217; etc.).
                s = WebUtility.HtmlDecode(s);

                // 13. Whitespace cleanup.
                s = CleanWhitespace(s);

                // If the result is essentially empty, fall back to the plain-text body.
                if (string.IsNullOrWhiteSpace(s))
                {
                    return plainTextFallback ?? string.Empty;
                }

                return s;
            }
            catch (Exception)
            {
                // Never let body conversion abort the sling — fall back to plain text.
                return plainTextFallback ?? string.Empty;
            }
        }

        private static string ConvertImage(string imgTag)
        {
            string src = GetAttribute(imgTag, "src");
            if (string.IsNullOrWhiteSpace(src))
            {
                return string.Empty;
            }

            // Skip common tracking pixels: explicit 1x1 dimensions, or Substack open-tracker gifs.
            string width = GetAttribute(imgTag, "width");
            string height = GetAttribute(imgTag, "height");
            if (width == "1" || height == "1")
            {
                return string.Empty;
            }
            if (src.IndexOf("/p.gif", StringComparison.OrdinalIgnoreCase) >= 0 ||
                src.IndexOf("/open?", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return string.Empty;
            }

            string alt = GetAttribute(imgTag, "alt") ?? string.Empty;
            return "\n\n![" + alt.Trim() + "](" + src.Trim() + ")\n\n";
        }

        private static string GetAttribute(string tag, string name)
        {
            Match m = Regex.Match(tag, name + @"\s*=\s*(?:""([^""]*)""|'([^']*)')", Opts);
            if (!m.Success)
            {
                return null;
            }
            return !string.IsNullOrEmpty(m.Groups[1].Value) ? m.Groups[1].Value : m.Groups[2].Value;
        }

        private static string StripTags(string input)
        {
            return string.IsNullOrEmpty(input) ? input : Regex.Replace(input, @"<[^>]+>", string.Empty, Opts);
        }

        private static string QuoteLines(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }
            var sb = new StringBuilder("\n\n");
            foreach (string line in text.Split('\n'))
            {
                sb.Append("> ").Append(line.Trim()).Append('\n');
            }
            sb.Append('\n');
            return sb.ToString();
        }

        private static string CleanWhitespace(string s)
        {
            // Normalize newlines.
            s = s.Replace("\r\n", "\n").Replace("\r", "\n");
            // Non-breaking spaces -> regular spaces.
            s = s.Replace(' ', ' ');
            // Trim trailing spaces on each line.
            s = Regex.Replace(s, @"[ \t]+\n", "\n");
            // Collapse runs of spaces/tabs.
            s = Regex.Replace(s, @"[ \t]{2,}", " ");
            // Collapse 3+ blank lines into a single blank line.
            s = Regex.Replace(s, @"\n{3,}", "\n\n");
            return s.Trim();
        }
    }
}
