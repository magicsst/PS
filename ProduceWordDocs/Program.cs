using HtmlAgilityPack;
using System.Net;
using System.Text.RegularExpressions;

namespace ProduceWordDocs
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            // 1) Ρυθμίσεις
            var baseUrl = "https://ploumis-sotiropoulos.gr";
            var listingUrl = $"{baseUrl}/en-us/sales/independent-buildings";
            var outRoot = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            var imgCache = Path.Combine(outRoot, "Images");
            Directory.CreateDirectory(outRoot);
            Directory.CreateDirectory(imgCache);

            Console.WriteLine("Fetching listing...");
            var http = new HttpClient(new HttpClientHandler
            {
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
            });
            http.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (compatible; ExportBot/1.0)");

            // 2) Φέρε τα links "Further information"
            var listHtml = await http.GetStringAsync(listingUrl);
            var listDoc = new HtmlDocument();
            listDoc.LoadHtml(listHtml);

            var furtherLinks = listDoc.DocumentNode
                .SelectNodes("//a[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'),'further information')]")
                ?.Select(a => a.GetAttributeValue("href", null))
                .Where(h => !string.IsNullOrWhiteSpace(h))
                .Select(h => h!.StartsWith("http", StringComparison.OrdinalIgnoreCase) ? h : baseUrl + h)
                .Distinct()
                .ToList() ?? new List<string>();

            if (furtherLinks.Count == 0)
            {
                Console.WriteLine("No properties found.");
                return;
            }

            var items = new List<Property>();

            // 3) Για κάθε ακίνητο -> φέρε λεπτομέρειες, εικόνες κ.λπ.
            foreach (var url in furtherLinks)
            {
                Console.WriteLine($"Fetching: {url}");
                var html = await http.GetStringAsync(url);
                var doc = new HtmlDocument();
                doc.LoadHtml(html);

                // Τίτλος: πρώτο <h1> ή # The ...
                var title = doc.DocumentNode.SelectSingleNode("//h1")?.InnerText?.Trim();
                if (string.IsNullOrWhiteSpace(title))
                {
                    title = doc.DocumentNode.SelectSingleNode("//h2")?.InnerText?.Trim() ?? "Unknown Title";
                }

                // Περιγραφή: οι πρώτες 3-5 παράγραφοι πριν/γύρω από την ένδειξη (Code: xxxxxxx)
                var allTextNodes = doc.DocumentNode.SelectNodes("//p|//div|//span") ?? new HtmlNodeCollection(null);
                string description = "";
                foreach (var n in allTextNodes)
                {
                    var t = WebUtility.HtmlDecode(n.InnerText ?? "").Trim();
                    if (string.IsNullOrWhiteSpace(t)) continue;
                    if (t.StartsWith("(Code:", StringComparison.OrdinalIgnoreCase)) break;
                    // αγνόησε κοινά footer sections
                    if (t.Contains("EXCLUSIVITIES", StringComparison.OrdinalIgnoreCase)) break;
                    if (t.Length > 40) description += t + Environment.NewLine + Environment.NewLine;
                }
                description = description.Trim();

                // Code: πιάσε "(Code: 1234567)" ή από Overview
                var codeMatch = Regex.Match(doc.DocumentNode.InnerText, @"\(Code:\s*(\d+)\)", RegexOptions.IgnoreCase);
                var code = codeMatch.Success ? codeMatch.Groups[1].Value : "";

                // Overview πεδία (αν υπάρχουν). Απλή προσέγγιση: μάζεψε ζεύγη γνωστών labels
                var overview = ExtractOverviewPairs(doc.DocumentNode.InnerText);

                // Τιμή: από overview ή τμήμα “PRICE …”
                string price = overview.TryGetValue("Price", out var p) ? p : ExtractPriceFromListing(doc);

                // Εικόνες (gallery)
                var imgUrls = doc.DocumentNode
                    .SelectNodes("//img")
                    ?.Select(i => i.GetAttributeValue("src", null))
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Select(s => s!.StartsWith("http", StringComparison.OrdinalIgnoreCase) ? s : baseUrl + s)
                    .Distinct()
                    .ToList() ?? new List<string>();

                // Κατέβασε μέχρι 10 εικόνες (ρυθμίσιμο)
                var localImages = new List<string>();
                int take = Math.Min(10, imgUrls.Count);
                for (int i = 0; i < take; i++)
                {
                    try
                    {
                        var u = imgUrls[i];
                        var ext = Path.GetExtension(new Uri(u).AbsolutePath);
                        if (string.IsNullOrWhiteSpace(ext)) ext = ".jpg";
                        var fname = SanitizeFileName($"{(string.IsNullOrWhiteSpace(code) ? "no-code" : code)}_{i + 1}{ext}");
                        var fpath = Path.Combine(imgCache, fname);
                        var bytes = await http.GetByteArrayAsync(u);
                        await File.WriteAllBytesAsync(fpath, bytes);
                        localImages.Add(fpath);
                    }
                    catch { /* ignore single image errors */ }
                }

                items.Add(new Property
                {
                    Title = title ?? "",
                    Code = code,
                    Url = url,
                    Price = price,
                    Description = description,
                    Overview = overview,
                    ImageFiles = localImages
                });
            }

            // 4) Εξαγωγή Word:
            var exporter = new WordExporter(outRoot);

            // (a) .dotx template με λίστα
            var dotxPath = exporter.CreateSummaryTemplateDotx(items, "Independent Buildings — Data Export");

            // (b) .docx ανά ακίνητο
            foreach (var it in items)
            {
                var docx = exporter.CreatePropertyDocx(it);
                exporter.ConvertDocxToPdf(docx);
            }

            Console.WriteLine("\nDone!");
            Console.WriteLine("Summary template: " + dotxPath);
            Console.WriteLine("Per-property DOCX files in: " + outRoot);
        }

        static Dictionary<string, string> ExtractOverviewPairs(string text)
        {
            // Απλό parser: ψάχνει γνωστές ετικέτες και πιάνει την επόμενη τιμή στη ροή του κειμένου
            // Θα καλύψει πεδία όπως εμφανίζονται στις σελίδες: Property Code, Property Category, Interior Space, Land Area, Energy Efficiency Class, Price, Objective Tax Value, Annual Tax, Offered
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            string[] labels =
            {
            "Property Code",
            "Property Category",
            "Interior Space",
            "Land Area",
            "Energy Efficiency Class",
            "Price",
            "Objective Tax Value",
            "Annual Tax",
            "Offered"
        };

            // Καθάρισε whitespace
            var clean = Regex.Replace(text ?? "", @"\s+", " ").Trim();

            foreach (var label in labels)
            {
                var rx = new Regex(label + @"\s+([^\n\r#]+?)(?=\s+(Property Code|Property Category|Interior Space|Land Area|Energy Efficiency Class|Price|Objective Tax Value|Annual Tax|Offered|EXCLUSIVITIES|RESIDENTIAL|VACATION|$))",
                                   RegexOptions.IgnoreCase);
                var m = rx.Match(clean);
                if (m.Success)
                {
                    var val = m.Groups[1].Value.Trim();
                    map[label] = val;
                }
            }
            return map;
        }

        static string ExtractPriceFromListing(HtmlDocument doc)
        {
            // Ψάξε για "PRICE" δίπλα στον τίτλο/κεφαλίδα
            var text = doc.DocumentNode.InnerText;
            var m = Regex.Match(text, @"PRICE\s+([^\n\r]+)", RegexOptions.IgnoreCase);
            return m.Success ? m.Groups[1].Value.Trim() : "";
        }

        static string SanitizeFileName(string name)
        {
            foreach (var c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name;
        }
    }

    public record Property
    {
        public string Title { get; set; } = "";
        public string Code { get; set; } = "";
        public string Url { get; set; } = "";
        public string Price { get; set; } = "";
        public string Description { get; set; } = "";
        public Dictionary<string, string> Overview { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        public List<string> ImageFiles { get; set; } = new();
    }
}
