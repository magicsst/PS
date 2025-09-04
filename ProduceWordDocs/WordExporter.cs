using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SixLabors.ImageSharp;


namespace ProduceWordDocs
{
    public class WordExporter
    {
        private readonly string _outRoot;
        public WordExporter(string outRoot) { _outRoot = outRoot; }

        public string CreateSummaryTemplateDotx(IEnumerable<Property> props, string title)
        {
            var path = Path.Combine(_outRoot, "Independent_Buildings_Summary.dotx");
            using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Template);
            var main = doc.AddMainDocumentPart();
            main.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(new Body());
            var body = main.Document.Body!;

            body.Append(MakeHeading(title, 1));
            body.Append(MakeParagraph($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm}"));
            body.Append(MakeParagraph("Source: https://ploumis-sotiropoulos.gr/en-us/sales/independent-buildings"));

            int i = 1;
            foreach (var p in props)
            {
                body.Append(MakeHeading($"{i}. {p.Title}", 2));
                body.Append(MakeKeyVal("Property Code", p.Code));
                body.Append(MakeKeyVal("Price", string.IsNullOrWhiteSpace(p.Price) ? "-" : p.Price));
                if (p.Overview.TryGetValue("Interior Space", out var ispc)) body.Append(MakeKeyVal("Interior Space", ispc));
                if (p.Overview.TryGetValue("Land Area", out var land)) body.Append(MakeKeyVal("Land Area", land));
                if (p.Overview.TryGetValue("Energy Efficiency Class", out var ee)) body.Append(MakeKeyVal("Energy Class", ee));
                body.Append(MakeHyperlink("URL", p.Url));
                body.Append(new Paragraph(new Run(new Text(" ")))); // spacer
                i++;
            }

            main.Document.Save();
            return path;
        }

        public string CreatePropertyDocx(Property p)
        {
            var fileName = $"{(string.IsNullOrWhiteSpace(p.Code) ? "no-code" : p.Code)} - {Sanitize(p.Title)}.docx";
            var path = Path.Combine(_outRoot, fileName);

            using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
            var main = doc.AddMainDocumentPart();
            main.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(new Body());
            var body = main.Document.Body!;

            body.Append(MakeHeading(p.Title, 1));
            if (!string.IsNullOrWhiteSpace(p.Code)) body.Append(MakeKeyVal("Property Code", p.Code));
            if (p.Overview.TryGetValue("Property Category", out var cat)) body.Append(MakeKeyVal("Category", cat));
            if (!string.IsNullOrWhiteSpace(p.Price)) body.Append(MakeKeyVal("Price", p.Price));
            if (p.Overview.TryGetValue("Interior Space", out var ispc)) body.Append(MakeKeyVal("Interior Space", ispc));
            if (p.Overview.TryGetValue("Land Area", out var land)) body.Append(MakeKeyVal("Land Area", land));
            if (p.Overview.TryGetValue("Energy Efficiency Class", out var ee)) body.Append(MakeKeyVal("Energy Class", ee));
            if (p.Overview.TryGetValue("Objective Tax Value", out var otv)) body.Append(MakeKeyVal("Objective Tax Value", otv));
            if (p.Overview.TryGetValue("Annual Tax", out var tax)) body.Append(MakeKeyVal("Annual Tax", tax));
            body.Append(MakeHyperlink("URL", p.Url));

            // Περιγραφή
            if (!string.IsNullOrWhiteSpace(p.Description))
            {
                body.Append(new Paragraph(new Run(new Break())));
                body.Append(MakeHeading("Description", 2));
                body.Append(MakeParagraph(p.Description));
            }

            // Εικόνες (έως 10)
            if (p.ImageFiles.Count > 0)
            {
                body.Append(new Paragraph(new Run(new Break())));
                body.Append(MakeHeading("Images", 2));

                foreach (var imgPath in p.ImageFiles)
                {
                    try
                    {
                        AddImage(main, imgPath, 6.0 /* inches width */);
                        body.Append(new Paragraph(new Run(new Break()))); // spacer between images
                    }
                    catch { /* ignore single image errors */ }
                }
            }

            main.Document.Save();
            return path;
        }

        private static Paragraph MakeHeading(string text, int level)
        {
            return new Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId() { Val = "Heading" + Math.Clamp(level, 1, 3) }),
                new Run(new Text(text)));
        }

        private static Paragraph MakeParagraph(string text)
        {
            return new Paragraph(new Run(new Text(text)));
        }

        private static Paragraph MakeKeyVal(string key, string val)
        {
            return new Paragraph(
                new Run(new RunProperties(new Bold()), new Text(key + ": ")),
                new Run(new Text(val ?? "-")));
        }

        private static Paragraph MakeHyperlink(string label, string url)
        {
            // Απλός συμβολισμός λινκ (ως κείμενο). Για ενεργό hyperlink απαιτείται Id στο rels, παραλείπεται για συντομία.
            return MakeKeyVal(label, url);
        }

        private static string Sanitize(string name)
        {
            foreach (var c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name;
        }

        private static void AddImage(MainDocumentPart mainPart, string imagePath, double targetWidthInches)
        {

            var partType = GuessImagePartType(imagePath);   // PartTypeInfo
            var imgPart = mainPart.AddImagePart(partType); // <-- ΔΕΧΕΤΑΙ PartTypeInfo στη v3

            //var imgPartType = GuessImagePartType(imagePath);
            //var imgPart = mainPart.AddImagePart(imgPartType);

            using (var stream = File.OpenRead(imagePath))
            {
                imgPart.FeedData(stream);
            }

            // Διαστάσεις εικόνας
            using var image = SixLabors.ImageSharp.Image.Load(imagePath);
            var pxW = image.Width;
            var pxH = image.Height;

            // Μετατροπή px -> EMU υποθέτοντας 96 dpi
            const double emusPerInch = 914400;
            const double dpi = 96.0;
            var aspect = (double)pxH / pxW;

            var widthEmu = (long)(targetWidthInches * emusPerInch);
            var heightEmu = (long)(targetWidthInches * aspect * emusPerInch);


            using (var stream = File.OpenRead(imagePath))
                imgPart.FeedData(stream);

            var element =
                new Drawing(
                  new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() { Cx = widthEmu, Cy = heightEmu },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent()
                    {
                        LeftEdge = 0L,
                        TopEdge = 0L,
                        RightEdge = 0L,
                        BottomEdge = 0L
                    },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties()
                    {
                        Id = (UInt32Value)1U,
                        Name = Path.GetFileName(imagePath)
                    },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties(
                        new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks() { NoChangeAspect = true }
                    ),
                    new DocumentFormat.OpenXml.Drawing.Graphic(
                        new DocumentFormat.OpenXml.Drawing.GraphicData(
                            new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties()
                                    {
                                        Id = (UInt32Value)0U,
                                        Name = Path.GetFileName(imagePath)
                                    },
                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()),
                                new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                    new DocumentFormat.OpenXml.Drawing.Blip()
                                    {
                                        Embed = mainPart.GetIdOfPart(imgPart),
                                        CompressionState = DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print
                                    },
                                    new DocumentFormat.OpenXml.Drawing.Stretch(
                                        new DocumentFormat.OpenXml.Drawing.FillRectangle())),
                                new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                    new DocumentFormat.OpenXml.Drawing.Transform2D(
                                        new DocumentFormat.OpenXml.Drawing.Offset() { X = 0L, Y = 0L },
                                        new DocumentFormat.OpenXml.Drawing.Extents() { Cx = widthEmu, Cy = heightEmu }),
                                    new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                                        new DocumentFormat.OpenXml.Drawing.AdjustValueList()
                                    )
                                    { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle })
                            )
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                    )
                  )
                );

            var para = new Paragraph(new Run(element));
            mainPart.Document.Body.Append(para);
        }

        private static PartTypeInfo GuessImagePartType(string path)
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();
            return ext switch
            {
                ".png" => ImagePartType.Png,     // <-- επιστρέφει PartTypeInfo
                ".gif" => ImagePartType.Gif,
                ".bmp" => ImagePartType.Bmp,
                ".tiff" => ImagePartType.Tiff,
                _ => ImagePartType.Jpeg
            };
        }
    }
}
