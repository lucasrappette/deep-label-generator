using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Reflection;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.VisualBasic.FileIO;
using QRCoder;
using QualityScoringBlazor.Models;
using Spire.Xls;
using System.Web;
using Font = System.Drawing.Font;
using Image = System.Drawing.Image;
using Rectangle = System.Drawing.Rectangle;

namespace QualityScoringBlazor.Services
{
    public class LabelService : ILabelService
    {
        private readonly string _pathBase = AppDomain.CurrentDomain.BaseDirectory;

        public bool GenerateLabels()
        {
            
            var complete = false;

            var pathProcess= Path.Combine(_pathBase, "Process");
            try
            {
                var pathUploads = Path.Combine(_pathBase, "Uploads");
                var pathDownloads = Path.Combine(_pathBase, "Downloads");
                var pathBaks = Path.Combine(_pathBase, "Baks");
                var pathPdfs = Path.Combine(_pathBase, "Pdfs");
                var xlsxPath = Path.Combine(pathUploads, "xlsx route version.xlsx");
                var csvPath = Path.Combine(_pathBase, "Process" + "\\CSVRouteVersion.csv");
                var dirPath = Path.Combine(_pathBase, "Pdfs");
                var finalSheet = "LabelSheet15.png";
                try
                {
                    ConvertXlsxToCsv
                    (
                        xlsxPath,
                        csvPath
                    );
                }
                catch (Exception e)
                {
                    Console.WriteLine
                    (
                        e
                    );
                    throw;
                }

                //GenerateQrCode();
                List<Order>? orderList;
                var firstField = 2;

                try
                {
                    orderList = CreateOrderList(csvPath, firstField, null);
                }
                catch (Exception e)
                {
                    Console.WriteLine
                    (
                        e
                    );
                    throw;
                }

                var count = 0;
                var sheetCount = 100;
                foreach (var order in orderList)
                {
                    try
                    {
                        count++;
                        sheetCount++;
                        var barCodeText = order.StringOrderNumber;
                        var stopNumberText = order.ShortCode;
                        var dateText = order.StringDate;
                        var strArray = order.OrderAddress?.Split(',');
                        var addressTextLine1 = strArray?[0];
                        var addressTextLine2 = strArray?[1].Trim() + " " + strArray?[2];
                        var dateAddressText = $"{dateText}\r\n{addressTextLine1}\r\n{addressTextLine2}";
                        GenerateBarCode(barCodeText);
                        GenerateBackground();
                        GenerateBmpLabelSheetBackground();
                        GenerateBmpOrderNumber(barCodeText);
                        GenerateBmpStopNumber(stopNumberText, count);
                        GenerateBmpDateAddress(dateAddressText);
                        AddBarcode();
                        AddOrderNumber();
                        AddStopNumber();
                        AddDateAddress();
                        ResizeLabel();
                        AddLabels();
                        GeneratePdf(pathProcess, dirPath, finalSheet, sheetCount);
                        if (count > 5)
                            count = 1;
                    }
                    catch (Exception e)
                    {
                        continue;
                    }
                    
                }

                var pdfsArray = new string[sheetCount];
                for (var pdfs = 0; pdfs < sheetCount; pdfs++)
                {
                    
                    pdfsArray[pdfs] = (Path.Combine(_pathBase, "Pdfs" + "\\Labels" + $"{pdfs + 1}" + ".pdf"));
                }

                MergeMultiplePdf
                (
                    pathPdfs,
                    Path.Combine(_pathBase, "Downloads" + "\\Labels.pdf"),
                    Path.Combine(_pathBase, "Baks" + "\\Labels.pdf")

                );

                complete = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            return complete;
        }

        private static void GenerateQrCode()
        {
            var pathBase = AppDomain.CurrentDomain.BaseDirectory;
            var qrGenerator = new QRCodeGenerator();
            var qrCodeData = qrGenerator.CreateQrCode
            (
                @"https://deep.mmthriftapps.com/#/",
                QRCodeGenerator.ECCLevel.Q
            );
            var qrCode = new QRCode
            (
                qrCodeData
            );
            var qrCodeBmp = new BitmapByteQRCode
            (
                qrCodeData
            );
            //byte[] qrCodeImageBmp = qrCodeBmp.GetGraphic(20, new byte[] { 118, 126, 152 }, new byte[] { 144, 201, 111 });
            //byte[] fileContents = File.ReadAllBytes("test.txt");
            using var ms = new MemoryStream();
            using var bitmap = qrCode.GetGraphic
            (
                20
            );
            bitmap.Save
            (
                ms,
                ImageFormat.Png
            );

            bitmap.Save(Path.Combine(pathBase, "Process" + "\\QRCodeZen.png"));
        }

        public static void ConvertXlsxToCsv(string xlsxPath, string csvPath)
        {
            if (!File.Exists(xlsxPath)) throw new FileNotFoundException(xlsxPath);
            if (File.Exists(csvPath))
            {
                File.Delete(csvPath);
            };
            try
            {
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(xlsxPath);
                if (workbook.Worksheets.Count > 1)
                {
                    for (int i = 0; i < workbook.Worksheets.Count; i++)
                    {
                        Worksheet sheet = workbook.Worksheets[i];
                        sheet.SaveToFile(csvPath + i + ".csv", ",", Encoding.UTF8);
                    }
                }
                else
                {
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.SaveToFile(csvPath, ",", Encoding.UTF8);
                }
            }
            catch (Exception e)
            {

                throw e;
            }
            return;
        }

        private static List<Order>? CreateOrderList(string csvPath, int firstField, List<Order>? orderList)
        {
            orderList ??= new List<Order>();
            ;
            try
            {
                using TextFieldParser csvParser = new TextFieldParser
                (
                    csvPath
                );
                csvParser.TextFieldType = FieldType.Delimited;
                csvParser.SetDelimiters
                (
                    ","
                );
                // Skip the row with the column names
                csvParser.ReadLine();
                while (!csvParser.EndOfData)
                {
                    try
                    {
                        //Process row
                        string?[]? fields = csvParser.ReadFields();
                        char[] orderNumberArray = fields?[firstField].ToArray();
                        if (orderNumberArray is { Length: < 8 })
                        {
                            firstField = firstField + 1;
                        }
                        //if (orderNumberArray is { Length: > 8 })
                        //{
                        //    firstField = firstField - 1;
                        //}
                        string? orderNumber = fields?[firstField];
                        string? orderDestination = fields?[firstField + 1];
                        Order order = new Order()
                        {
                            StringOrderNumber = orderNumber,
                            OrderAddress = orderDestination,
                            OrderNoCharArray = orderNumber?.ToCharArray(),
                            ShortCode = orderNumber?.ToCharArray()[10].ToString()
                                        + orderNumber?.ToCharArray()[11].ToString()
                                        + orderNumber?.ToCharArray()[12].ToString(),
                            StringDate = orderNumber?.ToCharArray()[2].ToString()
                                         + orderNumber?.ToCharArray()[3].ToString()
                                         + "\\"
                                         + orderNumber?.ToCharArray()[4].ToString()
                                         + orderNumber?.ToCharArray()[5].ToString()
                                         + "\\"
                                         + orderNumber?.ToCharArray()[0].ToString()
                                         + orderNumber?.ToCharArray()[1].ToString()
                        };
                        orderList.Add
                        (
                            order
                        );
                    }
                    catch (Exception e)
                    {
                        firstField--;
                        continue;
                    }
                   
                }

                return orderList;
            }
            catch (Exception e)
            {
                CreateOrderList
                (
                    csvPath,
                    firstField: 1,
                    orderList
                );
            }
            return orderList;
        }

        private static void CreateZenQrCode()
        {
            string pathBase = AppDomain.CurrentDomain.BaseDirectory;
            Zen.Barcode.CodeQrBarcodeDraw qrCodeZen = Zen.Barcode.BarcodeDrawFactory.CodeQr;
            var imageQrCode = qrCodeZen.Draw
            (
                "123456789",
                100
            );

            imageQrCode.Save(Path.Combine(pathBase, "Process" + "\\QRCodeZen.png"));
        }

        private static void GenerateBarCode(string? barCodeText)
        {
            var pathBase = AppDomain.CurrentDomain.BaseDirectory;
            if (barCodeText == null)
                throw new ArgumentNullException
                (
                    nameof(barCodeText)
                );
            Zen.Barcode.Code128BarcodeDraw barCode = Zen.Barcode.BarcodeDrawFactory.Code128WithChecksum;

            var imageBarcode = barCode.Draw
            (
                barCodeText,
                750,
                15
            );

            imageBarcode.Save(Path.Combine(pathBase, "Process" + "\\BarcodeZen.png"));
        }

        private static Bitmap GenerateBackground()
        {
            var pathBase = AppDomain.CurrentDomain.BaseDirectory;
            var bmpLabelSheet = new Bitmap
            (
                2625,
                2000
            );
            using var graphics = Graphics.FromImage
            (
                bmpLabelSheet
            );
            var font = new Font
            (
                "Verdana",
                155
            );
            graphics.FillRectangle
            (
                new SolidBrush
                (
                    Color.White
                ),
                0,
                0,
                bmpLabelSheet.Width,
                bmpLabelSheet.Height
            );
            graphics.Flush();
            font.Dispose();
            graphics.Dispose();

            bmpLabelSheet.Save(Path.Combine(pathBase, "Process" + "\\Background.png"));

            return bmpLabelSheet;
        }

        private static Bitmap GenerateBmpLabelSheetBackground()
        {
            var pathBase = AppDomain.CurrentDomain.BaseDirectory;
            var bmpLabelSheet = new Bitmap
            (
                2550,
                3300
            );
            using var graphics = Graphics.FromImage
            (
                bmpLabelSheet
            );
            var font = new Font
            (
                "Verdana",
                155
            );
            graphics.FillRectangle
            (
                new SolidBrush
                (
                    Color.White
                ),
                0,
                0,
                bmpLabelSheet.Width,
                bmpLabelSheet.Height
            );
            graphics.Flush();
            font.Dispose();
            graphics.Dispose();

            bmpLabelSheet.Save(Path.Combine(pathBase, "Process" + "\\LabelSheetBackground.png"));

            return bmpLabelSheet;
        }

        private static Bitmap GenerateBmpOrderNumber(string? barCodeText)
        {
            var pathBase = AppDomain.CurrentDomain.BaseDirectory;
            var bmpOrderNumber = new Bitmap
            (
                3000,
                300
            );
            using var graphics = Graphics.FromImage
            (
                bmpOrderNumber
            );
            var font = new Font
            (
                "Arial Black",
                135
            );
            graphics.FillRectangle
            (
                new SolidBrush
                (
                    Color.White
                ),
                0,
                0,
                bmpOrderNumber.Width,
                bmpOrderNumber.Height
            );
            graphics.DrawString
            (
                barCodeText,
                font,
                new SolidBrush
                (
                    Color.Black
                ),
                5,
                5
            );
            graphics.Flush();
            font.Dispose();
            graphics.Dispose();

            bmpOrderNumber.Save(Path.Combine(pathBase, "Process" + "\\OrderNumber.png"));

            return bmpOrderNumber;
        }

        private static Bitmap GenerateBmpStopNumber(string? stopNumberText, int count)
        {
            var pathBase = AppDomain.CurrentDomain.BaseDirectory;
            var currentColor = count switch
            {
                1 => Color.Red,
                2 => Color.Purple,
                3 => Color.Blue,
                4 => Color.Orange,
                5 => Color.Green,
                _ => Color.Red
            };
            var bmpStopNumber = new Bitmap
            (
                1200,
                900
            );
            using var graphics = Graphics.FromImage
            (
                bmpStopNumber
            );
            var font = new Font
            (
                "Arial Black",
                380
            );
            graphics.FillRectangle
            (
                new SolidBrush
                (
                    Color.White
                ),
                0,
                0,
                bmpStopNumber.Width,
                bmpStopNumber.Height
            );
            graphics.DrawString
            (
                stopNumberText,
                font,
                new SolidBrush
                (
                    currentColor
                ),
                0,
                50
            );
            graphics.Flush();
            font.Dispose();
            graphics.Dispose();

            bmpStopNumber.Save(Path.Combine(pathBase, "Process" + "\\StopNumber.png"));

            return bmpStopNumber;
        }
        
        private static Bitmap GenerateBmpDateAddress(string? addressText)
        {
            var pathBase = AppDomain.CurrentDomain.BaseDirectory;
            var bmpDateAddress = new Bitmap
            (
                1200,
                900
            );

            using var graphics = Graphics.FromImage
            (
                bmpDateAddress
            );
            var font = new Font
            (
                "Arial Black",
                60
            );
            graphics.FillRectangle
            (
                new SolidBrush
                (
                    Color.White
                ),
                0,
                0,
                bmpDateAddress.Width,
                bmpDateAddress.Height
            );
            graphics.DrawString
            (
                addressText,
                font,
                new SolidBrush
                (
                    Color.Black
                ),
                0,
                50
            );
            graphics.Flush();
            font.Dispose();
            graphics.Dispose();
            bmpDateAddress.Save(Path.Combine(pathBase, "Process" + "\\DateAddress.png"));

            return bmpDateAddress;
        }

        private static void AddBarcode()
        {
            var pathBase = AppDomain.CurrentDomain.BaseDirectory;
            using var background = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\Background.png")
            );
            using var barcode = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\BarcodeZen.png"
                )
            );
            using var graphics = Graphics.FromImage
            (
                background
            );
            var x = 120;
            var y = 450;
            graphics.DrawImage
            (
                barcode,
                x,
                y
            );

            background.Save(Path.Combine(pathBase, "Process" + "\\BarcodeResult.png"), ImageFormat.Png);
        }

        private static void AddOrderNumber()
        {
            var pathBase = AppDomain.CurrentDomain.BaseDirectory;
            using var background = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\BarcodeResult.png"
                )
            );
            using var orderNumber = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\OrderNumber.png"
                )
            );
            using var graphics = Graphics.FromImage
            (
                background
            );
            var x = 80;
            var y = 150;
            graphics.DrawImage
            (
                orderNumber,
                x,
                y
            );
            background.Save(Path.Combine(pathBase, "Process" + "\\OrderNumberResult.png"), ImageFormat.Png);
        }

        private static void AddStopNumber()
        {
            var pathBase = AppDomain.CurrentDomain.BaseDirectory;
            for (var i = 1; i < 16; i++)
            {
                using var background = Image.FromFile
                (
                    Path.Combine(pathBase, "Process" + "\\OrderNumberResult.png")
                );
                using var stopNumber = Image.FromFile
                (
                    Path.Combine(pathBase, "Process" + "\\StopNumber" + $"{i}" + ".png")
                );
                using var graphics = Graphics.FromImage
                (
                    background
                );
                var x = 35;
                var y = 1200;
                graphics.DrawImage
                (
                    stopNumber,
                    x,
                    y
                );
                background.Save(Path.Combine(pathBase, "Process" + "\\StopNumberResult" + $"{i}" + ".png"), ImageFormat.Png);
            }
        }

        private static void AddDateAddress()
        {
            var pathBase = AppDomain.CurrentDomain.BaseDirectory;
            for (var i = 1; i < 16; i++)
            {
                using var background = Image.FromFile
                (
                    Path.Combine(pathBase, pathBase, "Process" + "\\StopNumberResult" + $"{i}" + ".png"
                    )
                );
                using var addressPNG = Image.FromFile
                (
                    Path.Combine(pathBase, "Process" + "\\DateAddress.png"
                    )
                );
                using var graphics = Graphics.FromImage
                (
                    background
                );
                var x = 1250;
                var y = 1350;
                graphics.DrawImage
                (
                    addressPNG,
                    x,
                    y
                );
                background.Save(Path.Combine(pathBase, "Process" + "\\AddressResult" + $"{i}" + ".png"), ImageFormat.Png);
                background.Save(Path.Combine(pathBase, "Process" + "\\FinalLabelResult" + $"{i}" + ".png"), ImageFormat.Png);
            }
        }

        private static void ResizeLabel()
        {
            var pathBase = AppDomain.CurrentDomain.BaseDirectory;
            for (var i = 1; i < 16; i++)
            {
                using var roughLabelResult = Image.FromFile
                (
                    Path.Combine(pathBase, "Process" + "\\FinalLabelResult" + $"{i}" + ".png")
                );
                var finalLabel = ResizeImage
                (
                    roughLabelResult,
                    750,
                    600
                );
                finalLabel.Save(Path.Combine(pathBase, "Process" + "\\FinalLabel" + $"{i}" + ".png"), ImageFormat.Png);
            }
        }

        private static void AddLabels()
        {
            AddLabelX1Y1();

            AddLabelX2Y1();

            AddLabelX3Y1();

            AddLabelX1Y2();

            AddLabelX2Y2();

            AddLabelX3Y2();

            AddLabelX1Y3();

            AddLabelX2Y3();

            AddLabelX3Y3();

            AddLabelX1Y4();

            AddLabelX2Y4();

            AddLabelX3Y4();

            AddLabelX1Y5();

            AddLabelX2Y5();

            AddLabelX3Y5();
        }

        private static void AddLabelX1Y1()
        {
            var pathBase = AppDomain.CurrentDomain.BaseDirectory;
            using var background = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\LabelSheetBackground.png")
            );
            using var label = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\FinalLabel1.png")
            );
            using var graphics = Graphics.FromImage
            (
                background
            );
            var x = 64;
            var y = 120;
            graphics.DrawImage
            (
                label,
                x,
                y
            );
            background.Save(Path.Combine(pathBase, "Process" + "\\LabelSheet1.png"), ImageFormat.Png);
        }

        private static void AddLabelX2Y1()
        {
            var pathBase = AppDomain.CurrentDomain.BaseDirectory;
            using var background = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\LabelSheet1.png")
            );
            using var label = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\FinalLabel2.png")
            );
            using var graphics = Graphics.FromImage
            (
                background
            );
            var x = 900;
            var y = 120;
            graphics.DrawImage
            (
                label,
                x,
                y
            );
            background.Save(Path.Combine(pathBase, "Process" + "\\LabelSheet2.png"), ImageFormat.Png);
        }

        private static void AddLabelX3Y1()
        {
            var pathBase = AppDomain.CurrentDomain.BaseDirectory;
            using var background = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\LabelSheet2.png")
            );
            using var label = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\FinalLabel3.png")
            );
            using var graphics = Graphics.FromImage
            (
                background
            );
            var x = 1740;
            var y = 120;
            graphics.DrawImage
            (
                label,
                x,
                y
            );
            background.Save(Path.Combine(pathBase, "Process" + "\\LabelSheet3.png"), ImageFormat.Png);
        }

        private static void AddLabelX1Y2()
        {
            var pathBase = AppDomain.CurrentDomain.BaseDirectory;
            using var background = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\LabelSheet3.png")
            );
            using var label = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\FinalLabel4.png")
            );
            using var graphics = Graphics.FromImage
            (
                background
            );
            var x = 64;
            var y = 740;
            graphics.DrawImage
            (
                label,
                x,
                y
            );
            background.Save(Path.Combine(pathBase, "Process" + "\\LabelSheet4.png"), ImageFormat.Png);
        }

        private static void AddLabelX2Y2()
        {
            var pathBase = AppDomain.CurrentDomain.BaseDirectory;
            using var background = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\LabelSheet4.png")
            );
            using var label = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\FinalLabel5.png")
            );
            using var graphics = Graphics.FromImage
            (
                background
            );
            var x = 900;
            var y = 740;
            graphics.DrawImage
            (
                label,
                x,
                y
            );
            background.Save(Path.Combine(pathBase, "Process" + "\\LabelSheet5.png"), ImageFormat.Png);
        }

        private static void AddLabelX3Y2()
        {
            var pathBase = AppDomain.CurrentDomain.BaseDirectory;
            using var background = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\LabelSheet5.png")
            );
            using Image label = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\FinalLabel6.png")
            );
            using var graphics = Graphics.FromImage
            (
                background
            );
            var x = 1740;
            var y = 740;
            graphics.DrawImage
            (
                label,
                x,
                y
            );
            background.Save(Path.Combine(pathBase, "Process" + "\\LabelSheet6.png"), ImageFormat.Png);
        }

        private static void AddLabelX1Y3()
        {
            var pathBase = AppDomain.CurrentDomain.BaseDirectory;
            using var background = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\LabelSheet6.png")
            );
            using var label = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\FinalLabel7.png")
            );
            using var graphics = Graphics.FromImage
            (
                background
            );
            var x = 64;
            var y = 1340;
            graphics.DrawImage
            (
                label,
                x,
                y
            );
            background.Save(Path.Combine(pathBase, "Process" + "\\LabelSheet7.png"), ImageFormat.Png);
        }

        private static void AddLabelX2Y3()
        {
            var pathBase = AppDomain.CurrentDomain.BaseDirectory;
            using var background = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\LabelSheet7.png")
            );
            using var label = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\FinalLabel8.png")
            );
            using var graphics = Graphics.FromImage
            (
                background
            );
            var x = 900;
            var y = 1340;
            graphics.DrawImage
            (
                label,
                x,
                y
            );
            background.Save(Path.Combine(pathBase, "Process" + "\\LabelSheet8.png"), ImageFormat.Png);
        }

        private static void AddLabelX3Y3()
        {
            var pathBase = AppDomain.CurrentDomain.BaseDirectory;
            using var background = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\LabelSheet8.png")
            );
            using var label = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\FinalLabel9.png")
            );
            using var graphics = Graphics.FromImage
            (
                background
            );
            var x = 1740;
            var y = 1340;
            graphics.DrawImage
            (
                label,
                x,
                y
            );
            background.Save(Path.Combine(pathBase, "Process" + "\\LabelSheet9.png"), ImageFormat.Png);
        }

        private static void AddLabelX1Y4()
        {
            var pathBase = AppDomain.CurrentDomain.BaseDirectory;
            using Image background = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\LabelSheet9.png"
                )
            );
            using Image label = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\FinalLabel10.png"
                )
            );
            using Graphics graphics = Graphics.FromImage
            (
                background
            );
            var x = 64;
            var y = 1940;
            graphics.DrawImage
            (
                label,
                x,
                y
            );
            background.Save(Path.Combine(pathBase, "Process" + "\\LabelSheet10.png"), ImageFormat.Png);
        }

        private static void AddLabelX2Y4()
        {
            var pathBase = AppDomain.CurrentDomain.BaseDirectory;
            using var background = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\LabelSheet10.png")
            );
            using var label = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\FinalLabel11.png")
            );
            using var graphics = Graphics.FromImage
            (
                background
            );
            int x = 900;
            int y = 1940;
            graphics.DrawImage
            (
                label,
                x,
                y
            );
            background.Save(Path.Combine(pathBase, "Process" + "\\LabelSheet11.png"), ImageFormat.Png);
        }

        private static void AddLabelX3Y4()
        {
            var pathBase = AppDomain.CurrentDomain.BaseDirectory;
            using var background = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\LabelSheet11.png")
            );
            using var label = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\FinalLabel12.png")
            );
            using var graphics = Graphics.FromImage
            (
                background
            );
            var x = 1740;
            var y = 1940;
            graphics.DrawImage
            (
                label,
                x,
                y
            );
            background.Save(Path.Combine(pathBase, "Process" + "\\LabelSheet12.png"), ImageFormat.Png);
        }

        private static void AddLabelX1Y5()
        {
            var pathBase = AppDomain.CurrentDomain.BaseDirectory;
            using var background = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\LabelSheet12.png")
            );
            using var label = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\FinalLabel13.png")
            );
            using var graphics = Graphics.FromImage
            (
                background
            );
            var x = 64;
            var y = 2560;
            graphics.DrawImage
            (
                label,
                x,
                y
            );
            background.Save(Path.Combine(pathBase, "Process" + "\\LabelSheet13.png"), ImageFormat.Png);
        }

        private static void AddLabelX2Y5()
        {
            var pathBase = AppDomain.CurrentDomain.BaseDirectory;
            using var background = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\LabelSheet13.png")
            );
            using var label = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\FinalLabel14.png")
            );
            using var graphics = Graphics.FromImage
            (
                background
            );
            var x = 900;
            var y = 2560;
            graphics.DrawImage
            (
                label,
                x,
                y
            );
            background.Save(Path.Combine(pathBase, "Process" + "\\LabelSheet14.png"), ImageFormat.Png);
        }

        private static void AddLabelX3Y5()
        {
            var pathBase = AppDomain.CurrentDomain.BaseDirectory;
            var path = AppDomain.CurrentDomain.BaseDirectory;
            var newPath = HttpUtility.UrlPathEncode("labelsheet15.png");
            using var background = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\LabelSheet14.png")
            );
            using var label = Image.FromFile
            (
                Path.Combine(pathBase, "Process" + "\\FinalLabel15.png")
            );
            using var graphics = Graphics.FromImage
            (
                background
            );
            var x = 1740;
            var y = 2560;
            graphics.DrawImage
            (
                label,
                x,
                y
            );
            background.Save(Path.Combine(pathBase, "Process" + "\\LabelSheet15.png"), ImageFormat.Png);
            background.Save(path + newPath, ImageFormat.Png);
        }

        public static Bitmap ResizeImage(Image image, int width, int height)
        {
            var destRect = new Rectangle(0, 0, width, height);
            var destImage = new Bitmap(width, height);

            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using var graphics = Graphics.FromImage(destImage);
            graphics.CompositingMode = CompositingMode.SourceCopy;
            graphics.CompositingQuality = CompositingQuality.HighQuality;
            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphics.SmoothingMode = SmoothingMode.HighQuality;
            graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

            using var wrapMode = new ImageAttributes();
            wrapMode.SetWrapMode(WrapMode.TileFlipXY);
            graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);

            return destImage;
        }

        private static bool GeneratePdf(string processPath, string dirPath, string finalSheet, int sheetCount)
        {
            var success = false;
            iTextSharp.text.Rectangle pageSize = null;
            var finalSheetPath = $"{processPath}" + "\\" + $"{finalSheet}";
            using var srcImage = new Bitmap(finalSheetPath);
            pageSize = new iTextSharp.text.Rectangle(0, 0, srcImage.Width, srcImage.Height);
            using var ms = new MemoryStream();
            var document = new iTextSharp.text.Document(pageSize, 0, 0, 0, 0);
            PdfWriter.GetInstance(document, ms).SetFullCompression();
            document.Open();
            var systemPath = System.IO.Path.GetDirectoryName
            (
                Assembly.GetExecutingAssembly().Location
            );
            string path = AppDomain.CurrentDomain.BaseDirectory;
            var newPath = HttpUtility.UrlPathEncode("labelsheet15.png");
            iTextSharp.text.Image? image;
            try
            {
                image = iTextSharp.text.Image.GetInstance(path + "Process\\" + newPath);
            }

            catch (Exception e)
            {
                Console.WriteLine
                (
                    e
                );
                throw;
            }
            document.Add(image);
            document.Close();
            var labelsPdfPath = $"{dirPath}" + "\\"+ "Labels" + $"{sheetCount}" + ".pdf";
            File.WriteAllBytes(labelsPdfPath, ms.ToArray());
            success = true;

            return success;
        }

        public static void MergeMultiplePdf(string[] pdfFileNames, string outputFile)
        {
            var pdfDoc = new Document();
            using var myFileStream = new FileStream(outputFile, FileMode.Create);
            var pdfWriter = new PdfCopy(pdfDoc, myFileStream);
            pdfDoc.Open();
            foreach (var fileName in pdfFileNames)
            {  
                var pdfReader = new PdfReader(fileName);
                pdfReader.ConsolidateNamedDestinations();
                for (var i = 1; i <= pdfReader.NumberOfPages; i++)
                {
                    var page = pdfWriter.GetImportedPage(pdfReader, i);
                    pdfWriter.AddPage(page);
                }
                var form = pdfReader.AcroForm;
                if (form != null)
                {
                    pdfWriter.CopyAcroForm(pdfReader);
                } 
                pdfReader.Close();
            }  
            pdfWriter.Close();
            pdfDoc.Close();
        }

        private static void MergeMultiplePdf(string sourceFolder, string destinationFile, string backupFile)
        {
            using var stream = new MemoryStream();
            using (var doc = new Document())
            {
                var pdf = new PdfCopy(doc, stream);
                pdf.CloseStream = false;
                doc.Open();

                foreach (var file in Directory.GetFiles(sourceFolder))
                {
                    var reader = new PdfReader(file);
                    for (var i = 0; i < reader.NumberOfPages; i++)
                    {
                        var page = pdf.GetImportedPage(reader, i + 1);
                        pdf.AddPage(page);
                    }

                    pdf.FreeReader(reader);
                    reader.Close();
                }
            }
            using (var streamX = new FileStream(destinationFile, FileMode.Create))
            {
                stream.WriteTo(streamX);
            }
            using (var streamX = new FileStream(backupFile, FileMode.Create))
            {
                stream.WriteTo(streamX);
            }
        }

    }
}
