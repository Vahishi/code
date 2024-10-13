using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using IronPdf;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xceed.Words.NET;
using Xceed.Document.NET;
using System.Drawing;
using Border = OfficeOpenXml.Style.Border; // For DocX

namespace ConvertToPdfAndExcel.Pages
{
    public class ConvertPageModel : PageModel
    {
        public PurchaseOrderViewModel PurchaseOrder { get; set; }

        public void OnGet()
        {
            // Initialize the PurchaseOrder when the page is first loaded (optional)
            PurchaseOrder = LoadPurchaseOrder();
        }

        public IActionResult OnGetExportPdf()
        {
            if (PurchaseOrder == null)
            {
                PurchaseOrder = LoadPurchaseOrder();
            }

            var renderer = new HtmlToPdf();
            string htmlContent = RenderPurchaseOrderHtml(PurchaseOrder); // Generate HTML content from ViewModel
            var pdf = renderer.RenderHtmlAsPdf(htmlContent);
            var pdfBytes = pdf.BinaryData;
            return File(pdfBytes, "application/pdf", $"{PurchaseOrder.PONumber}.pdf");
        }

        public IActionResult OnGetExportExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            if (PurchaseOrder == null)
            {
                PurchaseOrder = LoadPurchaseOrder();
            }

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Purchase Order");

                // Set header row for the title
                worksheet.Cells[1, 1].Value = "PURCHASE ORDER";
                worksheet.Cells[1, 1, 1, 10].Merge = true;
                worksheet.Cells[1, 1].Style.Font.Bold = true;
                worksheet.Cells[1, 1].Style.Font.Size = 20;
                worksheet.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                // Add "To" section and other header information
                worksheet.Cells[3, 1].Value = $"To: {PurchaseOrder.CompanyName ?? "N/A"}";
                worksheet.Cells[4, 1].Value = $"Attention: {PurchaseOrder.Attention ?? "N/A"}";
                worksheet.Cells[5, 1].Value = $"P.O. No: {PurchaseOrder.PONumber ?? "N/A"}";
                worksheet.Cells[6, 1].Value = $"P.O. Date: {PurchaseOrder.PODate:dd-MMM-yyyy}";
                worksheet.Cells[7, 1].Value = $"Project: {PurchaseOrder.Project ?? "N/A"}";
                worksheet.Cells[8, 1].Value = $"VAT No: {PurchaseOrder.VATNo ?? "N/A"}";

                // Set column widths for better readability
                worksheet.Column(1).Width = 5;   // #
                worksheet.Column(2).Width = 15;  // Stock No.
                worksheet.Column(3).Width = 50;  // Description & Specifications (wider column to fit description)
                worksheet.Column(4).Width = 10;  // Unit
                worksheet.Column(5).Width = 10;  // Qty
                worksheet.Column(6).Width = 15;  // Unit Price
                worksheet.Column(7).Width = 15;  // Discount
                worksheet.Column(8).Width = 10;  // VAT%
                worksheet.Column(9).Width = 20;  // Amount Before Tax
                worksheet.Column(10).Width = 20; // VAT Amount

                // Enable wrapping for "Description & Specifications" column
                worksheet.Cells["C:C"].Style.WrapText = true; // Apply to the entire "Description & Specifications" column

                // Set column headers for the item details table
                var headers = new[] { "#", "Stock No.", "Description & Specifications", "Unit", "Qty", "Unit Price", "Discount", "VAT%", "Amount Before Tax", "VAT Amount" };
                for (int i = 0; i < headers.Length; i++)
                {
                    worksheet.Cells[10, i + 1].Value = headers[i];
                    worksheet.Cells[10, i + 1].Style.Font.Bold = true;
                    worksheet.Cells[10, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[10, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells[10, i + 1].Style.WrapText = true; // Enable text wrapping for headers
                }

                // Add data rows for items
                for (int i = 0; i < PurchaseOrder.Items.Count; i++)
                {
                    var item = PurchaseOrder.Items[i];
                    worksheet.Cells[i + 11, 1].Value = i + 1; // Serial number
                    worksheet.Cells[i + 11, 2].Value = item.StockNumber;
                    worksheet.Cells[i + 11, 3].Value = item.Description; // Wrap description text
                    worksheet.Cells[i + 11, 4].Value = item.Unit;
                    worksheet.Cells[i + 11, 5].Value = item.Quantity;
                    worksheet.Cells[i + 11, 6].Value = item.UnitPrice;
                    worksheet.Cells[i + 11, 7].Value = item.Discount;
                    worksheet.Cells[i + 11, 8].Value = item.VAT;
                    worksheet.Cells[i + 11, 9].Value = item.AmountBeforeTax;
                    worksheet.Cells[i + 11, 10].Value = item.VATAmount;

                    // Apply border for the row and enable text wrapping
                    for (int j = 1; j <= 10; j++)
                    {
                        worksheet.Cells[i + 11, j].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        worksheet.Cells[i + 11, j].Style.WrapText = true; // Enable wrapping for all item rows
                    }
                }

                // Total calculations
                int startRow = PurchaseOrder.Items.Count + 12; // Total rows start after items
                worksheet.Cells[startRow, 8].Value = "Total Amount:";
                worksheet.Cells[startRow, 9].Value = PurchaseOrder.TotalAmount;

                worksheet.Cells[startRow + 1, 8].Value = "Less: Discount:";
                worksheet.Cells[startRow + 1, 9].Value = PurchaseOrder.Discount;

                worksheet.Cells[startRow + 2, 8].Value = "Gross Total:";
                worksheet.Cells[startRow + 2, 9].Value = PurchaseOrder.TotalAmount - PurchaseOrder.Discount;

                worksheet.Cells[startRow + 3, 8].Value = "VAT Amount:";
                worksheet.Cells[startRow + 3, 9].Value = PurchaseOrder.VATAmount;

                worksheet.Cells[startRow + 4, 8].Value = "Net Amount:";
                worksheet.Cells[startRow + 4, 9].Value = PurchaseOrder.NetAmount;

                worksheet.Cells[startRow + 5, 1].Value = $"Total in USD: {PurchaseOrder.TotalAmount * 0.27m}";

                // Format total rows
                for (int i = startRow; i < startRow + 5; i++)
                {
                    worksheet.Cells[i, 8].Style.Font.Bold = true;
                    worksheet.Cells[i, 8].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells[i, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                }

                // Amount in Words
                worksheet.Cells[startRow + 6, 1].Value = $"Amount in Words: {PurchaseOrder.AmountInWords}";
                worksheet.Cells[startRow + 6, 1].Style.Font.Bold = true;
                worksheet.Cells[startRow + 6, 1, startRow + 6, 10].Merge = true; // Merge cells for amount in words

                var excelBytes = package.GetAsByteArray();
                return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"{PurchaseOrder.PONumber}.xlsx");
            }
        }


        public IActionResult OnGetExportDoc()
        {
            if (PurchaseOrder == null)
            {
                PurchaseOrder = LoadPurchaseOrder();
            }

            using (var stream = new MemoryStream())
            {
                using (var document = DocX.Create(stream))
                {
                    // Set margins for the content area
                    document.MarginTop = 25f;  // Margin for the top content area
                    document.MarginBottom = 25f; // Margin for the bottom content area to allow space for footer
                    document.MarginLeft = 50f;
                    document.MarginRight = 50f;

                    // Define header and footer image paths
                    string headerImagePath = "wwwroot/images/logo.jpg"; // Header image path
                    string footerImagePath = "wwwroot/images/footer.jpg"; // Footer image path

                    // Add header image at the top with zero margin
                    var headerImage = document.AddImage(headerImagePath);
                    var headerPicture = headerImage.CreatePicture(70, 500); // Adjust size as necessary
                    var headerParagraph = document.InsertParagraph();
                    headerParagraph.AppendPicture(headerPicture).Alignment = Alignment.center;

                    // Title: PURCHASE ORDER
                    var titleFormat = new Formatting { Size = 20D, Bold = true, FontColor = Color.Blue };
                    document.InsertParagraph("PURCHASE ORDER", false, titleFormat).Alignment = Alignment.center;
                    document.InsertParagraph().SpacingAfter(10); // Space after title

                    // Info box (To and P.O. Information)
                    var toPoTable = document.AddTable(1, 2);
                    toPoTable.SetWidths(new float[] { 300, 300 });
                    toPoTable.Design = TableDesign.TableGrid;
                    toPoTable.Rows[0].Cells[0].Paragraphs[0].Append($"To: {PurchaseOrder.CompanyName ?? "N/A"}");
                    toPoTable.Rows[0].Cells[1].Paragraphs[0].Append($"P.O. No: {PurchaseOrder.PONumber ?? "N/A"}\n" +
                                                                    $"P.O. Date: {PurchaseOrder.PODate:dd-MMM-yyyy}\n" +
                                                                    $"VAT No: {PurchaseOrder.VATNo ?? "N/A"}\n" +
                                                                    $"Project: {PurchaseOrder.Project ?? "N/A"}").Alignment = Alignment.right;

                    document.InsertTable(toPoTable);
                    document.InsertParagraph().SpacingAfter(5); // Space after info box

                    // Attention and Reference
                    var normalFormat = new Formatting { Size = 12D };
                    document.InsertParagraph($"Attention: {PurchaseOrder.Attention ?? "N/A"}", false, normalFormat).SpacingAfter(5);
                    document.InsertParagraph($"Reference: {PurchaseOrder.Reference ?? "N/A"}", false, normalFormat).SpacingAfter(5);
                    document.InsertParagraph("Subject: Purchase Order for Materials", false, normalFormat).SpacingAfter(10);

                    // Horizontal Line
                    document.InsertParagraph().InsertHorizontalLine(HorizontalBorderPosition.bottom, BorderStyle.Tcbs_single, 6, 1, Color.Black);
                    document.InsertParagraph().SpacingAfter(5); // Space after line

                    // Dear Sir section
                    document.InsertParagraph("Dear Sir,", false, normalFormat).SpacingAfter(5);
                    document.InsertParagraph("With reference to your quotation, please be informed that we accepted your terms & conditions with the below prices as mentioned:", false, normalFormat).SpacingAfter(10);

                    // Create table for item details with compact layout and borders
                    var itemTable = document.AddTable(PurchaseOrder.Items.Count + 1, 10);
                    itemTable.SetWidths(new float[] { 20, 80, 180, 50, 50, 70, 70, 50, 80, 80 });
                    itemTable.Design = TableDesign.TableGrid; // Add borders to the table

                    // Set Table Headers
                    itemTable.Rows[0].Cells[0].Paragraphs[0].Append("#").Bold();
                    itemTable.Rows[0].Cells[1].Paragraphs[0].Append("Stock No.").Bold();
                    itemTable.Rows[0].Cells[2].Paragraphs[0].Append("Description & Specifications").Bold();
                    itemTable.Rows[0].Cells[3].Paragraphs[0].Append("Unit").Bold();
                    itemTable.Rows[0].Cells[4].Paragraphs[0].Append("Qty").Bold();
                    itemTable.Rows[0].Cells[5].Paragraphs[0].Append("Unit Price").Bold();
                    itemTable.Rows[0].Cells[6].Paragraphs[0].Append("Discount").Bold();
                    itemTable.Rows[0].Cells[7].Paragraphs[0].Append("VAT%").Bold();
                    itemTable.Rows[0].Cells[8].Paragraphs[0].Append("Amount Before Tax").Bold();
                    itemTable.Rows[0].Cells[9].Paragraphs[0].Append("VAT Amount").Bold();

                    // Fill Table Data
                    for (int i = 0; i < PurchaseOrder.Items.Count; i++)
                    {
                        var item = PurchaseOrder.Items[i];
                        itemTable.Rows[i + 1].Cells[0].Paragraphs[0].Append((i + 1).ToString()).Alignment = Alignment.center;
                        itemTable.Rows[i + 1].Cells[1].Paragraphs[0].Append(item.StockNumber ?? "N/A").Alignment = Alignment.center;
                        itemTable.Rows[i + 1].Cells[2].Paragraphs[0].Append(item.Description ?? "N/A").Alignment = Alignment.left;
                        itemTable.Rows[i + 1].Cells[3].Paragraphs[0].Append(item.Unit ?? "N/A").Alignment = Alignment.center;
                        itemTable.Rows[i + 1].Cells[4].Paragraphs[0].Append(item.Quantity.ToString()).Alignment = Alignment.center;
                        itemTable.Rows[i + 1].Cells[5].Paragraphs[0].Append(item.UnitPrice.ToString("F2")).Alignment = Alignment.right;
                        itemTable.Rows[i + 1].Cells[6].Paragraphs[0].Append(item.Discount.ToString("F2")).Alignment = Alignment.right;
                        itemTable.Rows[i + 1].Cells[7].Paragraphs[0].Append(item.VAT.ToString("F2")).Alignment = Alignment.right;
                        itemTable.Rows[i + 1].Cells[8].Paragraphs[0].Append(item.AmountBeforeTax.ToString("F2")).Alignment = Alignment.right;
                        itemTable.Rows[i + 1].Cells[9].Paragraphs[0].Append(item.VATAmount.ToString("F2")).Alignment = Alignment.right;
                    }

                    document.InsertTable(itemTable);
                    document.InsertParagraph().SpacingAfter(10); // Space after table

                    // Total calculations
                    document.InsertParagraph($"Total Amount: {PurchaseOrder.TotalAmount:F2}", false, normalFormat).Alignment = Alignment.right;
                    document.InsertParagraph($"Less: Discount: {PurchaseOrder.Discount:F2}", false, normalFormat).Alignment = Alignment.right;
                    document.InsertParagraph($"Gross Total: {(PurchaseOrder.TotalAmount - PurchaseOrder.Discount):F2}", false, normalFormat).Alignment = Alignment.right;
                    document.InsertParagraph($"VAT Amount: {PurchaseOrder.VATAmount:F2}", false, normalFormat).Alignment = Alignment.right;
                    document.InsertParagraph($"Net Amount: {PurchaseOrder.NetAmount:F2}", false, new Formatting { Size = 12D, Bold = true }).Alignment = Alignment.right;

                    document.InsertParagraph().SpacingAfter(10);
                    // Amount in Words section
                    var wordsFormat = new Formatting { Size = 14D, FontColor = Color.Blue, Bold = true };
                    document.InsertParagraph($"Amount in Words: {PurchaseOrder.AmountInWords}", false, wordsFormat).Alignment = Alignment.center;
                    document.InsertParagraph().SpacingAfter(10); // Space after amount in words

                    document.InsertParagraph($"Total in USD: {PurchaseOrder.TotalAmount * 0.27m}", false, new Formatting { Size = 12D, Bold = true }).Alignment = Alignment.right;
                    document.InsertParagraph().SpacingAfter(10);
                    // Final message
                    document.InsertParagraph("This Purchase Order is raised to order the items as listed and subject to provision attached/listed below.", false, normalFormat).SpacingAfter(5);
                    document.InsertParagraph("Please return a copy of this Purchase Order duly signed & stamped in acknowledgement of your receipt.", false, normalFormat).SpacingAfter(10);
                    document.InsertParagraph("Thanks and Regards,", false, normalFormat).SpacingAfter(1);

                    // Footer image at the bottom with no margin
                    var footerImage = document.AddImage(footerImagePath);
                    var footerPicture = footerImage.CreatePicture(70, 500); // Adjust size as necessary
                    var footerParagraph = document.InsertParagraph();
                    footerParagraph.AppendPicture(footerPicture).Alignment = Alignment.center;

                    // Save the document to the stream
                    document.Save();
                    stream.Position = 0;

                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"{PurchaseOrder.PONumber}.docx");
                }
            }
        }




































        // Mocked method to simulate loading the purchase order data
        private PurchaseOrderViewModel LoadPurchaseOrder()
        {
            return new PurchaseOrderViewModel
            {
                PONumber = "RAQ-PO-2024-00219",
                PODate = new DateTime(2024, 9, 11),
                VATNo = "310709107100003",
                Project = "Acorn Distributions",
                Attention = "Mr. Elan",
                Reference = "4500141571",
                Subject = "Purchase Order for Materials",
                Items = new List<PurchaseOrderItem>
                {
                    new PurchaseOrderItem
                    {
                        StockNumber = "GEN-02009",
                        Description = "SYRINGE ASEPTO 50 CC STERILE, PART N0: AS011P",
                        Unit = "BOX",
                        Quantity = 22,
                        UnitPrice = 60.00M,
                        Discount = 0.00M,
                        AmountBeforeTax = 1320.00M,
                        VAT = 0.00M,
                        VATAmount = 0.00M
                    }
                },
                TotalAmount = 1320.00M,
                Discount = 0.00M,
                VATAmount = 0.00M,
                NetAmount = 1320.00M,
                CompanyName = "ACORN DISTRIBUTIONS",
                Address = "UK",
                ContactPerson = "Mr. Elan",
                AmountInWords = "One Thousand Three Hundred Twenty Saudi Riyals only"
            };
        }

        // Method to render the purchase order as HTML (if needed for PDF export)
        private string RenderPurchaseOrderHtml(PurchaseOrderViewModel po)
        {
            string headerImagePath = "wwwroot/images/logo.jpg"; // Path to your header image
            string footerImagePath = "wwwroot/images/footer.jpg"; // Path to your footer image

            return $@"
<html>
<head>
    <style>
        @page {{
            size: A4;
            margin: 0; /* Set margins to 0 for header and footer */
        }}
        body {{
            font-family: 'Arial', sans-serif;
            margin: 0; /* Remove body margin */
            box-sizing: border-box;
        }}
        .header {{
            position: fixed;
            top: 0; /* Align header at the very top of the page */
            left: 0;
            right: 0;
            height: 100px; /* Height of the header */
            text-align: center;
        }}
        .footer {{
            position: fixed;
            bottom: 0; /* Align footer at the very bottom of the page */
            left: 0;
            right: 0;
            height: 100px; /* Height of the footer */
            text-align: center;
        }}
        .content {{
            margin: 120px 25px 120px 25px; /* Space for header, footer, and content */
            font-size: 14px; /* Set the font size */
            overflow: hidden; /* To ensure content stays within the page */
        }}
        .title {{
            color: #2E8BC0; /* Set title color */
            text-align: center;
            font-size: 28px; /* Adjust title font size */
            font-weight: bold; /* Make title bold */
            margin-top: 10px; /* Space above title */
        }}
        .info-box {{
            display: flex;
            justify-content: space-between;
            margin: 10px 0;
        }}
        .info-left, .info-right {{
            width: 48%;
            border: 1px solid black;
            padding: 5px;
        }}
        .info-left p, .info-right p {{
            margin: 5px 0; /* Space between paragraphs */
        }}
        .horizontal-line {{
            border-top: 2px solid black;
            margin: 10px 0;
        }}
        .dear-sir {{
            margin-top: 10px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 10px;
        }}
        table, th, td {{
            border: 1px solid black;
        }}
        th {{
            background-color: #87CEEB; /* Sky blue background color for the header */
            padding: 8px;
        }}
        td {{
            padding: 6px;
            text-align: left;
        }}
        .amount-in-words {{
            color: blue; /* Set the text color of the amount in words */
            font-weight: bold; /* Make it bold for emphasis */
            text-align: center; /* Center align */
            font-size: 14px; /* Increase font size */
            margin-top: 10px; /* Space above */
        }}
        .total-usd {{
            margin-top: 10px; /* Space above USD total */
            text-align: right; /* Align to the right */
        }}
        .nothing-follows {{
            text-align: center; /* Center align for the message */
            margin: 5px 0; /* Space above and below */
            font-style: italic; /* Italic style */
        }}
        .final-message {{
            margin-top: 10px; /* Space before Thanks and Regards */
        }}
    </style>
</head>
<body>
    <div class='header'>
        <img src='{headerImagePath}' width='100%' style='max-height: 100px;' /> <!-- Header Image -->
    </div>

    <div class='footer'>
        <img src='{footerImagePath}' width='100%' style='max-height: 100px;' /> <!-- Footer Image -->
    </div>

    <div class='content'>
        <div class='title'>PURCHASE ORDER</div>

        <div class='info-box'>
            <div class='info-left'>
                <p><strong>To:</strong> {po.CompanyName ?? "N/A"}</p>
            </div>
            <div class='info-right'>
                <p><strong>P.O. No:</strong> {po.PONumber ?? "N/A"}</p>
                <p><strong>P.O. Date:</strong> {po.PODate.ToString("dd-MMM-yyyy")}</p>
                <p><strong>VAT No:</strong> {po.VATNo ?? "N/A"}</p>
                <p><strong>Project:</strong> {po.Project ?? "N/A"}</p>
            </div>
        </div>

        <p><strong>Attention:</strong> {po.Attention ?? "N/A"}</p>
        <p><strong>Reference:</strong> {po.Reference ?? "N/A"}</p>
        <p><strong>Subject:</strong> Purchase Order for Materials</p>

        <div class='horizontal-line'></div> <!-- Horizontal line below subject -->

        <div class='dear-sir'>
            <p>Dear Sir,</p>
            <p>With reference to your quotation, please be informed that we accepted your terms & conditions with the below prices as mentioned:</p>
        </div>

        <table>
            <thead>
                <tr>
                    <th>#</th> <!-- Added serial number column -->
                    <th>Stock No.</th>
                    <th>Description & Specifications</th>
                    <th>Unit</th>
                    <th>Qty</th>
                    <th>Unit Price</th>
                    <th>Discount</th>
                    <th>VAT%</th>
                    <th>Amount Before Tax</th>
                    <th>VAT Amount</th>
                </tr>
            </thead>
            <tbody>
                {string.Join("", po.Items.Select((item, index) => $@"
                    <tr>
                        <td>{index + 1}</td> <!-- Serial number -->
                        <td>{item.StockNumber ?? "N/A"}</td>
                        <td>{item.Description ?? "N/A"}</td>
                        <td>{item.Unit ?? "N/A"}</td>
                        <td>{item.Quantity}</td>
                        <td>{item.UnitPrice}</td>
                        <td>{item.Discount}</td>
                        <td>{item.VAT}</td>
                        <td>{item.AmountBeforeTax}</td>
                        <td>{item.VATAmount}</td>
                    </tr>
                "))}

                <tr class='nothing-follows'>
                    <td colspan='10'>**** Nothing Follows *****</td> <!-- Added Nothing Follows row -->
                </tr>
                <tr>
                    <td colspan='8'>Total Amount:</td>
                    <td>{po.TotalAmount}</td>
                    <td></td>
                </tr>
                <tr>
                    <td colspan='8'>Less: Discount:</td>
                    <td>{po.Discount}</td>
                    <td></td>
                </tr>
                <tr>
                    <td colspan='8'>Gross Total:</td>
                    <td>{po.TotalAmount - po.Discount}</td>
                    <td></td>
                </tr>
                <tr>
                    <td colspan='8'>VAT Amount:</td>
                    <td>{po.VATAmount}</td>
                    <td></td>
                </tr>
                <tr>
                    <td colspan='8'>Net Amount:</td>
                    <td>{po.NetAmount}</td>
                    <td></td>
                </tr>
            </tbody>
        </table>

        <div class='amount-in-words'>
            Amount in Words: {po.AmountInWords}
        </div>

        <div class='dear-sir final-message'>
            <p>This Purchase Order is raised to order the items as listed and subject to provision attached/listed below and related to our procurement procedures. Please return a copy of this Purchase Order duly signed & stamped in acknowledgement of your receipt, acceptance & compliance to all the terms stated herein.</p>
        </div>

        <div class='total-usd'>
            <p><strong>Total in USD:</strong> {po.TotalAmount * 0.27m}</p> <!-- Replace with appropriate conversion if needed -->
        </div>

        <div class='dear-sir final-message'>
            <p>Thanks and Regards,</p>
        </div>
    </div>
</body>
</html>";
        }
    }

    public class PurchaseOrderViewModel
    {
        public string PONumber { get; set; }
        public DateTime PODate { get; set; }
        public string VATNo { get; set; }
        public string Project { get; set; }
        public string Attention { get; set; }
        public string Reference { get; set; }
        public string Subject { get; set; }
        public List<PurchaseOrderItem> Items { get; set; }
        public decimal TotalAmount { get; set; }
        public decimal Discount { get; set; }
        public decimal VATAmount { get; set; }
        public decimal NetAmount { get; set; }
        public string CompanyName { get; set; }
        public string Address { get; set; }
        public string ContactPerson { get; set; }
        public string AmountInWords { get; set; }
    }

    public class PurchaseOrderItem
    {
        public string StockNumber { get; set; }
        public string Description { get; set; }
        public string Unit { get; set; }
        public int Quantity { get; set; }
        public decimal UnitPrice { get; set; }
        public decimal AmountBeforeTax { get; set; }
        public decimal VAT { get; set; }
        public decimal Discount { get; internal set; }
        public decimal VATAmount { get; internal set; }
    }
}
