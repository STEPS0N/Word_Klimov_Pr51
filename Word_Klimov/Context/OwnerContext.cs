using Microsoft.Office.Interop.Word;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing.Printing;
using System.Linq;
using Word_Klimov.Models;
using Word = Microsoft.Office.Interop.Word;

namespace Word_Klimov.Context
{
    public class OwnerContext : Owner
    {
        public OwnerContext(string img, string firstName, string lastName, string surName, int numberRoom)
            : base(img, firstName, lastName, surName, numberRoom) { }

        public static List<OwnerContext> AllOwners()
        {
            string defImg = System.IO.Path.Combine(
                            AppDomain.CurrentDomain.BaseDirectory,
                            "Images",
                            "owner.png"
                                    );

            List<OwnerContext> allOwners = new List<OwnerContext>();

            allOwners.Add(new OwnerContext(defImg, "Елена", "Иванова", "Петровна", 1));
            allOwners.Add(new OwnerContext(defImg, "Алексей", "Смирнов", "Владимирович", 2));
            allOwners.Add(new OwnerContext(defImg, "Анна", "Кузнецова", "Сергеевна", 3));
            allOwners.Add(new OwnerContext(defImg, "Дмитрий", "Павлов", "Александрович", 3));
            allOwners.Add(new OwnerContext(defImg, "Ольга", "Михайлова", "Ивановна", 4));
            allOwners.Add(new OwnerContext(defImg, "Артем", "Козлов", "Олегович", 5));
            allOwners.Add(new OwnerContext(defImg, "Наталья", "Соколова", "Викторовна", 6));
            allOwners.Add(new OwnerContext(defImg, "Игорь", "Лебедев", "Андреевич", 6));
            allOwners.Add(new OwnerContext(defImg, "Екатерина", "Федорова", "Дмитриевна", 7));
            allOwners.Add(new OwnerContext(defImg, "Андрей", "Александрович", "Игоревич", 7));
            allOwners.Add(new OwnerContext(defImg, "Оксана", "Степанова", "Николаевна", 8));
            allOwners.Add(new OwnerContext(defImg, "Сергей", "Никитин", "Васильевич", 9));
            allOwners.Add(new OwnerContext(defImg, "Мария", "Ковалева", "Александровна", 10));
            allOwners.Add(new OwnerContext(defImg, "Павел", "Фролов", "Михайлович", 11));
            allOwners.Add(new OwnerContext(defImg, "Елена", "Белова", "Александровна", 12));
            allOwners.Add(new OwnerContext(defImg, "Илья", "Поляков", "Данилович", 13));
            allOwners.Add(new OwnerContext(defImg, "Анастасия", "Гаврилова", "Валерьевна", 14));
            allOwners.Add(new OwnerContext(defImg, "Денис", "Орлов", "Владимирович", 15));
            allOwners.Add(new OwnerContext(defImg, "Алина", "Киселева", "Сергеевна", 16));
            allOwners.Add(new OwnerContext(defImg, "Артем", "Ткаченко", "Викторович", 16));
            allOwners.Add(new OwnerContext(defImg, "Валерия", "Романова", "Павловна", 16));
            allOwners.Add(new OwnerContext(defImg, "Александр", "Максимов", "Юрьевич", 17));
            allOwners.Add(new OwnerContext(defImg, "Евгения", "Сидорова", "Игоревна", 17));
            allOwners.Add(new OwnerContext(defImg, "Никита", "Антонов", "Алексеевич", 18));
            allOwners.Add(new OwnerContext(defImg, "Юлия", "Дмитриева", "Владимировна", 19));

            return allOwners;
        }

        public static void Report(string fileName)
        {
            Word.Application app = new Word.Application();
            Word.Document doc = app.Documents.Add();
            Word.Paragraph paraHeader = doc.Paragraphs.Add();
            paraHeader.Range.Font.Size = 16;
            paraHeader.Range.Text = "Список жильцов дома";
            paraHeader.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            paraHeader.Range.ParagraphFormat.SpaceAfter = 0;
            paraHeader.Range.Font.Bold = 1;
            paraHeader.Range.InsertParagraphAfter();

            Word.Paragraph paraAddress = doc.Paragraphs.Add();
            paraAddress.Range.Font.Size = 14;
            paraAddress.Range.Text = "по адресу: г. Пермь, ул. Луначарского, д. 24";
            paraHeader.Range.ParagraphFormat.SpaceAfter = 20;
            paraHeader.Range.Font.Bold = 0;
            paraHeader.Range.InsertParagraphAfter();

            Word.Paragraph paraCount = doc.Paragraphs.Add();
            paraCount.Range.Font.Size = 14;
            paraCount.Range.Text = $"Всего жильцов: {AllOwners().Count}";
            paraCount.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            paraHeader.Range.ParagraphFormat.SpaceAfter = 0;
            paraCount.Range.InsertParagraphAfter();

            Word.Paragraph tableParagraph = doc.Paragraphs.Add();
            Word.Table paymentsTable = doc.Tables.Add(tableParagraph.Range, AllOwners().Count + 1, 6);
            paymentsTable.Borders.InsideLineStyle = paymentsTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            paymentsTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            Cell("№", paymentsTable.Cell(1, 1).Range);
            Cell("Изображение", paymentsTable.Cell(1, 2).Range);
            Cell("Фамилия", paymentsTable.Cell(1, 3).Range);
            Cell("Имя", paymentsTable.Cell(1, 4).Range);
            Cell("Отчество", paymentsTable.Cell(1, 5).Range);
            Cell("№ Квартиры", paymentsTable.Cell(1, 6).Range);

            for (int i = 0; i < AllOwners().Count; i++)
            {
                OwnerContext owner = AllOwners()[i];
                Cell((i + 1).ToString(), paymentsTable.Cell(1 + 1 + i, 1).Range);

                Word.Range imgRange = paymentsTable.Cell(2 + i, 2).Range;
                imgRange.Text = "";

                try
                {
                    if (System.IO.File.Exists(owner.img))
                    {
                        var img = imgRange.InlineShapes.AddPicture(
                            FileName: owner.img,
                            LinkToFile: false,
                            SaveWithDocument: true,
                            Range: imgRange
                        );
                        img.Height = 35;
                        img.Width = 35;
                        imgRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                    else
                    {
                        imgRange.Text = "(нет фото)";
                        imgRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                }
                catch (Exception ex)
                {
                    imgRange.Text = $"Ошибка: {ex.Message}";
                }

                Cell(owner.LastName, paymentsTable.Cell(1 + 1 + i, 3).Range, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                Cell(owner.FirstName, paymentsTable.Cell(1 + 1 + i, 4).Range, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                Cell(owner.SurName, paymentsTable.Cell(1 + 1 + i, 5).Range, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                Cell(owner.NumberRoom.ToString(), paymentsTable.Cell(1 + 1 + i, 6).Range, Word.WdParagraphAlignment.wdAlignParagraphLeft);
            }

            doc.SaveAs2(fileName);
            doc.Close();
            app.Quit();
        }

        private static void Cell(string Text, Word.Range Cell,
            Word.WdParagraphAlignment Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter)
        {
            Cell.Text = Text;
            Cell.ParagraphFormat.Alignment = Alignment;
        }

        public static void ReportPDF(string fileName)
        {
            PdfDocument document = new PdfDocument();
            document.Info.Title = "Отчёт по жильцам дома";
            PdfPage page = document.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(page);
            int MarginTop = 20;
            int MarginLeft = 50;
            XFont fontHeader = new XFont("Arial", 16, XFontStyle.Bold);
            XFont font = new XFont("Arial", 12);
            gfx.DrawString("Список жильцов дома", fontHeader, XBrushes.Black,
                new XRect(0, MarginTop, page.Width, 15),
                XStringFormats.Center);
            gfx.DrawString("по адресу: г. Пермь, ул. Луначарского, д. 24", font, XBrushes.Black,
                new XRect(0, MarginTop + 30, page.Width, 10),
                XStringFormats.Center);
            gfx.DrawString("Всего жильцов: " + AllOwners().Count, font, XBrushes.Black,
                new XRect(MarginLeft, MarginTop + 70, page.Width, 10),
                XStringFormats.CenterLeft);

            int Width = (Convert.ToInt32(page.Width.Value) - MarginLeft * 2 - 30) / 4;
            gfx.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft, MarginTop + 100, Width, 20);
            gfx.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft + Width + 10, MarginTop + 100, Width, 20);
            gfx.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft + (Width + 10) * 2, MarginTop + 100, Width, 20);
            gfx.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft + (Width + 10) * 3, MarginTop + 100, Width, 20);

            gfx.DrawString("№" + AllOwners().Count, font, XBrushes.Black,
                new XRect(MarginLeft, MarginTop + 100, Width, 20),
                XStringFormats.Center);
            gfx.DrawString("Фамилия", font, XBrushes.Black,
                new XRect(MarginLeft + Width + 10, MarginTop + 100, Width, 20),
                XStringFormats.Center);
            gfx.DrawString("Имя", font, XBrushes.Black,
                new XRect(MarginLeft + (Width + 10) * 2, MarginTop + 100, Width, 20),
                XStringFormats.Center);
            gfx.DrawString("Отчество", font, XBrushes.Black,
                new XRect(MarginLeft + (Width + 10) * 3, MarginTop + 100, Width, 20),
                XStringFormats.Center);

            for (int i = 0; i < AllOwners().Count; i++)
            {
                gfx.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft, MarginTop + 100 + 25 * (i + 1), Width, 20);
                gfx.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft + Width + 10, MarginTop + 100 + 25 * (i + 1), Width, 20);
                gfx.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft + (Width + 10) * 2, MarginTop + 100 + 25 * (i + 1), Width, 20);
                gfx.DrawRectangle(new XSolidBrush(XColors.LightGray), MarginLeft + (Width + 10) * 3, MarginTop + 100 + 25 * (i + 1), Width, 20);

                gfx.DrawString((i + 1).ToString(), font, XBrushes.Black,
                    new XRect(MarginLeft, MarginTop + 100 + 25 * (i + 1), Width, 20),
                    XStringFormats.Center);
                gfx.DrawString(AllOwners()[i].LastName, font, XBrushes.Black,
                    new XRect(MarginLeft + Width + 10, MarginTop + 100 + 25 * (i + 1), Width, 20),
                    XStringFormats.Center);
                gfx.DrawString(AllOwners()[i].FirstName, font, XBrushes.Black,
                    new XRect(MarginLeft + (Width + 10) * 2, MarginTop + 100 + 25 * (i + 1), Width, 20),
                    XStringFormats.Center);
                gfx.DrawString(AllOwners()[i].SurName, font, XBrushes.Black,
                    new XRect(MarginLeft + (Width + 10) * 3, MarginTop + 100 + 25 * (i + 1), Width, 20),
                    XStringFormats.Center);
            }
            document.Save(fileName);
            Process.Start(fileName);
        }
    }
}
