using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXML.Excel.Export
{
    public class CellHelper
    {
        public static Cell ConstructCell(string value, CellValues dataType, uint styleIndex = 0)
        {
            return new Cell
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType),
                StyleIndex = styleIndex
            };
        }

        public static Stylesheet GenerateStylesheet()
        {
            var fonts = new Fonts(
                new Font(new FontSize { Val = 10 }),
                new Font(new FontSize { Val = 10 }, new Bold(), new Color { Rgb = "FFFFFF" }));

            var fills = new Fills(
                    new Fill(new PatternFill { PatternType = PatternValues.None }),
                    new Fill(new PatternFill { PatternType = PatternValues.Gray125 }),
                    new Fill(new PatternFill( new ForegroundColor { Rgb = new HexBinaryValue { Value = "66666666" } })
                    {
                        PatternType = PatternValues.Solid
                    }));

            var borders = new Borders(
                    new Border(), 
                    new Border( 
                        new LeftBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
                        new TopBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())
                );

            var cellFormats = new CellFormats(
                    new CellFormat(),
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }, // body
                    new CellFormat { FontId = 1, FillId = 2, BorderId = 1, ApplyFill = true } // header
                );

            return new Stylesheet(fonts, fills, borders, cellFormats);
        }
    }
}
