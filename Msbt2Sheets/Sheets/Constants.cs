using Google.Apis.Sheets.v4.Data;

namespace Msbt2Sheets.Sheets;

public class Constants
{
    public static CellFormat HEADER_CELL_FORMAT = new CellFormat()
    {
        BackgroundColorStyle = new ColorStyle()
        {
            RgbColor = new Color()
            {
                Red = 0.937f, Blue = 0.937f, Green = 0.937f, Alpha = 1
            }
        },
        TextFormat = new TextFormat()
        {
            Bold = true
        }
    };
    
    public static CellFormat HEADER_CELL_FORMAT_LEFT_BORDER = new CellFormat()
    {
        BackgroundColorStyle = new ColorStyle()
        {
            RgbColor = new Color()
            {
                Red = 0.937f, Blue = 0.937f, Green = 0.937f, Alpha = 1
            }
        },
        TextFormat = new TextFormat()
        {
            Bold = true
        },
        Borders = new Borders()
        {
            Left = new Border()
            {
                Style = "SOLID"
            },
        }
    };
    
    public static CellFormat HEADER_CELL_FORMAT_CENTERED = new CellFormat()
    {
        BackgroundColorStyle = new ColorStyle()
        {
            RgbColor = new Color()
            {
                Red = 0.937f, Blue = 0.937f, Green = 0.937f, Alpha = 1
            }
        },
        TextFormat = new TextFormat()
        {
            Bold = true
        },
        HorizontalAlignment = "CENTER"
    };
    
    public static CellFormat HEADER_CELL_FORMAT_CENTERED_LEFT_BORDER = new CellFormat()
    {
        BackgroundColorStyle = new ColorStyle()
        {
            RgbColor = new Color()
            {
                Red = 0.937f, Blue = 0.937f, Green = 0.937f, Alpha = 1
            }
        },
        TextFormat = new TextFormat()
        {
            Bold = true
        },
        HorizontalAlignment = "CENTER",
        Borders = new Borders()
        {
            Left = new Border()
            {
                Style = "SOLID"
            },
        }
    };
    
    public static CellFormat HEADER_CELL_FORMAT_PERCENT = new CellFormat()
    {
        BackgroundColorStyle = new ColorStyle()
        {
            RgbColor = new Color()
            {
                Red = 0.937f, Blue = 0.937f, Green = 0.937f, Alpha = 1
            }
        },
        TextFormat = new TextFormat()
        {
            Bold = true
        },
        NumberFormat = new NumberFormat()
        {
            Type = "PERCENT",
            Pattern = "0.00%"
        },
    };
    
    public static CellFormat HEADER_CELL_FORMAT_PERCENT_CENTERED = new CellFormat()
    {
        BackgroundColorStyle = new ColorStyle()
        {
            RgbColor = new Color()
            {
                Red = 0.937f, Blue = 0.937f, Green = 0.937f, Alpha = 1
            }
        },
        TextFormat = new TextFormat()
        {
            Bold = true
        },
        NumberFormat = new NumberFormat()
        {
            Type = "PERCENT",
            Pattern = "0.00%"
        },
        HorizontalAlignment = "CENTER"
    };
}