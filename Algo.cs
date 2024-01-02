using OfficeOpenXml;
using System.Linq;
using System.IO;
using OfficeOpenXml.Style;
using System;
using Microsoft.UI.Xaml.Media;
using System.Drawing;
using Microsoft.UI;
using Microsoft.UI.Xaml.Shapes;
using Path = System.IO.Path;
namespace ConductivityReportAlgo
{
    public class Algo
    {
        private ExcelPackage outPack;
        private ExcelWorksheet outSheet;
        private string path;
        private string savePath;

        private float r1;
        private SolidColorBrush Color1;
        private float r2;
        private SolidColorBrush Color2;
        private float r3;
        private SolidColorBrush Color3;
        private float r4;
        private SolidColorBrush Color4;
        private float r5;
        private SolidColorBrush Color5;
        private float r6;

        private bool allow1;
        private bool allow2;
        private bool allow3;
        private bool allow4;
        private bool allow5;
        public Algo(bool allow1, bool allow2, bool allow3, bool allow4, bool allow5,SolidColorBrush color1, SolidColorBrush color2, SolidColorBrush color3, SolidColorBrush color4, SolidColorBrush color5,string path,
            float r1,float r2,float r3,float r4,float r5 , string savePath = "")
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            outPack = new ExcelPackage();
            outSheet = outPack.Workbook.Worksheets.Add("build");
            this.path = path;
            this.savePath = savePath;
            this.Color1 = color1;
            this.Color2 = color2;
            this.Color3 = color3;
            this.Color4 = color4;
            this.Color5 = color5;
            this.r1 =r1;
            this.r2 =r2;
            this.r3 =r3;
            this.r4 =r4;
            this.r5 =r5;
            this.allow4 = allow1;
            this.allow4 = allow2;
            this.allow4 = allow3;
            this.allow4 = allow4;
            this.allow5 = allow5;
        }
        public void scanValueSummery(string topFileName,string bottomFileName)
        {
            //Assign Titels
            outSheet.Cells["U3"].Value = "Scan Values Summary:";
            outSheet.Cells["U3:X3"].Merge = true;
            outSheet.Cells["U4"].Value = "Min Conductivity:";
            outSheet.Cells["U4:X4"].Merge = true;
            outSheet.Cells["U5"].Value = "Max Conductivity:";
            outSheet.Cells["U5:X5"].Merge = true;
            outSheet.Cells["U6"].Value = "Conductivity Range:";
            outSheet.Cells["U6:X6"].Merge = true;

            //Assign Top Values
            var topPack = new ExcelPackage(new FileInfo(topFileName));
            var topWorkSheet = topPack.Workbook.Worksheets.First();
            outSheet.Cells["Y3"].Value = "Top";
            outSheet.Cells["Y3:Z3"].Merge = true;
            outSheet.Cells["Y4"].Value = topWorkSheet.Cells["A2"].Value;
            outSheet.Cells["Y5"].Value = topWorkSheet.Cells["B2"].Value;
            outSheet.Cells["Y6"].Value = topWorkSheet.Cells["C2"].Value;
            outSheet.Cells["Z4"].Value = "%IACS";
            outSheet.Cells["Z5"].Value = "%IACS";
            outSheet.Cells["Z6"].Value = "%IACS";
            //Assign Bottom Values
            var botPack = new ExcelPackage(new FileInfo(bottomFileName));
            var botWorkSheet = botPack.Workbook.Worksheets.First();
            outSheet.Cells["AB3"].Value = "Bottom";
            outSheet.Cells["AB3:AC3"].Merge = true;
            outSheet.Cells["AB4"].Value = botWorkSheet.Cells["A2"].Value;
            outSheet.Cells["AB5"].Value = botWorkSheet.Cells["B2"].Value;
            outSheet.Cells["AB6"].Value = botWorkSheet.Cells["C2"].Value;
            outSheet.Cells["AC4"].Value = "%IACS";
            outSheet.Cells["AC5"].Value = "%IACS";
            outSheet.Cells["AC6"].Value = "%IACS";
            //Assign Plate Values
            outSheet.Cells["AE3:AF3"].Merge = true;
            outSheet.Cells["AE3"].Value = "Plate";
            outSheet.Cells["AE4:AE6"].Value = "NaN"; // needs values here
            outSheet.Cells["AF4"].Value = "%IACS";
            outSheet.Cells["AF5"].Value = "%IACS";
            outSheet.Cells["AF6"].Value = "%IACS";
            //Assign Part Values
            outSheet.Cells["AH3:AI3"].Merge = true;
            outSheet.Cells["AH3"].Value = "Part";
            outSheet.Cells["AH4:AH6"].Value = "NaN"; // needs values here
            outSheet.Cells["AI4"].Value = "%IACS";
            outSheet.Cells["AI5"].Value = "%IACS";
            outSheet.Cells["AI6"].Value = "%IACS";

            // Borders
            var borderRange = outSheet.Cells["U3:AI6"];
            // Set Border Style
            var borderTopRange = outSheet.Cells["U3:AI3"];
            var borderLeftRange = outSheet.Cells["U3:U6"];
            var borderRightRange = outSheet.Cells["AI3:AI6"];
            var borderBottomRange = outSheet.Cells["U6:AI6"];
            borderTopRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            borderBottomRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            borderLeftRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            borderRightRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            savePack();
        }
        public void scanValueSummery(string path)
        {
            //Assign Titels
            outSheet.Cells["U3"].Value = "Scan Values Summary:";
            outSheet.Cells["U3:X3"].Merge = true;
            outSheet.Cells["U4"].Value = "Min Conductivity:";
            outSheet.Cells["U4:X4"].Merge = true;
            outSheet.Cells["U5"].Value = "Max Conductivity:";
            outSheet.Cells["U5:X5"].Merge = true;
            outSheet.Cells["U6"].Value = "Conductivity Range:";
            outSheet.Cells["U6:X6"].Merge = true;

            //Assign Top Values
            var topPack = new ExcelPackage(new FileInfo(path));
            var sheet = topPack.Workbook.Worksheets.First();
            outSheet.Cells["Y3"].Value = "Top";
            outSheet.Cells["Y3:Z3"].Merge = true;
            outSheet.Cells["Y4"].Value = autoDetectValue(sheet, "Min Top conductivity = (%IACS)");
            outSheet.Cells["Y5"].Value = autoDetectValue(sheet, "Max Top conductivity = (%IACS)");
            outSheet.Cells["Y6"].Value = autoDetectValue(sheet, "Conductivity Top range = (%IACS)");
            outSheet.Cells["Z4"].Value = "%IACS";
            outSheet.Cells["Z5"].Value = "%IACS";
            outSheet.Cells["Z6"].Value = "%IACS";
            //Assign Bottom Values
            outSheet.Cells["AB3"].Value = "Bottom";
            outSheet.Cells["AB3:AC3"].Merge = true;
            outSheet.Cells["AB4"].Value = autoDetectValue(sheet, "Min Bottom conductivity = (%IACS)");
            outSheet.Cells["AB5"].Value = autoDetectValue(sheet, "Max Bottom conductivity = (%IACS)");
            outSheet.Cells["AB6"].Value = autoDetectValue(sheet, "Conductivity Bottom range = (%IACS)");
            outSheet.Cells["AC4"].Value = "%IACS";
            outSheet.Cells["AC5"].Value = "%IACS";
            outSheet.Cells["AC6"].Value = "%IACS";
            //Assign Plate Values
            outSheet.Cells["AE3:AF3"].Merge = true;
            outSheet.Cells["AE3"].Value = "Plate";
            outSheet.Cells["AE4"].Value = autoDetectValue(sheet, "Min Plate conductivity = (%IACS)");
            outSheet.Cells["AE5"].Value = autoDetectValue(sheet, "Max Plate conductivity = (%IACS)");
            outSheet.Cells["AE6"].Value = autoDetectValue(sheet, "Conductivity Plate range = (%IACS)");
            outSheet.Cells["AF4"].Value = "%IACS";
            outSheet.Cells["AF5"].Value = "%IACS";
            outSheet.Cells["AF6"].Value = "%IACS";
            //Assign Part Values
            outSheet.Cells["AH3:AI3"].Merge = true;
            outSheet.Cells["AH3"].Value = "Part";
            outSheet.Cells["AH4"].Value = autoDetectValue(sheet, "Min Part conductivity = (%IACS)");
            outSheet.Cells["AH5"].Value = autoDetectValue(sheet, "Max Part conductivity = (%IACS)");
            outSheet.Cells["AH6"].Value = autoDetectValue(sheet, "Conductivity Part range = (%IACS)"); // needs values here
            outSheet.Cells["AI4"].Value = "%IACS";
            outSheet.Cells["AI5"].Value = "%IACS";
            outSheet.Cells["AI6"].Value = "%IACS";

            // Borders
            // Set Border Style
            var borderTopRange = outSheet.Cells["U3:AI3"];
            var borderLeftRange = outSheet.Cells["U3:U6"];
            var borderRightRange = outSheet.Cells["AI6:AI6"];
            var borderBottomRange = outSheet.Cells["U3:AI6"];
            borderTopRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            borderBottomRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            borderLeftRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            borderRightRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            savePack();
        }
        public void scanTempSummery(string temperatureFile)
        {
            var pack = new ExcelPackage(new FileInfo(temperatureFile));
            var sheet = pack.Workbook.Worksheets.First();
            //Calibration Temp
            outSheet.Cells["U10"].Value = "Calibration Temperature";
            outSheet.Cells["U10:X10"].Merge = true;
            outSheet.Cells["Y10"].Value = autoDetectValue(sheet, "Calibration temperature Min (°C)");
            outSheet.Cells["Z10"].Value = "°C";
            outSheet.Cells["AA10"].Value = autoDetectValue(sheet, "Calibration temperature Max (°C)");
            outSheet.Cells["AB10"].Value = "°C";
            //Assign Scan temperature
            outSheet.Cells["U11"].Value = "Scan temperature";
            outSheet.Cells["U11:X11"].Merge = true; 
            outSheet.Cells["Y11"].Value = autoDetectValue(sheet, "Scan temperature Min (°C)");
            outSheet.Cells["Z11"].Value = "°C";
            outSheet.Cells["AA11"].Value = autoDetectValue(sheet, "Scan temperature Max (°C)");
            outSheet.Cells["AB11"].Value = "°C";
            
            outSheet.Cells["AC9"].Value = "Range";
            outSheet.Cells["AA9"].Value = "Max";
            outSheet.Cells["Y9"].Value = "Min";
            outSheet.Cells["AC10"].Value = autoDetectValue(sheet, "Range (°C)");
            outSheet.Cells["AC10:AC11"].Merge = true;
            outSheet.Cells["AD10"].Value = "°C";
            outSheet.Cells["AD10:AD11"].Merge = true;
            // Assign borders
            // Set Border Style
            var borderTopRange = outSheet.Cells["U9:AD9"];
            var borderLeftRange = outSheet.Cells["U9:U11"];
            var borderRightRange = outSheet.Cells["AD9:AD11"];
            var borderBottomRange = outSheet.Cells["U11:AD11"];
            borderTopRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            borderBottomRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            borderLeftRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            borderRightRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            savePack();
        }
        public void scanDetailsSummery(string partFile)
        {
            var pack = new ExcelPackage(new FileInfo(partFile));
            var sheet = pack.Workbook.Worksheets.First();
            outSheet.Cells["N10"].Value = "Calibration Date";
            outSheet.Cells["N10:P10"].Merge = true;
            outSheet.Cells["Q10"].Value = autoDetectValue(sheet, "Calibration date");
            outSheet.Cells["Q10:R10"].Merge = true;

            outSheet.Cells["N11"].Value = "Calibration Time";
            outSheet.Cells["N11:P11"].Merge = true;
            outSheet.Cells["Q11"].Value = autoDetectValue(sheet, "Calibration time");
            outSheet.Cells["Q11:R11"].Merge = true;

            outSheet.Cells["N12"].Value = "Inspection date";
            outSheet.Cells["N12:P12"].Merge = true;
            outSheet.Cells["Q12"].Value = autoDetectValue(sheet, "Inspection date");
            outSheet.Cells["Q12:R12"].Merge = true;

            outSheet.Cells["N13"].Value = "Inspection time";
            outSheet.Cells["N13:P13"].Merge = true;
            outSheet.Cells["Q13"].Value = autoDetectValue(sheet, "Inspection time");
            outSheet.Cells["Q13:R13"].Merge = true;

            // Assign Border
            // Set Border Style
            var borderTopRange = outSheet.Cells["N10:R10"];
            var borderLeftRange = outSheet.Cells["N10:N13"];
            var borderRightRange = outSheet.Cells["R10:R13"];
            var borderBottomRange = outSheet.Cells["N13:R13"];
            borderTopRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            borderBottomRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            borderLeftRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            borderRightRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            savePack();
        }
        public void partDetailsSummery(string partFile)
        {
            var pack = new ExcelPackage(new FileInfo(partFile));
            var sheet = pack.Workbook.Worksheets.First();
            outSheet.Cells["N3"].Value = "Part #";
            outSheet.Cells["N3:P3"].Merge = true;
            outSheet.Cells["Q3"].Value = autoDetectValue(sheet, "Part #");
            outSheet.Cells["Q3:R3"].Merge = true;
            outSheet.Cells["N4"].Value = "Plate in the part #";
            outSheet.Cells["N4:P4"].Merge = true;
            outSheet.Cells["Q4"].Value = autoDetectValue(sheet, "Plate in the part #");
            outSheet.Cells["Q4:R4"].Merge = true;
            outSheet.Cells["N5"].Value = "Plate thickness";
            outSheet.Cells["N5:P5"].Merge = true;
            outSheet.Cells["Q5"].Value = autoDetectValue(sheet, "Plate thickness (mm)");
            outSheet.Cells["Q5:R5"].Merge = true;
            outSheet.Cells["N6"].Value = "Plate width";
            outSheet.Cells["N6:P6"].Merge = true;
            outSheet.Cells["Q6"].Value = autoDetectValue(sheet, "Plate width (mm)");
            outSheet.Cells["Q6:R6"].Merge = true;
            outSheet.Cells["N7"].Value = "Plate length";
            outSheet.Cells["N7:P7"].Merge = true;
            outSheet.Cells["Q7"].Value = autoDetectValue(sheet, "Plate length (mm)");
            outSheet.Cells["Q7:R7"].Merge = true;
            outSheet.Cells["S5"].Value = "mm";
            outSheet.Cells["S5:T5"].Merge = true;
            outSheet.Cells["S6"].Value = "mm";
            outSheet.Cells["S6:T6"].Merge = true;
            outSheet.Cells["S7"].Value = "mm";
            outSheet.Cells["S7:T7"].Merge = true;
            outSheet.Cells["N8"].Value = "Equpment";
            outSheet.Cells["N8:P8"].Merge = true;
            //Assign Border
            // Set Border Style
            var borderTopRange = outSheet.Cells["N3:T3"];
            var borderLeftRange = outSheet.Cells["N3:N7"];
            var borderRightRange = outSheet.Cells["T3:T7"];
            var borderBottomRange = outSheet.Cells["N7:T7"];
            borderTopRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            borderBottomRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            borderLeftRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            borderRightRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            savePack();
        }

        public void titleSummery(string partFile)
        {
            var pack = new ExcelPackage(new FileInfo(partFile));
            var sheet = pack.Workbook.Worksheets.First();
            outSheet.Cells["A3"].Value = "Alloy";
            outSheet.Cells["A3:C3"].Merge = true;
            outSheet.Cells["D3"].Value = autoDetectValue(sheet, "Alloy");
            outSheet.Cells["D3:H3"].Merge = true;

            outSheet.Cells["A4"].Value = "Temper";
            outSheet.Cells["A4:C4"].Merge = true;
            outSheet.Cells["D4"].Value = autoDetectValue(sheet, "Temper");
            outSheet.Cells["D4:H4"].Merge = true;

            outSheet.Cells["A5"].Value = "Reference standard";
            outSheet.Cells["A5:C5"].Merge = true;
            outSheet.Cells["D5"].Value = autoDetectValue(sheet, "Reference standard");
            outSheet.Cells["D5:H5"].Merge = true;

            outSheet.Cells["A7"].Value = "Measured points distance not more than:";
            outSheet.Cells["A7:G7"].Merge = true;
            outSheet.Cells["A8"].Value = "in X";
            outSheet.Cells["A8:E8"].Merge = true;
            outSheet.Cells["F8"].Value = autoDetectValue(sheet, "in X (mm)");
            outSheet.Cells["G8"].Value = "mm";

            outSheet.Cells["A9"].Value = "in Y";
            outSheet.Cells["A9:E9"].Merge = true;
            outSheet.Cells["F9"].Value = autoDetectValue(sheet, "in Y (mm)");
            outSheet.Cells["G9"].Value = "mm";
            //Rejection Sector
            outSheet.Cells["A10"].Value = "Rejection criterias:";
            outSheet.Cells["A10:G10"].Merge = true;
            outSheet.Cells["G11:G14"].Value = "%IACS";

            outSheet.Cells["A11"].Value = "Min Conductivity";
            outSheet.Cells["A11:E11"].Merge = true;
            outSheet.Cells["F11"].Value = autoDetectValue(sheet, "Min conductivity (%IACS)");

            outSheet.Cells["A12"].Value = "Max Conductivity";
            outSheet.Cells["A12:E12"].Merge = true;
            outSheet.Cells["F12"].Value = autoDetectValue(sheet, "Max conductivity (%IACS)");

            outSheet.Cells["A13"].Value = "Conductivity range at plate";
            outSheet.Cells["A13:E13"].Merge = true;
            outSheet.Cells["F13"].Value = autoDetectValue(sheet, "Conductivity range at plate (%IACS)");

            outSheet.Cells["A14"].Value = "Conductivity range in the part";
            outSheet.Cells["A14:E14"].Merge = true;
            outSheet.Cells["F14"].Value = autoDetectValue(sheet, "Conductivity range in the part (%IACS)");

            // Set Border Style
            var borderTopRange = outSheet.Cells["A3:H3"];
            var borderLeftRange = outSheet.Cells["A3:A14"];
            var borderRightRange = outSheet.Cells["H3:H14"];
            var borderBottomRange = outSheet.Cells["A14:H14"];
            borderTopRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            borderBottomRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            borderLeftRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            borderRightRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            savePack();

        }
        public void dataTableTop(string partFile,ExcelRangeBase topCorner,ExcelRangeBase bottomCorner)
        {
            if(topCorner == null || bottomCorner == null) { return; }
            var pack = new ExcelPackage(new FileInfo(partFile));
            var sheet = pack.Workbook.Worksheets.First();
            outSheet.Cells["A18"].Value = "Top Side";
            outSheet.Cells["A18:B18"].Merge = true;
            int topRow = topCorner.Start.Row;
            int topCol = topCorner.Start.Column;
            int botRow = bottomCorner.Start.Row;
            int botCol = bottomCorner.Start.Column;
            int topRowColor = topRow + 1;
            int topColColor = topCol + 1;
            for(int i= 0; i <= botRow-topRow;i++)
            {
                for(int j = 0; j <= botCol-topCol; j++)
                {
                    this.outSheet.Cells[20+i,1+j].Value = sheet.Cells[topRow+i, topCol+j].Value;
                    var cellVal = this.outSheet.Cells[20 + i, 1 + j].Value;
                    if (cellVal != null && float.TryParse(cellVal.ToString(),out float floatValue))
                    {
                        if (floatValue<=r5 && this.allow5)
                        {
                            outSheet.Cells[20+i,1+j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            outSheet.Cells[20 + i, 1 + j].Style.Fill.BackgroundColor.SetColor(colorConvert(Color5));
                            continue;
                        }
                        if (floatValue <= r4 && this.allow4)
                        {
                            outSheet.Cells[20 + i, 1 + j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            outSheet.Cells[20 + i, 1 + j].Style.Fill.BackgroundColor.SetColor(colorConvert(Color4));
                            continue;
                        }
                        if (floatValue <= r3)
                        {
                            outSheet.Cells[20 + i, 1 + j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            outSheet.Cells[20 + i, 1 + j].Style.Fill.BackgroundColor.SetColor(colorConvert(Color3));
                            continue;
                        }
                        if (floatValue <= r2)
                        {
                            outSheet.Cells[20 + i, 1 + j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            outSheet.Cells[20 + i, 1 + j].Style.Fill.BackgroundColor.SetColor(colorConvert(Color2));
                            continue;
                        }
                        if (floatValue <= r1)
                        {
                            outSheet.Cells[20 + i, 1 + j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            outSheet.Cells[20 + i, 1 + j].Style.Fill.BackgroundColor.SetColor(colorConvert(Color1));
                            continue;
                        }
                    }
                }
            }
            
            this.outSheet.Cells["B19"].Value = "Y, mm";
            this.outSheet.Cells[19, 2, 19, botCol].Merge = true;
            savePack();
        }
        private System.Drawing.Color colorConvert(SolidColorBrush brush) 
        {
            var color = brush.Color;
            return System.Drawing.Color.FromArgb(color.R, color.G, color.B);
        }
        public void dataTableBottom(string partFile, ExcelRangeBase topCorner, ExcelRangeBase bottomCorner)
        {
            if (topCorner == null || bottomCorner == null) { return; }
            var pack = new ExcelPackage(new FileInfo(partFile));
            var sheet = pack.Workbook.Worksheets.First();
            outSheet.Cells["Z18"].Value = "Top Side";
            outSheet.Cells["Z18:AA18"].Merge = true;
            int topRow = topCorner.Start.Row;
            int topCol = topCorner.Start.Column;
            int botRow = bottomCorner.Start.Row;
            int botCol = bottomCorner.Start.Column;
            int topRowColor = topRow + 1;
            int topColColor = topCol + 1;
            for (int i = 0; i <= botRow - topRow; i++)
            {
                for (int j = 0; j <= botCol - topCol; j++)
                {
                    this.outSheet.Cells[20 + i, 1 + j].Value = sheet.Cells[topRow + i, topCol + j].Value;
                    var cellVal = this.outSheet.Cells[20 + i, 1 + j].Value;
                    if (cellVal != null && float.TryParse(cellVal.ToString(), out float floatValue))
                    {
                        if (floatValue >= r1 && floatValue < r2)
                        {
                            sheet.Cells[20 + i, 1 + j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[20 + i, 1 + j].Style.Fill.BackgroundColor.SetColor(colorConvert(Color1));
                            continue;
                        }
                        if (floatValue >= r2 && floatValue < r3)
                        {
                            sheet.Cells[20 + i, 1 + j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[20 + i, 1 + j].Style.Fill.BackgroundColor.SetColor(colorConvert(Color2));
                            continue;
                        }
                        if (floatValue >= r3 && floatValue < r4)
                        {
                            sheet.Cells[20 + i, 1 + j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[20 + i, 1 + j].Style.Fill.BackgroundColor.SetColor(colorConvert(Color3));
                            continue;
                        }
                        if (floatValue >= r4 && floatValue < r5)
                        {
                            sheet.Cells[20 + i, 1 + j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[20 + i, 1 + j].Style.Fill.BackgroundColor.SetColor(colorConvert(Color4));
                            continue;
                        }
                        if (floatValue >= r5 && floatValue < r6)
                        {
                            sheet.Cells[20 + i, 1 + j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            sheet.Cells[20 + i, 1 + j].Style.Fill.BackgroundColor.SetColor(colorConvert(Color5));
                            continue;
                        }
                    }
                }
            }

            this.outSheet.Cells["AA19"].Value = "Y, mm";
            this.outSheet.Cells[19, 2, 19, botCol].Merge = true;
            savePack();
        }
        public bool savePack()
        {
            try
            {
                this.outSheet.Cells.AutoFitColumns();
                this.outPack.SaveAs(System.IO.Path.Combine(this.savePath, "Report.xlsx"));
                return true;
            } catch
            {
                this.outPack.SaveAs(System.IO.Path.Combine(this.savePath, "Report2.xlsx"));
                
                return false;
            }

        }
        public ExcelRangeBase getTopCorner(string file)
        {
            var pack = new ExcelPackage(new FileInfo(file));
            var sheet = pack.Workbook.Worksheets.First();
            var val = autoDetectValue(sheet, "for Top Measured Table (Top Left Corner)");
            if (val.Equals(""))
            {
                return null;
            } 
            else
            {
                return sheet.Cells[$"{val}"];
            }
        }
        public ExcelRangeBase getBottomCorner(string file)
        {
            var pack = new ExcelPackage(new FileInfo(file));
            var sheet = pack.Workbook.Worksheets.First();
            var val = autoDetectValue(sheet, "for Top Measured Table (Bottom Right Corner)");
            if (val.Equals(""))
            {
                return null;
            }
            else
            {
                return sheet.Cells[$"{val}"];
            }
        }
        public string autoDetectValue(ExcelWorksheet sheet, string valueName)
        {
            try {
                var query = from cell in sheet.Cells["A1:Z50"] where cell.Value?.ToString() == valueName select cell;
                int row = query.First().Start.Row;
                int col = query.First().Start.Column;
                if(sheet.Cells[row, col + 1].Value != null)
                {
                    return sheet.Cells[row, col + 1].Value.ToString();

                } else
                {
                    return "";
                }

            } catch
            {
                return "";
            }
        }

        public static void addTopFile(string fileCSV, ExcelPackage pack)
        {
            var sheet = pack.Workbook.Worksheets.First();
            string[] lines = File.ReadAllLines(fileCSV);
            string[] parts = lines[1].Split(',');
            sheet.Cells["K4"].Value = parts[0];
            sheet.Cells["K5"].Value = parts[1];
            sheet.Cells["K6"].Value = parts[2];
            pack.Save();
        }
        public static void addBottomFile(ExcelPackage pack, string fileCSV)
        {
            var sheet = pack.Workbook.Worksheets.First();
            string[] lines = File.ReadAllLines(fileCSV);
            string[] parts = lines[1].Split(',');
            sheet.Cells["K7"].Value = parts[0];
            sheet.Cells["K8"].Value = parts[1];
            sheet.Cells["K9"].Value = parts[2];
            pack.Save();
        }

        public static string semiPackage(string path)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var pack = new ExcelPackage();
            var sheet = pack.Workbook.Worksheets.Add("sample");
            sheet.Cells["Z1"].Value = "keySM";
            sheet.Cells["Z1"].Value = "keySM";
            sheet.Cells["Z1"].Style.Fill.PatternType = ExcelFillStyle.LightTrellis;
            sheet.Cells["Z1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.PaleVioletRed);
            sheet.Cells["Z1"].Style.Locked = true;
            sheet.Cells["A1"].Value = "Enter Sample Data in the Coresponding Columns";
            sheet.Cells["A1:D1"].Merge = true;
            //plate data
            sheet.Cells["B3:B12"].Style.Fill.PatternType = ExcelFillStyle.LightGray;
            sheet.Cells["B3:B12"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightSlateGray);
            sheet.Cells["A3"].Value = "Plate Data";
            sheet.Cells["A3:B3"].Merge = true;
            sheet.Cells["A4"].Value = "Part #";
            sheet.Cells["A5"].Value = "Plate in the part #";
            sheet.Cells["A6"].Value = "Plate thickness (mm)";
            sheet.Cells["A7"].Value = "Plate width (mm)";
            sheet.Cells["A8"].Value = "Plate length (mm)";
            sheet.Cells["A9"].Value = "Calibration date";
            sheet.Cells["A10"].Value = "Calibration time";
            sheet.Cells["A11"].Value = "Inspection date";
            sheet.Cells["A12"].Value = "Inspection time";

            // scan values
             sheet.Cells["J3"].Value = "Scan values summary:";
            sheet.Cells["J3:K3"].Merge = true;
            sheet.Cells["J4"].Value = "Min Top conductivity = (%IACS)";
            sheet.Cells["J5"].Value = "Max Top conductivity = (%IACS)";
            sheet.Cells["J6"].Value = "Conductivity Top range = (%IACS)";

            sheet.Cells["J7"].Value = "Min Bottom conductivity = (%IACS)";
            sheet.Cells["J8"].Value = "Max Bottom conductivity = (%IACS)";
            sheet.Cells["J9"].Value = "Conductivity Bottom range = (%IACS)";

            sheet.Cells["J10"].Value = "Min Plate conductivity = (%IACS)";
            sheet.Cells["J11"].Value = "Max Plate conductivity = (%IACS)";
            sheet.Cells["J12"].Value = "Conductivity Plate range = (%IACS)";

            sheet.Cells["J13"].Value = "Min Part conductivity = (%IACS)";
            sheet.Cells["J14"].Value = "Max Part conductivity = (%IACS)";
            sheet.Cells["J15"].Value = "Conductivity Part range = (%IACS)";


            sheet.Cells["K4:K15"].Style.Fill.PatternType = ExcelFillStyle.LightGray;
            sheet.Cells["K4:K15"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightSlateGray);

            //Alloy Data
            sheet.Cells["D3"].Value = "Alloy Data";
            sheet.Cells["D3:E3"].Merge = true;
            sheet.Cells["D4"].Value = "Alloy";
            sheet.Cells["D5"].Value = "Temper";
            sheet.Cells["D6"].Value = "Reference standard";
            sheet.Cells["D7"].Value = "Measured points distance not more than:";
            sheet.Cells["D7:E7"].Merge = true;
            sheet.Cells["D8"].Value = "in X (mm)";
            sheet.Cells["D9"].Value = "in Y (mm)";
            sheet.Cells["D10"].Value = "Rejection criterias:";
            sheet.Cells["D10:E10"].Merge = true;
            sheet.Cells["D11"].Value = "Min conductivity (%IACS)";
            sheet.Cells["D12"].Value = "Max conductivity (%IACS)";
            sheet.Cells["D13"].Value = "Conductivity range at plate (%IACS)";
            sheet.Cells["D14"].Value = "Conductivity range in the part (%IACS)";

            sheet.Cells["E4:E7"].Style.Fill.PatternType = ExcelFillStyle.LightGray;
            sheet.Cells["E4:E7"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightSlateGray);
            sheet.Cells["E8:E9"].Style.Fill.PatternType = ExcelFillStyle.LightGray;
            sheet.Cells["E8:E9"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightSlateGray);
            sheet.Cells["E10:E14"].Style.Fill.PatternType = ExcelFillStyle.LightGray;
            sheet.Cells["E10:E14"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightSlateGray);
            //Temperature Data
            sheet.Cells["G3"].Value = "Temperature Data";
            sheet.Cells["G3:H3"].Merge = true;
            sheet.Cells["G4"].Value = "Calibration temperature Min (°C)";
            sheet.Cells["G5"].Value = "Calibration temperature Max (°C)";
            sheet.Cells["G6"].Value = "Scan temperature Min (°C)";
            sheet.Cells["G7"].Value = "Scan temperature Max (°C)";
            sheet.Cells["G8"].Value = "Range (°C)";

            sheet.Cells["H4:H8"].Style.Fill.PatternType = ExcelFillStyle.LightGray;
            sheet.Cells["H4:H8"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightSlateGray);
            //Measured values Tabels
            sheet.Cells["A16"].Value = "Measured values: Tabels, Paste the Tabels in the file and complete the next fields";
            sheet.Cells["A17"].Value = "for Top Measured Table (Top Left Corner)";
            sheet.Cells["A18"].Value = "for Top Measured Table (Bottom Right Corner)";
            sheet.Cells["A19"].Value = "for Bottom Measured Table (Top Left Corner)";
            sheet.Cells["A20"].Value = "for Bottom Measured Table (Bottom Right Corner)";
            sheet.Cells["B17:B20"].Style.Fill.PatternType = ExcelFillStyle.LightGray;
            sheet.Cells["B17:B20"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightSlateGray);
            sheet.Cells.AutoFitColumns();
            string fileName = "sample" + generateDate() + ".xlsx";
            pack.SaveAs(new FileInfo(Path.Combine(path, fileName)));
            return fileName;

        }

        public bool testIntegrity(string file)
        {
            var pack = new ExcelPackage(new FileInfo(file));
            var sheet = pack.Workbook.Worksheets.First();
            var val = sheet.Cells["Z1"].Value;
            if(val != null)
            {
                val = val.ToString();
                return val.Equals("keySM");
            }
            else
            {
                return false;
            }
            
        }
        public static string generateDate()
        {
            var date = DateTime.Now;
            string day = date.Hour.ToString()+"."+ date.Minute.ToString() + "." + date.Second.ToString();
            return day;
        }


    }
}
