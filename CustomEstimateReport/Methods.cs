using System;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures.Model;
using Excel = Microsoft.Office.Interop.Excel;

namespace CustomEstimateReport
{
    class Methods
    {
        public static IList<ModelObject> ToList(ModelObjectEnumerator enumerator)
        {
            var modelObjects = new List<ModelObject>();

            while (enumerator.MoveNext())
            {
                var modelObject = enumerator.Current;
                if (modelObject == null) continue;

                modelObjects.Add(modelObject);
            }

            return modelObjects;
        }

        public static int writeDataToExcel(Dictionary<String, Slab> data, Excel.Worksheet sheet, int startRow)
        {
            int rowindex = startRow;

            foreach (string key in data.Keys)
            {
                Excel.Range formatRange = sheet.Range["a" + rowindex.ToString(), "c" + rowindex.ToString()];
                formatRange.Interior.Color = ColorTranslator.ToOle(Color.YellowGreen);
                sheet.Cells[rowindex, 1] = data[key].name;
                sheet.Cells[rowindex, 2] = "Quantity";
                sheet.Cells[rowindex, 3] = "Unit";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Material";
                sheet.Cells[rowindex, 2] = data[key].material;
                sheet.Cells[rowindex, 3] = "Concrete";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Gross Volume";
                sheet.Cells[rowindex, 2] = data[key].volumeGross;
                sheet.Cells[rowindex, 3] = "Cubic Yard";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Net Volume";
                sheet.Cells[rowindex, 2] = data[key].volumeNet;
                sheet.Cells[rowindex, 3] = "Cubic Yard";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Top Area";
                sheet.Cells[rowindex, 2] = data[key].areaTop;
                sheet.Cells[rowindex, 3] = "Square Foot";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Bottom Area";
                sheet.Cells[rowindex, 2] = data[key].areaBott;
                sheet.Cells[rowindex, 3] = "Square Foot";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Edge Area";
                sheet.Cells[rowindex, 2] = data[key].areaEdge;
                sheet.Cells[rowindex, 3] = "Square Foot";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Perimeter";
                sheet.Cells[rowindex, 2] = data[key].perimeter;
                sheet.Cells[rowindex, 3] = "Foot";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Quantity";
                sheet.Cells[rowindex, 2] = data[key].quantity;
                sheet.Cells[rowindex, 3] = "";
                rowindex++;
            }

            startRow = rowindex;

            return startRow;
        }

        public static int writeDataToExcel(Dictionary<String, Beam> data, Excel.Worksheet sheet, int startRow)
        {
            int rowindex = startRow;

            foreach (string key in data.Keys)
            {

                Excel.Range formatRange = sheet.Range["a" + rowindex.ToString(), "c" + rowindex.ToString()];
                formatRange.Interior.Color = ColorTranslator.ToOle(Color.YellowGreen);
                sheet.Cells[rowindex, 1] = data[key].name;
                sheet.Cells[rowindex, 2] = "Quantity";
                sheet.Cells[rowindex, 3] = "Unit";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Material";
                sheet.Cells[rowindex, 2] = data[key].material;
                sheet.Cells[rowindex, 3] = "Concrete";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Gross Volume";
                sheet.Cells[rowindex, 2] = data[key].volumeGross;
                sheet.Cells[rowindex, 3] = "Cubic Yard";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Length";
                sheet.Cells[rowindex, 2] = data[key].length;
                sheet.Cells[rowindex, 3] = "Foot";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Bottom Area";
                sheet.Cells[rowindex, 2] = data[key].areaBott;
                sheet.Cells[rowindex, 3] = "Square Foot";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Side Area";
                sheet.Cells[rowindex, 2] = data[key].areaSide;
                sheet.Cells[rowindex, 3] = "Square Foot";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Quantity";
                sheet.Cells[rowindex, 2] = data[key].quantity;
                sheet.Cells[rowindex, 3] = "";
                rowindex++;
            }

            startRow = rowindex;

            return startRow;

        }

        public static int writeDataToExcel(Dictionary<String, Column> data, Excel.Worksheet sheet, int startRow)
        {
            int rowindex = startRow;

            foreach (string key in data.Keys)
            {

                Excel.Range formatRange = sheet.Range["a" + rowindex.ToString(), "c" + rowindex.ToString()];
                formatRange.Interior.Color = ColorTranslator.ToOle(Color.YellowGreen);
                sheet.Cells[rowindex, 1] = data[key].name;
                sheet.Cells[rowindex, 2] = "Quantity";
                sheet.Cells[rowindex, 3] = "Unit";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Material";
                sheet.Cells[rowindex, 2] = data[key].material;
                sheet.Cells[rowindex, 3] = "Concrete";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Gross Volume";
                sheet.Cells[rowindex, 2] = data[key].volumeGross;
                sheet.Cells[rowindex, 3] = "Cubic Yard";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Height";
                sheet.Cells[rowindex, 2] = data[key].height;
                sheet.Cells[rowindex, 3] = "Foot";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Side Area";
                sheet.Cells[rowindex, 2] = data[key].areaSide;
                sheet.Cells[rowindex, 3] = "Square Foot";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Quantity";
                sheet.Cells[rowindex, 2] = data[key].quantity;
                sheet.Cells[rowindex, 3] = "";
                rowindex++;
            }

            startRow = rowindex;

            return startRow;
        }

        public static int writeDataToExcel(Dictionary<String, Footing> data, Excel.Worksheet sheet, int startRow)
        {
            int rowindex = startRow;

            foreach (string key in data.Keys)
            {

                Excel.Range formatRange = sheet.Range["a" + rowindex.ToString(), "c" + rowindex.ToString()];
                formatRange.Interior.Color = ColorTranslator.ToOle(Color.YellowGreen);
                sheet.Cells[rowindex, 1] = data[key].name;
                sheet.Cells[rowindex, 2] = "Quantity";
                sheet.Cells[rowindex, 3] = "Unit";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Material";
                sheet.Cells[rowindex, 2] = data[key].material;
                sheet.Cells[rowindex, 3] = "Concrete";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Gross Volume";
                sheet.Cells[rowindex, 2] = data[key].volumeGross;
                sheet.Cells[rowindex, 3] = "Cubic Yard";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Top Area";
                sheet.Cells[rowindex, 2] = data[key].areaTop;
                sheet.Cells[rowindex, 3] = "Square Foot";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Bottom Area";
                sheet.Cells[rowindex, 2] = data[key].areaBott;
                sheet.Cells[rowindex, 3] = "Square Foot";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Side Area";
                sheet.Cells[rowindex, 2] = data[key].areaEdge;
                sheet.Cells[rowindex, 3] = "Square Foot";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Quantity";
                sheet.Cells[rowindex, 2] = data[key].quantity;
                sheet.Cells[rowindex, 3] = "";
                rowindex++;
            }

            startRow = rowindex;

            return startRow;
        }

        public static int writeDataToExcel(Dictionary<String, Wall> data, Excel.Worksheet sheet, int startRow)
        {
            int rowindex = startRow;

            foreach (string key in data.Keys)
            {
                Excel.Range formatRange = sheet.Range["a" + rowindex.ToString(), "c" + rowindex.ToString()];
                formatRange.Interior.Color = ColorTranslator.ToOle(Color.YellowGreen);
                sheet.Cells[rowindex, 1] = data[key].name;
                sheet.Cells[rowindex, 2] = "Quantity";
                sheet.Cells[rowindex, 3] = "Unit";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Material";
                sheet.Cells[rowindex, 2] = data[key].material;
                sheet.Cells[rowindex, 3] = "Concrete";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Gross Volume";
                sheet.Cells[rowindex, 2] = data[key].volumeGross;
                sheet.Cells[rowindex, 3] = "Cubic Yard";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Net Volume";
                sheet.Cells[rowindex, 2] = data[key].volumeNet;
                sheet.Cells[rowindex, 3] = "Cubic Yard";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Length";
                sheet.Cells[rowindex, 2] = data[key].length;
                sheet.Cells[rowindex, 3] = "Foot";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Top Area";
                sheet.Cells[rowindex, 2] = data[key].areaTop;
                sheet.Cells[rowindex, 3] = "Square Foot";
                rowindex++;

                sheet.Cells[rowindex, 1] = "End1 Area";
                sheet.Cells[rowindex, 2] = data[key].areaEnd1;
                sheet.Cells[rowindex, 3] = "Square Foot";
                rowindex++;

                sheet.Cells[rowindex, 1] = "End2 Area";
                sheet.Cells[rowindex, 2] = data[key].areaEnd2;
                sheet.Cells[rowindex, 3] = "Square Foot";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Side1 Area";
                sheet.Cells[rowindex, 2] = data[key].areaSide1;
                sheet.Cells[rowindex, 3] = "Square Foot";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Side2 Area";
                sheet.Cells[rowindex, 2] = data[key].areaSide2;
                sheet.Cells[rowindex, 3] = "Square Foot";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Gross Side Area";
                sheet.Cells[rowindex, 2] = data[key].areaSideGross;
                sheet.Cells[rowindex, 3] = "Square Foot";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Opening Area";
                sheet.Cells[rowindex, 2] = data[key].areaOpening;
                sheet.Cells[rowindex, 3] = "Square Foot";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Quantity";
                sheet.Cells[rowindex, 2] = data[key].quantity;
                sheet.Cells[rowindex, 3] = "";
                rowindex++;
            }

            startRow = rowindex;

            return startRow;
        }

        public static int writeDataToExcel(Dictionary<String, ContinuousFooting> data, Excel.Worksheet sheet, int startRow)
        {
            int rowindex = startRow;

            foreach (string key in data.Keys)
            {

                Excel.Range formatRange = sheet.Range["a" + rowindex.ToString(), "c" + rowindex.ToString()];
                formatRange.Interior.Color = ColorTranslator.ToOle(Color.YellowGreen);
                sheet.Cells[rowindex, 1] = data[key].name;
                sheet.Cells[rowindex, 2] = "Quantity";
                sheet.Cells[rowindex, 3] = "Unit";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Material";
                sheet.Cells[rowindex, 2] = data[key].material;
                sheet.Cells[rowindex, 3] = "Concrete";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Gross Volume";
                sheet.Cells[rowindex, 2] = data[key].volumeGross;
                sheet.Cells[rowindex, 3] = "Cubic Yard";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Length";
                sheet.Cells[rowindex, 2] = data[key].length;
                sheet.Cells[rowindex, 3] = "Foot";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Top Area";
                sheet.Cells[rowindex, 2] = data[key].areaTop;
                sheet.Cells[rowindex, 3] = "Square Foot";
                rowindex++;

                sheet.Cells[rowindex, 1] = "End1 Area";
                sheet.Cells[rowindex, 2] = data[key].areaEnd1;
                sheet.Cells[rowindex, 3] = "Square Foot";
                rowindex++;

                sheet.Cells[rowindex, 1] = "End2 Area";
                sheet.Cells[rowindex, 2] = data[key].areaEnd2;
                sheet.Cells[rowindex, 3] = "Square Foot";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Side1 Area";
                sheet.Cells[rowindex, 2] = data[key].areaSide1;
                sheet.Cells[rowindex, 3] = "Square Foot";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Side2 Area";
                sheet.Cells[rowindex, 2] = data[key].areaSide2;
                sheet.Cells[rowindex, 3] = "Square Foot";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Quantity";
                sheet.Cells[rowindex, 2] = data[key].quantity;
                sheet.Cells[rowindex, 3] = "";
                rowindex++;
            }

            startRow = rowindex;

            return startRow;
        }

        public static int writeDataToExcel(Dictionary<String, Styrofoam> data, Excel.Worksheet sheet, int startRow)
        {
            int rowindex = startRow;

            foreach (string key in data.Keys)
            {
                Excel.Range formatRange = sheet.Range["a" + rowindex.ToString(), "c" + rowindex.ToString()];
                formatRange.Interior.Color = ColorTranslator.ToOle(Color.YellowGreen);
                sheet.Cells[rowindex, 1] = data[key].name;
                sheet.Cells[rowindex, 2] = "Quantity";
                sheet.Cells[rowindex, 3] = "Unit";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Material";
                sheet.Cells[rowindex, 2] = data[key].material;
                sheet.Cells[rowindex, 3] = "Concrete";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Gross Volume";
                sheet.Cells[rowindex, 2] = data[key].volumeGross;
                sheet.Cells[rowindex, 3] = "Cubic Yard";
                rowindex++;

                sheet.Cells[rowindex, 1] = "Quantity";
                sheet.Cells[rowindex, 2] = data[key].quantity;
                sheet.Cells[rowindex, 3] = "";
                rowindex++;
            }

            startRow = rowindex;

            return startRow;
        }

        /* --------------------- Method For Generating An Excel File Out of A Grid --------------------- */
        public static void exportToExcel(Dictionary<String, Slab> slabs, Dictionary<String, Beam> beams,
            Dictionary<String, Column> columns, Dictionary<String, Footing> footings, Dictionary<String, Wall> walls,
            Dictionary<String, ContinuousFooting> continuousFootings, Dictionary<String, Styrofoam> styrofoam)
        {

            SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            saveFileDialog.Filter = "Excel |*.xlsx";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string outputPath = saveFileDialog.FileName;

                Excel.Application excel = new Excel.Application();
                Excel.Workbook workBook = excel.Workbooks.Add(Type.Missing);
                Excel.Worksheet workSheet = (Excel.Worksheet)workBook.ActiveSheet;

                int rowIndex = writeDataToExcel(footings, workSheet, 1);
                rowIndex = writeDataToExcel(continuousFootings, workSheet, rowIndex);
                rowIndex = writeDataToExcel(slabs, workSheet, rowIndex);
                rowIndex = writeDataToExcel(columns, workSheet, rowIndex);
                rowIndex = writeDataToExcel(beams, workSheet, rowIndex);
                rowIndex = writeDataToExcel(walls, workSheet, rowIndex);
                rowIndex = writeDataToExcel(styrofoam, workSheet, rowIndex);

                Excel.Range formatRange = workSheet.UsedRange;
                formatRange.EntireColumn.AutoFit();
                formatRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                workBook.SaveAs(outputPath);
                workBook.Close();
                excel.Quit();

            }
        }
    }
}


