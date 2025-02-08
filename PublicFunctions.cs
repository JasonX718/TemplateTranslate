using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;


namespace TemplateTranslateApp
{
    //Convert Hexadecimal data to Decimal
    public class HexToDecimalConverter
    {
        public static List<string> ConvertHexListToDecimal(List<string> hexList)
        {
            // Initialize an empty list to store the decimal conversions
            List<int> decimalList = new List<int>();

            // Iterate over each string in the input list
            foreach (string hex in hexList)
            {
                // Convert the hex string to a decimal and add to the output list
                int decimalValue = Convert.ToInt32(hex, 16);
                decimalList.Add(decimalValue);
            }
            List<string> StrList = new List<string>();

            foreach (int dec in decimalList)
            {
                StrList.Add(dec.ToString());
            }

            return StrList;
        }
    }

    //To Get RCAI/TI VID list
    public class ReadInputs
    {
        public static List<string> ReadFile(string filePath)
        {
            List<string> lines = new List<string>();
            lines = File.ReadAllLines(filePath).ToList();
            List<string> ListInputs = new List<string>();
            try
            {
                foreach (string line in lines)
                {
                    int startidx = line.IndexOf(" ");
                    int endidx = line.IndexOf("n");
                    string strTmp = line.Substring(startidx + 1, endidx - startidx);
                    ListInputs.Add(strTmp);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }


            return ListInputs;
        }
    }

    //Substitute the "n" to int value "1-4" i  the RCAI/TI VID list
    public class Substitute
    {
        public static List<string> SubstituteN(List<string> Inputs, int n)
        {
            List<string> RCnInputs = new List<string>();
            foreach (string Input in Inputs)
            {
                int idx = Input.IndexOf("n");
                char[] charArray = Input.ToCharArray();
                charArray[idx] = Convert.ToChar(n.ToString());
                RCnInputs.Add(new string(charArray));

            }

            return RCnInputs;
        }
    }

    //Seperate the Columns in the txt format AI/TI tables
    public class Seperate
    {
        public static List<string> SeperateCols(string filePath, int colNum)
        {
            List<string> lines = File.ReadAllLines(filePath).ToList();
            List<string> ListInputs = new List<string>();

            foreach (string line in lines)
            {
                string[] cols = line.Split(' ');
                ListInputs.Add(cols[colNum]);

            }

            return ListInputs;

        }
    }

    //To get WHCAI and WHCAI Summary
    public class GET_WHC
    {
        public static List<string> GET_WHC_AI_SAI(int VID_Start, int start, int end)
        {
            List<string> WHCInputs = new List<string>();
            int VID = 0;
            for (int i = 0; i < start - end + 1; i++)
            {
                VID = VID_Start + 16 * i;
                WHCInputs.Add(VID.ToString());
            }
            return WHCInputs;
        }

        public static void GET_XP4_AI(string filePath, string fileName)
        {
            List<string> WHCAI = new List<string>();
            int AI_Temp = 0;
            for (int i = 0; i < 16; i++)
            {
                AI_Temp = 34155520 + 16 * i;
                WHCAI.Add(AI_Temp.ToString());
            }
            for (int i = 0; i < 16; i++)
            {
                AI_Temp = 34978048 + 16 * i;
                WHCAI.Add(AI_Temp.ToString());
            }
            List<string> Nums = new List<string>();
            for (int i = 0; i < 32; i++)
            {
                Nums.Add("AI" + i.ToString("D2"));
            }
            Worksheets.Insert_IDs(filePath, fileName, 1, Nums);
            Worksheets.Insert_IDs(filePath, fileName, 4, WHCAI);
        }

        public static void GET_XP8_AI(string filePath, string fileName)
        {
            List<string> WHCAI = new List<string>();
            int AI_Temp = 0;
            for (int i = 0; i < 16; i++)
            {
                AI_Temp = 33883136 + 16 * i;
                WHCAI.Add(AI_Temp.ToString());
            }
            for (int i = 0; i < 16; i++)
            {
                AI_Temp = 33892608 + 16 * i;
                WHCAI.Add(AI_Temp.ToString());
            }
            for (int i = 0; i < 16; i++)
            {
                AI_Temp = 38928384 + 16 * i;
                WHCAI.Add(AI_Temp.ToString());
            }
            List<string> Nums = new List<string>();
            for (int i = 0; i < 48; i++)
            {
                Nums.Add("AI" + i.ToString("D2"));
            }
            Worksheets.Insert_IDs(filePath, fileName, 1, Nums);
            Worksheets.Insert_IDs(filePath, fileName, 4, WHCAI);
        }

        public static void GET_Synergis_AI(string filePath, string fileName)
        {
            List<string> WHCAI = new List<string>();
            int AI_Temp = 0;
            for (int i = 0; i < 16; i++)
            {
                AI_Temp = 33883136 + 16 * i;
                WHCAI.Add(AI_Temp.ToString());
            }
            for (int i = 0; i < 16; i++)
            {
                AI_Temp = 33892608 + 16 * i;
                WHCAI.Add(AI_Temp.ToString());
            }
            for (int i=0;i<288;i++)
            {
                AI_Temp = 36848640 + 16 * i;
                WHCAI.Add(AI_Temp.ToString());
            }
            for (int i = 0; i < 80; i++)
            {
                AI_Temp = 36853248 + 16 * i;
                WHCAI.Add(AI_Temp.ToString());
            }
            List<string> Nums = new List<string>();
            for (int i = 0; i < 400; i++)
            {
                Nums.Add("AI" + i.ToString("D2"));
            }
            Worksheets.Insert_IDs(filePath, fileName, 1, Nums);
            Worksheets.Insert_IDs(filePath, fileName, 4, WHCAI);
        }

        public static void GET_Intrepid_AI(string filePath, string fileName)
        {
            List<string> WHCAI = new List<string>();
            int AI_Temp = 0;
            for (int i = 0; i < 16; i++)
            {
                AI_Temp = 34155520 + 16 * i;
                WHCAI.Add(AI_Temp.ToString());
            }
            for (int i = 0; i < 16; i++)
            {
                AI_Temp = 34978048 + 16 * i;
                WHCAI.Add(AI_Temp.ToString());
            }
            List<string> Nums = new List<string>();
            for (int i = 0; i < 32; i++)
            {
                Nums.Add("AI" + i.ToString("D2"));
            }
            Worksheets.Insert_IDs(filePath, fileName, 1, Nums);
            Worksheets.Insert_IDs(filePath, fileName, 4, WHCAI);
        }

        public static void GET_XP4_SAI(string filePath, string fileName)
        {
            List<string> WHCAI = new List<string>();
            int AI_Temp = 0;
            for (int i = 0; i < 32; i++)
            {
                AI_Temp = 34155776 + 16 * i;
                WHCAI.Add(AI_Temp.ToString());
            }
            List<string> Nums = new List<string>();
            for (int i = 0; i < 32; i++)
            {
                Nums.Add("SAI" + i.ToString("D2"));
            }
            Worksheets.Insert_IDs(filePath, fileName, 1, Nums);
            Worksheets.Insert_IDs(filePath, fileName, 4, WHCAI);
        }

        public static void GET_XP8_SAI(string filePath, string fileName)
        {
            List<string> WHCAI = new List<string>();
            int AI_Temp = 0;
            for (int i = 0; i < 32; i++)
            {
                AI_Temp = 33883648 + 16 * i;
                WHCAI.Add(AI_Temp.ToString());
            }
            for (int i = 0; i < 32; i++)
            {
                AI_Temp = 33883650 + 16 * i;
                WHCAI.Add(AI_Temp.ToString());
            }
            List<string> Nums = new List<string>();
            for (int i = 0; i < 64; i++)
            {
                Nums.Add("SAI" + i.ToString("D2"));
            }
            Worksheets.Insert_IDs(filePath, fileName, 1, Nums);
            Worksheets.Insert_IDs(filePath, fileName, 4, WHCAI);
        }

        public static void GET_Synergis_SAI(string filePath, string fileName)
        {
            List<string> WHCAI = new List<string>();
            int AI_Temp = 0;
            for (int i = 0; i < 32; i++)
            {
                AI_Temp = 33883694 + 16 * i;
                WHCAI.Add(AI_Temp.ToString());
            }
            for (int i = 0; i < 288; i++)
            {
                AI_Temp = 36860416 + 16 * i;
                WHCAI.Add(AI_Temp.ToString());
            }
            for (int i = 0; i < 80; i++)
            {
                AI_Temp = 36865024 + 16 * i;
                WHCAI.Add(AI_Temp.ToString());
            }
            List<string> Nums = new List<string>();
            for (int i = 0; i < 400; i++)
            {
                Nums.Add("SAI" + i.ToString("D2"));
            }
            Worksheets.Insert_IDs(filePath, fileName, 1, Nums);
            Worksheets.Insert_IDs(filePath, fileName, 4, WHCAI);
        }

        public static void GET_Intrepid_SAI(string filePath, string fileName)
        {
            List<string> WHCAI = new List<string>();
            int AI_Temp = 0;
            for (int i = 0; i < 32; i++)
            {
                AI_Temp = 34155776 + 16 * i;
                WHCAI.Add(AI_Temp.ToString());
            }
            List<string> Nums = new List<string>();
            for (int i = 0; i < 32; i++)
            {
                Nums.Add("SAI" + i.ToString("D2"));
            }
            Worksheets.Insert_IDs(filePath, fileName, 1, Nums);
            Worksheets.Insert_IDs(filePath, fileName, 4, WHCAI);
        }

    }

    public class Worksheets
    {
        //Initialize the Titles in the worksheet;
        //InputType has 3 optional value: "RCAI", "RCTI" and "WHCAI"
        public static string[] init(string InputType)
        {

            string[] Title = new string[1024];

            if (InputType == "RCAI")
            {
                Title[0] = "AI number";
                Title[1] = "RC1 AIname";
                Title[2] = "RC1 Unit";
                Title[3] = "RC1 ActualData(SVID)";
                Title[4] = "RC2 AIname";
                Title[5] = "RC2 Unit";
                Title[6] = "RC2 ActualData(SVID)";
                Title[7] = "RC3 AIname";
                Title[8] = "RC3 Unit";
                Title[9] = "RC3 ActualData(SVID)";
                Title[10] = "RC4 AIname";
                Title[11] = "RC4 Unit";
                Title[12] = "RC4 ActualData(SVID)";
                Title[13] = "Set(DVID)";
                Title[14] = "mean(DVID)";
                Title[15] = "max(DVID)";
                Title[16] = "min(DVID)";
                Title[17] = "deviation(DVID)";

            }
            if (InputType == "RCTI")
            {
                Title[0] = "TI number";
                Title[1] = "RC1 TIname";
                Title[2] = "RC1 Unit";
                Title[3] = "RC1 ActualData(SVID)";
                Title[4] = "RC2 TIname";
                Title[5] = "RC2 Unit";
                Title[6] = "RC2 ActualData(SVID)";
                Title[7] = "RC3 TIname";
                Title[8] = "RC3 Unit";
                Title[9] = "RC3 ActualData(SVID)";
                Title[10] = "RC4 TIname";
                Title[11] = "RC4 Unit";
                Title[12] = "RC4 ActualData(SVID)";

            }
            if (InputType == "DCMAI")
            {
                Title[0] = "AI number";
                Title[1] = "DCM1 AIname";
                Title[2] = "DCM1 Unit";
                Title[3] = "DCM1 ActualData(SVID)";
                Title[4] = "DCM2 AIname";
                Title[5] = "DCM2 Unit";
                Title[6] = "DCM2 ActualData(SVID)";
                Title[7] = "DCM3 AIname";
                Title[8] = "DCM3 Unit";
                Title[9] = "DCM3 ActualData(SVID)";
                Title[10] = "DCM4 AIname";
                Title[11] = "DCM4 Unit";
                Title[12] = "DCM4 ActualData(SVID)";
                Title[13] = "Set(DVID)";
                Title[14] = "mean(DVID)";
                Title[15] = "max(DVID)";
                Title[16] = "min(DVID)";
                Title[17] = "deviation(DVID)";

            }
            if (InputType == "DCMTI")
            {
                Title[0] = "TI number";
                Title[1] = "DCM1 TIname";
                Title[2] = "DCM1 Unit";
                Title[3] = "DCM1 ActualData(SVID)";
                Title[4] = "DCM2 TIname";
                Title[5] = "DCM2 Unit";
                Title[6] = "DCM2 ActualData(SVID)";
                Title[7] = "DCM3 TIname";
                Title[8] = "DCM3 Unit";
                Title[9] = "DCM3 ActualData(SVID)";
                Title[10] = "DCM4 TIname";
                Title[11] = "DCM4 Unit";
                Title[12] = "DCM4 ActualData(SVID)";

            }

            if (InputType == "WHCAI")
            {
                Title[0] = "AI Number";
                Title[1] = "WHCAI Name";
                Title[2] = "WHC Unit";
                Title[3] = "WHC ActualData(SVID)";

            }

            if (InputType == "WHCSAI")
            {
                Title[0] = "SAI Number";
                Title[1] = "WHCSAI Name";
                Title[2] = "WHC Unit";
                Title[3] = "WHC ActualData(SVID)";
            }

            return Title;
        }

        public static void CreateSheet(string filePath, string fileName, string[] Title)
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath + fileName)))
            {
                try
                {
                    package.Workbook.Worksheets.Delete("Template");
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }

                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Template");
                int i = 0;
                int title_row = 3;
                while (Title[i] != null)
                {
                    worksheet.Cells[title_row, i + 1].Value = Title[i];
                    worksheet.Cells[title_row, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[title_row, i + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[title_row, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[title_row, i + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(128, 128, 128));
                    worksheet.Cells[title_row, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Black);
                    i++;
                }
                package.Save();


            }
        }

        public static void Insert_IDs(string filePath, string fileName, int col, List<string> IDs)
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            int start_row = 4;

            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath + fileName)))
            {
                try
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Template"];
                    int i = 0;
                    foreach (string ID in IDs)
                    {
                        worksheet.Cells[start_row + i, col].Value = ID;
                        worksheet.Cells[start_row + i, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[start_row + i, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        worksheet.Cells[start_row + i, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        i++;
                    }

                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
                package.Save();
            }

        }
    }

    public class ReadConfig
    {

        public static List<string> ReadAITable(string filePath, bool IsXP4, bool IsSynergis)
        {
            List<string> lines = new List<string>();
            lines = File.ReadAllLines(filePath).ToList();
            List<string> Table = new List<string>();
            try
            {
                int startidx = 0;
                int endidx = 0;
                int idx = 0;
                if (IsXP4)
                {
                    foreach (string line in lines)
                    {
                        if (line.Contains("Analog Input"))
                        {
                            startidx = idx + 1;
                        }
                        if (line.Contains("AI - HSE VID"))
                        {
                            endidx = idx - 1;
                            break;
                        }
                        idx++;
                    }
                }
                else if (IsSynergis)
                {
                    foreach (string line in lines)
                    {
                        if (line.Contains("// AI") && !line.Contains("Average Setting"))
                        {
                            startidx = idx;
                        }
                        if (line.Contains("// AO"))
                        {
                            endidx = idx - 1;
                            break;
                        }
                        idx++;
                    }
                }
                else
                {
                    foreach (string line in lines)
                    {
                        if (line.Contains("Analog Input"))
                        {
                            startidx = idx + 1;
                        }
                        if (line.Contains("Analog Output"))
                        {
                            endidx = idx - 1;
                            break;
                        }
                        idx++;
                    }
                }
                idx = 0;
                foreach (string line in lines)
                {
                    if (idx < startidx)
                    {
                        idx++;
                        continue;
                    }
                    else if (idx == startidx)
                    {
                        if (line.Contains("Present"))
                        {
                            Table.Add(line);
                        }
                        idx++;
                        continue;
                    }
                    else if (idx > startidx && idx < endidx)
                    {
                        Table.Add(line);
                        idx++;
                        continue;
                    }
                    else
                    {
                        break;
                    }
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            return Table;
        }

        public static List<string> ReadTITable(string filePath,bool IsSynergis)
        {
            List<string> lines = new List<string>();
            lines = File.ReadAllLines(filePath).ToList();
            List<string> Table = new List<string>();
            try
            {
                int startidx = 0;
                int endidx = 0;
                int idx = 0;
                if (IsSynergis)
                {
                    foreach (string line in lines)
                    {
                        if (line.Contains("// Temp"))
                        {
                            startidx = idx;
                        }
                        if (line.Contains("// AI") && !line.Contains("Average Setting"))
                        {
                            endidx = idx - 1;
                            break;
                        }
                        idx++;
                    }
                }
                else
                {
                    foreach (string line in lines)
                    {
                        if (line.Contains("Thermo Data"))
                        {
                            startidx = idx + 1;
                        }
                        if (line.Contains("Analog Input"))
                        {
                            endidx = idx - 1;
                            break;
                        }
                        idx++;
                    }
                }

                idx = 0;
                foreach (string line in lines)
                {
                    if (idx < startidx)
                    {
                        idx++;
                        continue;
                    }
                    else if (idx == startidx)
                    {
                        if (line.Contains("Present"))
                        {
                            Table.Add(line);
                        }
                        idx++;
                        continue;
                    }
                    else if (idx > startidx && idx < endidx)
                    {
                        Table.Add(line);
                        idx++;
                        continue;
                    }
                    else
                    {
                        break;
                    }
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            return Table;
        }

        public static List<string> ReadWHCAITable(string filePath)
        {
            List<string> lines = new List<string>();
            lines = File.ReadAllLines(filePath).ToList();
            List<string> Table = new List<string>();
            try
            {
                int startidx = 0;
                int endidx = 0;
                int idx = 0;

                foreach (string line in lines)
                {
                    if (line.Contains("Analog Input") && !line.Contains("Assign"))
                    {
                        startidx = idx + 1;
                    }
                    if (line.Contains("Analog Sensor Input"))
                    {
                        endidx = idx - 1;
                        break;
                    }
                    idx++;
                }

                idx = 0;
                foreach (string line in lines)
                {
                    if (idx < startidx)
                    {
                        idx++;
                        continue;
                    }
                    else if (idx == startidx)
                    {
                        if (line.Contains("Present"))
                        {
                            Table.Add(line);
                        }
                        idx++;
                        continue;
                    }
                    else if (idx > startidx && idx < endidx)
                    {
                        Table.Add(line);
                        idx++;
                        continue;
                    }
                    else
                    {
                        break;
                    }
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            return Table;
        }

        public static List<string> ReadWHCSAITable(string filePath)
        {
            List<string> lines = new List<string>();
            lines = File.ReadAllLines(filePath).ToList();
            List<string> Table = new List<string>();
            try
            {
                int startidx = 0;
                int endidx = 0;
                int idx = 0;

                foreach (string line in lines)
                {
                    if (line.Contains("Analog Sensor Input"))
                    {
                        startidx = idx + 1;
                    }
                    if (line.Contains("Analog Output") && !line.Contains("Assign"))
                    {
                        endidx = idx - 1;
                        break;
                    }
                    idx++;
                }

                idx = 0;
                foreach (string line in lines)
                {
                    if (idx < startidx)
                    {
                        idx++;
                        continue;
                    }
                    else if (idx == startidx)
                    {
                        if (line.Contains("Present"))
                        {
                            Table.Add(line);
                        }
                        idx++;
                        continue;
                    }
                    else if (idx > startidx && idx < endidx)
                    {
                        Table.Add(line);
                        idx++;
                        continue;
                    }
                    else
                    {
                        break;
                    }
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            return Table;
        }

        public static (List<string>, List<string>, List<string>) LoadData(List<string> Table)
        {
            List<string> Present = new List<string>();
            List<string> Name = new List<string>();
            List<string> Unit = new List<string>();
            int startidx = 0;
            int endidx = Table[0].IndexOf("Range");

            startidx = Table[0].IndexOf("Name");

            foreach (string line in Table)
            {
                string str = line.Substring(startidx, endidx - startidx - 1);
                Name.Add(str);
            }

            startidx = Table[0].IndexOf("Unit");
            string target = "Unit";

            foreach (string line in Table)
            {
                string str = line.Substring(startidx - 1, target.Length + 3);
                Unit.Add(str);
            }

            startidx = Table[0].IndexOf("Present");
            target = "Present";
            foreach (string line in Table)
            {
                string str = line.Substring(startidx, target.Length);
                str = str.Replace(" ", "");
                Present.Add(str);
            }

            return (Present, Name, Unit);
        }
    }

    public class Assign
    {
        public static void RC(List<string> Present, List<string> Name, List<string> Unit, string filePath, string fileName, int ChamberNum)
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            int start_row = 4;

            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath + fileName)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Template"];
                int len = Present.Count;
                for (int i = 1; i < len; i++)
                {
                    if (Present[i] == "1" || Present[i]=="True")
                    {
                        int j = ChamberNum;
                        worksheet.Cells[start_row + i - 1, 3 * j - 1].Value = Name[i];
                        worksheet.Cells[start_row + i - 1, 3 * j - 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[start_row + i - 1, 3 * j - 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        worksheet.Cells[start_row + i - 1, 3 * j - 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        worksheet.Cells[start_row + i - 1, 3 * j].Value = Unit[i];
                        worksheet.Cells[start_row + i - 1, 3 * j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[start_row + i - 1, 3 * j].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        worksheet.Cells[start_row + i - 1, 3 * j].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    }
                    else
                    {
                        continue;
                    }
                }
                package.Save();
            }

        }

        public static void WHC(List<string> Present, List<string> Name, List<string> Unit, string filePath, string fileName)
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            int start_row = 4;

            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath + fileName)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Template"];
                int len = Present.Count;
                for (int i = 1; i < len; i++)
                {
                    if (Present[i] == "1" || Present[i] == "True")
                    {
                        worksheet.Cells[start_row + i - 1, 2].Value = Name[i];
                        worksheet.Cells[start_row + i - 1, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[start_row + i - 1, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        worksheet.Cells[start_row + i - 1, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        worksheet.Cells[start_row + i - 1, 3].Value = Unit[i];
                        worksheet.Cells[start_row + i - 1, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[start_row + i - 1, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        worksheet.Cells[start_row + i - 1, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    }
                    else
                    {
                        continue;
                    }
                }
                package.Save();
            }
        }

    }
    class PublicFunctions
    {

    }
}
