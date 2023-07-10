using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ConsoleApp1
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
            lines=File.ReadAllLines(filePath).ToList();
            List<string> ListInputs = new List<string>();
            try
            {
                foreach (string line in lines)
                {
                    int startidx=line.IndexOf(" ");
                    int endidx=line.IndexOf("n");
                    string strTmp=line.Substring(startidx+1,endidx-startidx);
                    ListInputs.Add(strTmp);
                }
            }
            catch(Exception e)
            {
                Console.WriteLine(e);
            }


            return ListInputs;
        }
    }

    //Substitute the "n" to int value "1-4" i  the RCAI/TI VID list
    public class Substitute
    {
        public static List<string> SubstituteN(List<string> Inputs,int n)
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
        public static List<string> GET_WHC_AI_SAI(int VID_Start, int start,int end)
        {
            List<string> WHCInputs = new List<string>();
            int VID = 0;
            for (int i = 0; i<start-end+1;i++)
            {
                VID = VID_Start + 16 * i;
                WHCInputs.Add(VID.ToString());
            }
            return WHCInputs;
        }
    }

    public class Worksheets
    {
        //Initialize the Titles in the worksheet;
        //InputType has 3 optional value: "RCAI", "RCTI" and "WHCAI"
        public static string[] init(string InputType){

            string[] Title=new string[1024];

            if(InputType == "RCAI"){
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
            if(InputType == "RCTI"){
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

            if (InputType == "WHCAI"){
                
            }

            return Title;
        }

        public static void Create_RCAI_Sheet(string filePath,string fileName,string[] Title)
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath+fileName))){
                try
                {
                    package.Workbook.Worksheets.Delete("Template");
                }
                catch(Exception e)
                {
                    Console.WriteLine(e.Message);
                }

                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Template");
                int i=0;
                int title_row = 3;
                while(Title[i] != null){
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
        
        public static void Insert_IDs(string filePath,string fileName,int col,List<string> IDs){
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            int start_row = 4;

            using(ExcelPackage package=new ExcelPackage(new FileInfo(filePath+fileName))){
                try{
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Template"];
                    int i=0;
                    foreach(string ID in IDs){
                        worksheet.Cells[start_row+i,col].Value=ID;
                        worksheet.Cells[start_row+i,col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[start_row+i,col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        worksheet.Cells[start_row+i,col].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        i++;
                    }
                    
                }
                catch(Exception e){
                    Console.WriteLine(e.Message);
                }
                package.Save();
            }
            
        }

        public static void CreateConfigSheet()
        {
            
        }
    }


    public class Program
    {
        // public static void Main(string[] args)
        // {
        //     string str = "D:\\TemplateTranslate\\Manual_Original\\XP4_RCAI.txt";
        //     List<string> lines = new List<string>();
        //     lines=ReadInputs.ReadFile(str);
        //     foreach (string line in lines)
        //     {
        //         Console.WriteLine(line);
        //     }
        //     Console.WriteLine("-----------------------endl----------------------");
            
        //     lines = Substitute.SubstituteN(lines, 1);
        //     foreach(string line in lines)
        //     {
        //         Console.WriteLine(line);
        //     }

        //     Console.WriteLine("-----------------------endl----------------------");

        //     lines = HexToDecimalConverter.ConvertHexListToDecimal(lines);
        //     foreach (string line in lines)
        //     {
        //         Console.WriteLine(line);
        //     }

        //     Console.WriteLine("-----------------------endl----------------------");

        //     lines = Seperate.SeperateCols(str, 4);
        //     foreach (string line in lines)
        //     {
        //         Console.WriteLine(line);
        //     }

        //     Console.WriteLine("-----------------------endl----------------------");
        //     string filePath="D:\\TemplateTranslate\\Template\\XP4";
        //     string fileName = "\\RCAI.xlsx";
        //     int row = lines.Count();

        //     string[] Title=new string[1024];
        //     string Type = "RCAI";
        //     Title =Worksheets.init(Type);
        //     try
        //     {
        //         Worksheets.Create_RCAI_Sheet(filePath,fileName,Title);
        //     }
        //     catch (Exception e)
        //     {
        //         Console.WriteLine(e.Message);
        //     }

        //     Console.WriteLine("-----------------------endl----------------------");
        //     try
        //     {
        //         Worksheets.Insert_IDs(filePath, fileName, 4, lines);
        //     }
        //     catch (Exception e)
        //     {
        //         Console.WriteLine(e.Message);
        //     }
            

        // }
        public static void Main(string[] args)
        {
            string Mannual = "D:\\TemplateTranslate\\Manual_Original\\XP8_DCMAI.txt";
            //Actual input for str To be Modified
            string TemplatePath= "D:\\TemplateTranslate\\Template\\XP8";
            string TemplateName= "\\DCMAI.xlsx";

            string[] Title = new string[1024];
            string Type = "DCMAI";
            Title = Worksheets.init(Type);
            Worksheets.Create_RCAI_Sheet(TemplatePath, TemplateName, Title);

            List<string> AIs = new List<string>();
            AIs = ReadInputs.ReadFile(Mannual);

            List<string> AI_Temp = new List<string>();

            int col = 0;
            for (int i = 0; i < 4; i++)
            {
                AI_Temp = Substitute.SubstituteN(AIs, i + 1);
                AI_Temp = HexToDecimalConverter.ConvertHexListToDecimal(AI_Temp);
                col = (i + 1) * 3 + 1;
                Worksheets.Insert_IDs(TemplatePath, TemplateName, col, AI_Temp);

            }

            List<string> Numbers = new List<string>();
            Numbers = Seperate.SeperateCols(Mannual, 0);

            for (int i = 0; i < Numbers.Count; i++)
            {
                Numbers[i] = "AI" + Numbers[i];
            }


            Worksheets.Insert_IDs(TemplatePath, TemplateName, 1, Numbers);

            for (int i = 0; i < 5; i++)
            {
                AI_Temp = Seperate.SeperateCols(Mannual, 3 + i);
                AI_Temp = HexToDecimalConverter.ConvertHexListToDecimal(AI_Temp);
                Worksheets.Insert_IDs(TemplatePath, TemplateName, col + i + 1, AI_Temp);
            }


        }
    }
}
