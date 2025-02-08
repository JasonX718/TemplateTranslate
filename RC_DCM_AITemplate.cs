using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TemplateTranslateApp
{
    public class RC_DCM_AITemplate
    {
        public static void PrintTemplate(string Mannual,string TemplatePath,string TemplateName,string Type)
        {
            //string Mannual = "D:\\TemplateTranslate\\Manual_Original\\XP4_RCAI.txt";
            ////Actual input for str To be Modified
            //string TemplatePath="D:\\TemplateTranslate\\Template\\XP4";
            //string TemplateName="\\RCAI.xlsx";

            string[] Title=new string[1024];
            //string Type = "RCAI";
            Title = Worksheets.init(Type);
            Worksheets.CreateSheet(TemplatePath,TemplateName,Title);

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
        public static void AssignConfig(string config, string TemplatePath,string TemplateName,bool isXP4,bool IsSynergis,int ChamberNum)
        {
            //string TemplatePath = "D:\\TemplateTranslate\\Template\\XP4";
            //string TemplateName = "\\RCAI.xlsx";

            //string config = "C:\\Users\\m1361\\Desktop\\config\\xp8\\3E3981.03-0RR.p4f.txt";
            

            List<string> Table = ReadConfig.ReadAITable(config, isXP4, IsSynergis);
            var Result = ReadConfig.LoadData(Table);
            List<string> AIPresent = Result.Item1;
            List<string> AIName = Result.Item2;
            List<string> AIUnit = Result.Item3;

            Assign.RC(AIPresent, AIName, AIUnit, TemplatePath, TemplateName,ChamberNum);
        } 
    }
}
