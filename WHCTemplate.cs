using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TemplateTranslateApp
{
    public class WHCTemplate {
        public static void PrintTemplate(string TemplatePath, string TemplateName, string Machine, string Type)
        {
            //string TemplatePath = "D:\\TemplateTranslate\\Template\\XP8";
            //string TemplateName = "\\WHCAI.xlsx";

            if (Machine == "XP4" && Type == "WHCAI")
            {
                string[] Title = new string[1024];
                Title = Worksheets.init(Type);
                Worksheets.CreateSheet(TemplatePath, TemplateName, Title);
                GET_WHC.GET_XP4_AI(TemplatePath, TemplateName);
            }
            if (Machine == "XP8" && Type == "WHCAI")
            {
                string[] Title = new string[1024];
                Title = Worksheets.init(Type);
                Worksheets.CreateSheet(TemplatePath, TemplateName, Title);
                GET_WHC.GET_XP8_AI(TemplatePath, TemplateName);
            }
            if (Machine == "Synergis" && Type == "WHCAI")
            {
                string[] Title = new string[1024];
                Title = Worksheets.init(Type);
                Worksheets.CreateSheet(TemplatePath, TemplateName, Title);
                GET_WHC.GET_Synergis_AI(TemplatePath, TemplateName);
            }
            if (Machine == "Intrepid" && Type == "WHCAI")
            {
                string[] Title = new string[1024];
                Title = Worksheets.init(Type);
                Worksheets.CreateSheet(TemplatePath, TemplateName, Title);
                GET_WHC.GET_Intrepid_AI(TemplatePath, TemplateName);
            }
            if (Machine == "XP4" && Type == "WHCSAI")
            {
                string[] Title = new string[1024];
                Title = Worksheets.init(Type);
                Worksheets.CreateSheet(TemplatePath, TemplateName, Title);
                GET_WHC.GET_XP4_SAI(TemplatePath, TemplateName);
            }
            if (Machine == "XP8" && Type == "WHCSAI")
            {
                string[] Title = new string[1024];
                Title = Worksheets.init(Type);
                Worksheets.CreateSheet(TemplatePath, TemplateName, Title);
                GET_WHC.GET_XP8_SAI(TemplatePath, TemplateName);
            }
            if (Machine == "Synergis" && Type == "WHCSAI")
            {
                string[] Title = new string[1024];
                Title = Worksheets.init(Type);
                Worksheets.CreateSheet(TemplatePath, TemplateName, Title);
                GET_WHC.GET_Synergis_SAI(TemplatePath, TemplateName);
            }
            if (Machine == "Intrepid" && Type == "WHCSAI")
            {
                string[] Title = new string[1024];
                Title = Worksheets.init(Type);
                Worksheets.CreateSheet(TemplatePath, TemplateName, Title);
                GET_WHC.GET_Intrepid_SAI(TemplatePath, TemplateName);
            }

        }

        public static void AssignConfig(string config, string TemplatePath, string TemplateName,string Type)
        {
            List<string> Table = new List<string>();
            if (Type=="WHCSAI")
            {
                Table = ReadConfig.ReadWHCSAITable(config);
            }
            else
            {
                Table = ReadConfig.ReadWHCAITable(config);
            }
            var Result = ReadConfig.LoadData(Table);
            List<string> WHCPresent = Result.Item1;
            List<string> WHCName = Result.Item2;
            List<string> WHCUnit = Result.Item3;

            Assign.WHC(WHCPresent, WHCName, WHCUnit,TemplatePath, TemplateName);
        }
    }
}
