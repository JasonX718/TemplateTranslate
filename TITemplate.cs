using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TemplateTranslateApp
{
    public class TITemplate
    {
        public static void PrintTemplate(string Mannual, string TemplatePath, string TemplateName, string Type)
        {

            string[] Title = new string[1024];
            Title = Worksheets.init(Type);
            Worksheets.CreateSheet(TemplatePath, TemplateName, Title);

            List<string> TIs = new List<string>();
            TIs = ReadInputs.ReadFile(Mannual);

            List<string> TI_Temp = new List<string>();

            int col = 0;
            for (int i = 0; i < 4; i++)
            {
                TI_Temp = Substitute.SubstituteN(TIs, i + 1);
                TI_Temp = HexToDecimalConverter.ConvertHexListToDecimal(TI_Temp);
                col = (i + 1) * 3 + 1;
                Worksheets.Insert_IDs(TemplatePath, TemplateName, col, TI_Temp);

            }

            List<string> Numbers = new List<string>();
            Numbers = Seperate.SeperateCols(Mannual, 0);

            for (int i = 0; i < Numbers.Count; i++)
            {
                Numbers[i] = "TI" + Numbers[i];
            }


            Worksheets.Insert_IDs(TemplatePath, TemplateName, 1, Numbers);

        }

        public static void AssignConfig(string config, string TemplatePath, string TemplateName, bool IsSynergis,int ChamberNum)
        {

            List<string> Table = ReadConfig.ReadTITable(config, IsSynergis);
            var Result = ReadConfig.LoadData(Table);
            List<string> TIPresent = Result.Item1;
            List<string> TIName = Result.Item2;
            List<string> TIUnit = Result.Item3;

            Assign.RC(TIPresent, TIName, TIUnit, TemplatePath, TemplateName,ChamberNum);
        }
    }
}
