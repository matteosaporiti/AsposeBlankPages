using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AsposeBlankPages
{
    class Program
    {
        private static readonly string ProjDirectory = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
        static void Main(string[] args)
        {
            var bytes = File.ReadAllBytes(ProjDirectory + @"\docs\Test_Source.docx");
            var attachments = new Dictionary<string, byte[]>
            {
                {"Att1.docx", File.ReadAllBytes(ProjDirectory + @"\docs\Test_Attachments\Att1.docx")},
                {"Att2.docx", File.ReadAllBytes(ProjDirectory + @"\docs\Test_Attachments\Att2.docx")},
                {"Att3.docx", File.ReadAllBytes(ProjDirectory + @"\docs\Test_Attachments\Att3.docx")},
                {"Att4.docx", File.ReadAllBytes(ProjDirectory+ @"\docs\Test_Attachments\Att4.docx") },
                {"Att5.docx", File.ReadAllBytes(ProjDirectory + @"\docs\Test_Attachments\Att5.docx")},
                {"Att6.docx", File.ReadAllBytes(ProjDirectory + @"\docs\Test_Attachments\Att6.docx")},
                {"Att7.docx", File.ReadAllBytes(ProjDirectory + @"\docs\Test_Attachments\Att7.docx")},
                {"Att8.docx", File.ReadAllBytes(ProjDirectory + @"\docs\Test_Attachments\Att8.docx") }
            };

            MergeManager.MergeWithAspose(bytes, attachments);
        }
    }
}
