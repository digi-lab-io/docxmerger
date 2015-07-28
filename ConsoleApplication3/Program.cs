using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace DocxMerger
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {

                //string[] customArgs = { "C:\\tools\\DDD.docx", "C:\\tools\\test01.docx", "C:\\tools\\test02.docx" };
                //args = customArgs;

                string outputFileName = @args[0];
                List<FileStream> list = new List<FileStream>();

                foreach (var fileName in args.Skip(1))
                {
                    list.Add(File.Open(@fileName, FileMode.Open));
                }

                mergeDocx(outputFileName, list);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        static void mergeDocx(string paramOutputFile, List<FileStream> paramDocumentstreams)
        {

            var sources = new List<Source>();
            foreach (var stream in paramDocumentstreams)
            {
                var tempms = new MemoryStream();
                stream.CopyTo(tempms);
                WmlDocument doc = new WmlDocument(stream.Length.ToString(), tempms);
                if (!doc.MainDocumentPart.IsEmpty)
                {
                    sources.Add(new Source(new WmlDocument(stream.Length.ToString(), tempms), true));
                }
            }

            var mergedDoc = DocumentBuilder.BuildDocument(sources);
            mergedDoc.SaveAs(paramOutputFile);
        }
    }
}
