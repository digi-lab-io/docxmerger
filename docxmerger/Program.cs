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
                string outputFileName = @args[0];
                List<FileStream> list = new List<FileStream>();

                foreach (var fileName in args.Skip(1))
                {
                    list.Add(File.Open(@fileName, FileMode.Open));
                }

                mergeDocx(outputFileName, list);

            }
            catch (Exception e)
            {
                Console.WriteLine("Application Error ocurred, please check the log: ");
                Console.WriteLine(e.ToString());
            }
        }

        static void mergeDocx(string paramOutputFile, List<FileStream> paramDocumentstreams)
        {
            var sources = new List<Source>();
            

            try
            {

                foreach (var stream in paramDocumentstreams)
                {
                    var tempms = new MemoryStream();
                    if (0 != stream.Length)
                    {
                        stream.CopyTo(tempms);
                        WmlDocument doc = new WmlDocument(stream.Length.ToString(), tempms);
                        sources.Add(new Source(doc, true));
                    }
                    tempms.Close();


                }

                var mergedDoc = DocumentBuilder.BuildDocument(sources);
                mergedDoc.SaveAs(paramOutputFile);
                


            }
            catch (Exception e)
            {
                Console.WriteLine("An Error ocurred while merging, please check the log: ");
                Console.WriteLine(e.ToString());
            }
        }
    
    }
}
