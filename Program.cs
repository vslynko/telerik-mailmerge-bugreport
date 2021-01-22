namespace MailMerge
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Reflection;

    using Telerik.Windows.Documents.Flow.FormatProviders.Docx;
    using Telerik.Windows.Documents.Flow.Model;
    using Telerik.Windows.Documents.Flow.Model.Editing;

    class MyDataObject
    {
        public string AUserName { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Perform Merge");
            var doc = MailMergeDataObject();

            var path = "MailMergeSample.docx";
            using (var stream = File.OpenWrite(path))
            {
                var formatProvider = new DocxFormatProvider();
                formatProvider.Export(doc, stream);
            }

            var psi = new ProcessStartInfo()
            {
                FileName = path,
                UseShellExecute = true,
            };

            Process.Start(psi);

            Console.WriteLine("Finished");
        }

        static private RadFlowDocument CreateDocTemplate()
        {
            var stream = File.OpenRead("MailMergeTemplate.docx");

            var fmtProvider = new DocxFormatProvider();

            var template = fmtProvider.Import(stream);

            return template;
        }

        static private IEnumerable GetDataSource()
        {
            var ds = new List<MyDataObject>
            {
                new MyDataObject { AUserName = "My User Name" },
            };
            return ds;
        }

        static private RadFlowDocument MailMergeDataObject()
        {
            var dataSource = GetDataSource();
            var template = CreateDocTemplate();
            return template.MailMerge(dataSource);
        }
    }
}
