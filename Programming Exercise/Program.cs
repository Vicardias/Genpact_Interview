using Microsoft.Office.Interop.Excel;
using System;
using System.IO;

namespace Programming_Exercise
{
    class Program
    {

        private static void OnChanged(object source, FileSystemEventArgs e) {

            string extension = Path.GetExtension(e.FullPath);
            string NewFile = e.FullPath;
            string NameFile = e.Name;
            string Processed = @"C:\Users\ML20090197.WMX\Downloads\Programming Exercise\Test Path\Processed\" + NameFile;
            string Not_applicable = @"C:\Users\ML20090197.WMX\Downloads\Programming Exercise\Test Path\Not applicable\" + NameFile;
            string MasterBook = @"C:\Users\ML20090197.WMX\Downloads\Programming Exercise\Test Path\Master File.xlsx";

            Console.WriteLine(NewFile);

            if (extension == ".xlsx") {


                var bookNewFile = new Aspose.Cells.Workbook(NewFile);
                var bookMasterBook = new Aspose.Cells.Workbook(MasterBook);
                bookMasterBook.Combine(bookNewFile);
                bookMasterBook.Save(MasterBook);

                if (File.Exists(Processed)) {
                    File.Delete(Processed);
                    File.Move(NewFile, Processed);
                } else {
                    File.Move(NewFile, Processed);
                }
            } else {
                if (File.Exists(Not_applicable)) {
                    File.Delete(Processed);
                    File.Move(NewFile, Not_applicable);
                }
                else {
                    File.Move(NewFile, Not_applicable);
                }
            }
        }

        static void Main(string[] args)
        {
            FileSystemWatcher fsw = new FileSystemWatcher("C:\\Users\\ML20090197.WMX\\Downloads\\Programming Exercise\\Test Path");
            fsw.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite
                | NotifyFilters.FileName | NotifyFilters.DirectoryName;
            fsw.Created += new FileSystemEventHandler(OnChanged);
            fsw.EnableRaisingEvents = true;

            Console.WriteLine("Press \'Enter\' to quit the sample.");
            Console.ReadLine();
        }
    }
}
