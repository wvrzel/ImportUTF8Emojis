using System;
using Interop = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.IO;
using System.Windows;


public class ImportEmojis
{
    public static int Main (String[] args) {
        try {
            Interop.Application Word = new Interop.Application();

            Console.OutputEncoding = System.Text.Encoding.UTF8;

            // read emoji file
            StreamReader streamreader = new StreamReader(args[0]);
            char[] delimiter = new char[] { '\t' };
            while (streamreader.Peek() > 0)
            {
                string[] rowcolumn = streamreader.ReadLine().Split(delimiter);
                Console.WriteLine("Adding {0}\t{1}", rowcolumn[0], rowcolumn[1]);
                Word.AutoCorrect.Entries.Add(rowcolumn[0], rowcolumn[1]); // set AutoCorrectEntries
            }

            streamreader.Close();
            
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Word);

            return 0;
        }
        catch(Exception ex) {
            Console.Error.WriteLine("[ERROR] {0}", ex.Message);
            return -1;
        }
    }
}