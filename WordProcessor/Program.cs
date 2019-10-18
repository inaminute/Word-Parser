using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;

namespace WordParser
{
    class Program
    {
        static void Main(string[] args)
        {

            Application application = new Application();
            Document document = application.Documents.Open(@"YOUR DOCUMENT PATH.docx");

            string paragraph = string.Empty;
            List<string> lines = new List<string>();
            for(int i = 0; i< document.Paragraphs.Count; i++)
            {
                string text = document.Paragraphs[i + 1].Range.Text.Trim();
                lines.Add(text);

                Console.WriteLine("Line: {0} = {1}", i, text);
            }
            application.Quit();

            string pattern = @"([a-zA-Z]+): ([\w\s\p{P}]+)|([a-zA-Z]+) : ([\w\s\p{P}]+)";

            Dictionary<string, string> hold = new Dictionary<string, string>();

            foreach (string s in lines)
            {
                Match match = Regex.Match(s, pattern);
                if (match.Success)
                {
                    string txt = match.Value;
                    Console.WriteLine(txt);

                    string[] data = txt.Split(':');
                    string key = data.GetValue(0).ToString();
                    string value = data.GetValue(1).ToString().Trim();

                    hold.Add(key, value);

                    Console.WriteLine("Key: {0} \nvalue: {1}", key, value);
                }
                Console.WriteLine();
            }

            Console.WriteLine("Count of total (key: value) pairs: {0} ",hold.Count);

            Console.ReadLine();
        }
    }
}
