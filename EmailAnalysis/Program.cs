using System;

namespace EmailAnalysis
{
    class Program
    {
        static void Main(string[] args)
        {
            var analyser = new EmailAnalyser();

            var emailData = analyser.GetEmailData();

            EmailAnalyser.OutputDataset(emailData);

            Console.WriteLine("DONE");
            Console.ReadLine();
        }
    }
}
