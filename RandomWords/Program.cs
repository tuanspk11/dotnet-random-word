using Microsoft.Extensions.Configuration;
using System.Diagnostics;

using RandomWords.Helpers;
using RandomWords.Services;
using RandomWords.Utilities;
using RandomWords.Models;

namespace RandomWords
{
    public class RandomWords
    {
        public static IConfigurationRoot? config;

        static async Task Main()
        {
            ConfigureServices();

            try
            {
                var message = string.Empty;

                var filePath = FileMgr.GetFilePath(config[Defines.DATA_FILE_NAME]);
                var dataSheetsFile = FileMgr.GetFilePath(config[Defines.DATA_SHEETS]);

                var lstDataSheets = File.ReadAllLines(dataSheetsFile).Where(sheet => !string.IsNullOrEmpty(sheet)).ToList();

                var settingDataService = new SettingDataService();
                var excelWords = settingDataService.GetExcelWords(filePath, lstDataSheets, out message);

                var random = new Random();
                var randomWords1 = new List<RandomWord1>();
                var randomWordCount1 = config[Defines.RANDOM_WORD_COUNT_1];
                for (int i = 0; i < Int32.Parse(randomWordCount1); i++)
                {
                    var index = random.Next(excelWords.Count);
                    while (string.IsNullOrEmpty(excelWords[index].kanji))
                        index = random.Next(excelWords.Count);
                    var randomWord1 = new RandomWord1();
                    randomWord1.kanji = excelWords[index].kanji;
                    randomWords1.Add(randomWord1);
                    excelWords.RemoveAt(index);
                }

                var randomWords2 = new List<RandomWord2>();
                var randomWordCount2 = config[Defines.RANDOM_WORD_COUNT_2];
                for (int i = 0; i < Int32.Parse(randomWordCount2); i++)
                {
                    var index = random.Next(excelWords.Count);
                    var randomWord2 = new RandomWord2();
                    randomWord2.hiragana = excelWords[index].hiragana;
                    randomWord2.kanji = excelWords[index].kanji;
                    randomWords2.Add(randomWord2);
                    excelWords.RemoveAt(index);
                }

                var printFileName = FileMgr.GetFilePath(config[Defines.MINI_TEST_FILE_NAME]);
                MakeCSVFile(randomWords1, randomWords2, printFileName);

                var process = new Process();
                process.StartInfo = new ProcessStartInfo(printFileName)
                {
                    UseShellExecute = true
                };
                process.Start();
            }
            catch (Exception ex)
            {
                var err = ex;
                Console.Write(err.Message);
                Console.ReadKey();
            }
        }

        private static void ConfigureServices()
        {
            // Build configuration
            config = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory())
                                               .AddJsonFile("appsettings.json", false)
                                               .Build();
        }

        private static void MakeCSVFile(List<RandomWord1> randomWords1, List<RandomWord2> randomWords2, string outputFile)
        {
            System.Text.Encoding enc = System.Text.Encoding.GetEncoding("UTF-8");

            using (StreamWriter sw = new StreamWriter(outputFile, false, enc))
            {
                var cnt1 = randomWords1.Count;
                var cnt2 = randomWords2.Count;
                var len = cnt1 > cnt2 ? cnt1 : cnt2;
                for (int i = 0; i < len; i++)
                {
                    sw.WriteLine(String.Format("{0},{1},{2},{3},{4},{5},{6}",
                                               EncloseDoubleQuotes(cnt1 > i ? randomWords1[i].kanji : String.Empty),
                                               EncloseDoubleQuotes(String.Empty),
                                               EncloseDoubleQuotes(String.Empty),
                                               EncloseDoubleQuotes(String.Empty),
                                               EncloseDoubleQuotes(cnt2 > i ? randomWords2[i].hiragana : String.Empty),
                                               EncloseDoubleQuotes(String.Empty),
                                               EncloseDoubleQuotes(cnt2 > i ? randomWords2[i].kanji : String.Empty)));
                }
            }
        }

        private static string EncloseDoubleQuotes(string field)
        {
            if (field.IndexOf('"') > -1)
            {
                field = field.Replace("\"", "\"\"");
            }
            return "\"" + field + "\"";
        }
    }
}