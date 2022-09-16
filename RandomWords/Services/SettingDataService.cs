using System.Data;

using RandomWords.Models;
using RandomWords.Utilities;

namespace RandomWords.Services
{
    internal class SettingDataService
    {
        public List<RandomWord> GetExcelWords(string filePath, List<string> lstDataSheets, out string message)
        {
            var result = new List<RandomWord>();

            try
            {
                var excelFileMgr = new ExcelFileMgr();
                var excelFileData = excelFileMgr.ReadExcelFile(filePath, lstDataSheets, out message);
                if (excelFileData.Count() != 0)
                {
                    foreach (var fileData in excelFileData)
                    {
                        var sheetData = fileData.Value;
                        foreach (DataRow dataRow in sheetData.Rows)
                        {
                            var randomWord = new RandomWord();
                            randomWord.hiragana = dataRow["HIRAGANA"].ToString();
                            randomWord.kanji = dataRow["KANJI"].ToString();

                            result.Add(randomWord);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }

            return result;
        }
    }
}
