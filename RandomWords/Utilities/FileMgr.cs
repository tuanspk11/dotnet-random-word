namespace RandomWords.Utilities
{
    internal class FileMgr
    {
        public static string GetFilePath(string fileName, string dirPath = "")
        {
            string basePath = RandomWords.config == null ? string.Empty: RandomWords.config["BasePath"];
            string filePath = !string.IsNullOrEmpty(dirPath) ? dirPath :
                              !string.IsNullOrEmpty(basePath) ? basePath : Environment.CurrentDirectory;

            filePath = Path.Join(filePath, fileName);

            return filePath;
        }

        public static bool IsFileLocked(FileInfo file)
        {
            FileStream? stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
                return true;
            }
            finally
            {
                if (stream != null)
                {
                    stream.Close();
                }
            }

            return false;
        }
    }
}
