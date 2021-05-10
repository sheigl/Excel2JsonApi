using System.Collections.Generic;
using System.IO;

namespace Excel2JsonApi.Extensions
{
    public static class FileExtensions
    {
        public static IEnumerable<FileInfo> EnsureCreatedAndEnumerateFiles(this DirectoryInfo dir)
        {
            if (!dir.Exists)
                dir.Create();

            return dir.EnumerateFiles();
        }
    }
}
