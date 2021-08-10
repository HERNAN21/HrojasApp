using Microsoft.AspNetCore.Hosting;
using System;
using System.IO;

namespace HrojasApp.Providers
{
    public interface IPathProvider
    {
        string FilePath(string filename, string folder = "files");
        byte[] FileData(string filename, string folder = "files");
        string Base64FileData(string filename, string folder = "files");
        void Save(string path, byte[] data);
    }

    public class PathProvider : IPathProvider
    {
        private readonly IWebHostEnvironment _env;

        public PathProvider(IWebHostEnvironment environment)
        {
            _env = environment;
        }

        public string FilePath(string filename, string folder = "files")
        {
            return Path.Combine(_env.ContentRootPath, "wwwroot", folder, filename);
        }

        public byte[] FileData(string filename, string folder = "files")
        {
            if (string.IsNullOrEmpty(filename)) throw new ArgumentNullException(nameof(filename)); ;
            return File.ReadAllBytes(FilePath(filename, folder));
        }

        public string Base64FileData(string filename, string folder = "files")
        {
            return Convert.ToBase64String(FileData(filename, folder));
        }

        public void Save(string path, byte[] data)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));
            if (File.Exists(path))
                File.Delete(path);
            File.WriteAllBytes(path, data);
        }

    }
}
