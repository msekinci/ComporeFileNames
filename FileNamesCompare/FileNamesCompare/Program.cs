using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FileNamesCompare
{
    class Program
    {
        static void Main(string[] args)
        {
            string folderPath = @"C:\Users\msekinci\Desktop\_output";
            string excelPath = @"C:\Users\msekinci\Desktop\FileNames.xlsx";

            List<string> fileDatas = getDataFromFolder(folderPath);
            List<string> excelDatas = getDataFromExcel(excelPath);

            //Listler arasındaki farkı karşılaştırır.
            var differents = fileDatas.Except(excelDatas);
            foreach (var item in differents)
            {
                Console.WriteLine(item);
            }
            Console.ReadLine();
        }

        //Excel kolonlardaki isimleri list döner
        private static List<string> getDataFromExcel(string excelPath)
        {
            //Dosyayı okuyacağımı ve gerekli izinlerin ayarlanması.
            FileStream stream = File.Open(excelPath, FileMode.Open, FileAccess.Read);

            //Install - Package ExcelDataReader ----> PM Console
            IExcelDataReader excelReader;
            List<string> excelDatas = new List<string>();
            int counter = 0;

            //Gönderdiğim dosya xls'mi xlsx formatında mı kontrol ediliyor.
            if (Path.GetExtension(excelPath).ToUpper() == ".XLS")
            {
                //Reading from a binary Excel file ('97-2003 format; *.xls)
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            else
            {
                //Reading from a OpenXml Excel file (2007 format; *.xlsx)
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }
            //Veriler okunmaya başlıyor.
            while (excelReader.Read())
            {
                counter++;

                //ilk satır başlık olduğu için 2.satırdan okumaya başlıyorum.
                if (counter > 0)
                {
                    excelDatas.Add(excelReader.GetString(0));
                }
            }

            //Okuma bitiriliyor.
            excelReader.Close();

            return excelDatas;
        }

        //Klasör içerisindeki dosya isimlerini list döner
        private static List<string> getDataFromFolder(string folderPath)
        {
            List<string> fileDatas = new List<string>();
            DirectoryInfo directory = new DirectoryInfo(folderPath);
            FileInfo[] rgFiles = directory.GetFiles();
            foreach (FileInfo fi in rgFiles)
            {
                fileDatas.Add(clearExtension(fi.Name));
            }
            return fileDatas;
        }

        //Eğer dosya isminde istenmeyen kısımlar varsa siler (uzantılar siliniyor)
        private static string clearExtension(string fileName)
        {
            fileName = toUpperCase(fileName.Substring(0, fileName.Length - 12));

            return fileName;
        }

        //Kakarakterlerin hepsini büyük yapar.
        public static string toUpperCase(string str)
        {
            char[] chars = str.ToCharArray();

            for (int i = 0; i < chars.Length; i++)
            {
                char c = chars[i];
                if ('a' <= c && c <= 'z')
                {
                    chars[i] = (char)(c - 'a' + 'A');
                }
            }

            return new string(chars);
        }
    }
}
