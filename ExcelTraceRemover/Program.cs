using System;
using Microsoft.Win32;

namespace ExcelTraceRemover
{
    class ExcelTraceRemover
    {
        // Registry has "KEY" and "ITEM"
        static string[] extension = 
            { ".xlsx", ".xls", ".xlsm",".xlsb",".xltx",".xltm",".xlt",".xlam",".xla",".xlw",".xlr"};
        public static RegistryKey ExcelFileMRUKey()
        // Get ExcelMRU Key in Registry
        {
            var recentItemsPath = Registry.CurrentUser;
            try
            {
                recentItemsPath = Registry.CurrentUser.OpenSubKey("SOFTWARE").OpenSubKey("Microsoft").OpenSubKey("Office");

                foreach (var subKeyName in recentItemsPath.GetSubKeyNames())
                {
                    if (subKeyName.Equals("16.0"))// excel2016,2019
                    {
                        recentItemsPath = recentItemsPath.OpenSubKey(subKeyName);

                        foreach (var subKeyName2 in recentItemsPath.GetSubKeyNames())
                        {
                            if (subKeyName2.Equals("Excel"))// Excel
                            {
                                recentItemsPath = recentItemsPath.OpenSubKey(subKeyName2).OpenSubKey("User MRU");

                                foreach (var subKeyName3 in recentItemsPath.GetSubKeyNames())
                                {
                                    if (subKeyName3.StartsWith("ADAL"))
                                    {
                                        recentItemsPath = recentItemsPath.OpenSubKey(subKeyName3).OpenSubKey("File MRU", true);
                                        return recentItemsPath;
                                    }
                                }
                            }
                        }
                    }
                }

                return null;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                if (recentItemsPath != null)
                    recentItemsPath.Close();
                return null;
            }
        }
        public static List<string> ExcelLNK()
        // Get Recent Folder Path 
        {
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Recent);
            List<string> lnkFiles = new List<string>();

            foreach (var ex in extension)
            {
                string userName = Environment.UserName;

                string[] windowsRecent = Directory.GetFiles(@$"C:\Users\{userName}\AppData\Roaming\Microsoft\Windows\Recent\", "*" + ex + ".lnk");
                string[] officeRecent = Directory.GetFiles(@$"C:\Users\{userName}\AppData\Roaming\Microsoft\Office\Recent\", "*" + ex + ".lnk");
                foreach (var file in windowsRecent)
                    lnkFiles.Add(file);
                foreach (var file in officeRecent)
                    lnkFiles.Add(file);
            }

            return lnkFiles;
        }
        public static Shell32.Folder2 QuickAccess()
        // Open QuickAccess in PowerShell
        {
            Type shellAppType = Type.GetTypeFromProgID("Shell.Application");
            Object shell = Activator.CreateInstance(shellAppType);
            Shell32.Folder2 f2 = (Shell32.Folder2)shellAppType.InvokeMember("NameSpace", System.Reflection.BindingFlags.InvokeMethod, null, shell, new object[] { "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}" });

            return f2;
        }

        public static int PrintData(RegistryKey r)
        // Print Excel Trace in Registry
        {
            int total = 0;
            try
            {
                if (r != null)
                    foreach (var item in r.GetValueNames())
                    {
                        if (item.StartsWith("Item"))
                        {
                            Console.Write("Excel > 최근 항목 > ");
                            Console.WriteLine(r.GetValue(item).ToString().Split('*')[^1]);
                            //r.DeleteValue(item);
                            total++;
                        }
                    }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            return total;
        }
        public static int PrintData(List<string> lnkFiles)
        // Print Excel Trace in RecentDirectory
        {
            int total = 0;
            try
            {
                foreach (var item in lnkFiles)
                {
                    Console.WriteLine(item);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            return lnkFiles.Count;
        }
        public static int PrintData(Shell32.Folder2 f2)
        // Print Excel Trace in QuickAccess
        {
            int total = 0;
            foreach (Shell32.FolderItem fi in f2.Items())
            {
                foreach (var ex in extension)
                {
                    if (("." + fi.Name.Split('.')[^1]).Equals(ex))
                    {
                        Console.Write("즐겨찾기 > 최근에 사용한 파일 > ");
                        Console.WriteLine(fi.Name);
                        total++;
                    }
                }
            }
            return total;
        }

        public static int EraseData(RegistryKey r)
        {
            int total = 0;
            try
            {
                if (r != null)
                    foreach (var item in r.GetValueNames())
                    {
                        if (item.StartsWith("Item"))
                        {
                            Console.Write("Delete : Excel > 최근 항목 > ");
                            Console.WriteLine(r.GetValue(item).ToString().Split('*')[^1]);
                            r.DeleteValue(item);
                            total++;
                        }
                    }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            return total;
        }
        public static int EraseData(List<string> lnkFiles)
        {
            int total = 0;
            try
            {
                foreach (var item in lnkFiles)
                {
                    Console.WriteLine("Delete : " + item);
                    File.Delete(item);
                    total++;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            return total;
        }
        public static int EraseData(Shell32.Folder2 f2)
        {
            int total = 0;

            foreach (Shell32.FolderItem fi in f2.Items())
            {
                foreach (var ex in extension)
                {
                    if (("." + fi.Name.Split('.')[^1]).Equals(ex))
                    {
                        Console.Write("Delete : 즐겨찾기 > 최근에 사용한 파일 > ");
                        Console.WriteLine(fi.Name);
                        fi.InvokeVerb("remove");
                        total++;
                    }
                }
            }
            return total;
        }
    }
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                using(var recentItemsPath = ExcelTraceRemover.ExcelFileMRUKey())
                {
                    var lnkFiles = ExcelTraceRemover.ExcelLNK();
                    var quickAccess = ExcelTraceRemover.QuickAccess();

                    int count = 0;

                    count += ExcelTraceRemover.PrintData(recentItemsPath);
                    count += ExcelTraceRemover.PrintData(lnkFiles);
                    count += ExcelTraceRemover.PrintData(quickAccess);

                    Console.WriteLine($"\n{count}개의 항목이 검출되었습니다.");
                    if(count > 0)
                    {
                        Console.Write("삭제하시겠습니까? (y/n) ");

                        while (true)
                        {
                            string input = Console.ReadLine().Trim();
                            Console.WriteLine();

                            if (input.Equals("y") || input.Equals("Y"))
                            {
                                int recount = 0;
                                recount += ExcelTraceRemover.EraseData(recentItemsPath);
                                recount += ExcelTraceRemover.EraseData(lnkFiles);
                                recount += ExcelTraceRemover.EraseData(quickAccess);

                                Console.WriteLine($"{recount}개 항목 삭제 완료.");
                                Console.Write("계속 진행하시려면 아무키나 입력해주세요. ");
                                Console.ReadLine();
                                break;
                            }
                            else if (input.Equals("n") || input.Equals("N"))
                                break;

                            else
                                Console.Write("삭제하시겠습니까? (y/n) ");
                        }
                    }
                    else
                    {
                        Console.WriteLine("항목이 없습니다.");
                        Console.Write("계속 진행하시려면 아무키나 입력해주세요. ");
                        Console.ReadLine();
                    }
                }              
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
