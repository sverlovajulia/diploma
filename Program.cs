using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace Diplom
{
    class Program
    {
        static void Main(string[] args)
        {
            ConsoleKeyInfo cki;
            string pathRankigOfTheUniversity = Directory.GetCurrentDirectory() + "\\Rankig of the university";
            string pathMethodGaleShapley = Directory.GetCurrentDirectory() + "\\Method Gale Shapley";
            string pathMethodOneThreadOnePriority = Directory.GetCurrentDirectory() + "\\Method one thread one priority";
            string pathEGEAdmissionOneThread = Directory.GetCurrentDirectory() + "\\EGE admission one thread";
            string pathEGEAdmissionTwoThread = Directory.GetCurrentDirectory() + "\\EGE admission two thread";
            string pathEGEAdmissionThreeThread = Directory.GetCurrentDirectory() + "\\EGE admission three thread";
            int numberUniversity = 5;
            do
            {
                Console.Clear();
                Console.WriteLine("  _______________ Алгоритм распределения абитуриентов по вузам ______________");
                Console.WriteLine(" |                                                                           |");
                Console.WriteLine(" | 1: Запустить создание таблицы приоритетов абитуриентов                    |");
                Console.WriteLine(" | 2: Запустить создание таблиц ранжирования абитуриентов по вузам           |");
                Console.WriteLine(" | 3: Запустить алгоритм Гейла Шепли распределения абитуриентов по вузам     |");
                Console.WriteLine(" | 4: Запустить алгоритм в одну волну с одним приоритетом                    |");
                Console.WriteLine(" | 5: Запустить алгоритм распределения абитуриентов по вузам в одну волну    |");
                Console.WriteLine(" | 6: Запустить алгоритм распределения абитуриентов по вузам в две волны     |");
                Console.WriteLine(" | 7: Запустить алгоритм распределения абитуриентов по вузам в три волны     |");
                Console.WriteLine(" | X: Выход                                                                  |");
                Console.WriteLine(" |___________________________________________________________________________|");
                Console.WriteLine();
                Console.WriteLine();
                Console.Write(" Ввод: ");
                cki = Console.ReadKey();
                Console.ReadLine();
                switch (cki.Key)
                {
                    case (ConsoleKey.D1):
                        if (File.Exists(Directory.GetCurrentDirectory() + "\\TablePreferencesStudents.xls"))
                        {
                            File.Delete(Directory.GetCurrentDirectory() + "\\TablePreferencesStudents.xls");
                            GenerateTablePreferencesStudents(numberUniversity);
                        }
                        else
                        {
                            GenerateTablePreferencesStudents(numberUniversity);
                        }
                        break;
                    case (ConsoleKey.D2):
                        UpdateFolder(pathRankigOfTheUniversity);
                        if (File.Exists(Directory.GetCurrentDirectory() + "\\TablePreferencesStudents.xls"))
                        {
                            GenerateRankingUniversities(pathRankigOfTheUniversity, numberUniversity);
                        }
                        else
                        {
                            GenerateTablePreferencesStudents(numberUniversity);
                            GenerateRankingUniversities(pathRankigOfTheUniversity, numberUniversity);
                        }
                        break;
                    case (ConsoleKey.D3):
                        UpdateFolder(pathMethodGaleShapley);
                        if (!Directory.Exists(pathRankigOfTheUniversity))
                        {
                            Directory.CreateDirectory(pathRankigOfTheUniversity);
                        }
                        if (Directory.GetFiles(pathRankigOfTheUniversity).GetLength(0) != 0)
                        {
                            MethodGaleShapley(pathRankigOfTheUniversity, pathMethodGaleShapley);
                        }
                        else
                        {
                            if (File.Exists(Directory.GetCurrentDirectory() + "\\TablePreferencesStudents.xls"))
                            {
                                GenerateRankingUniversities(pathRankigOfTheUniversity, numberUniversity);
                            }
                            else
                            {
                                GenerateTablePreferencesStudents(numberUniversity);
                                GenerateRankingUniversities(pathRankigOfTheUniversity, numberUniversity);
                            }
                            MethodGaleShapley(pathRankigOfTheUniversity, pathMethodGaleShapley);
                        }
                        break;
                    case (ConsoleKey.D4):
                        UpdateFolder(pathMethodOneThreadOnePriority);
                        if (!Directory.Exists(pathRankigOfTheUniversity))
                        {
                            Directory.CreateDirectory(pathRankigOfTheUniversity);
                        }
                        if (Directory.GetFiles(pathRankigOfTheUniversity).GetLength(0) != 0)
                        {
                            MethodOneThreadOnePriority(pathRankigOfTheUniversity, pathMethodOneThreadOnePriority);
                        }
                        else
                        {
                            if (File.Exists(Directory.GetCurrentDirectory() + "\\TablePreferencesStudents.xls"))
                            {
                                GenerateRankingUniversities(pathRankigOfTheUniversity, numberUniversity);
                            }
                            else
                            {
                                GenerateTablePreferencesStudents(numberUniversity);
                                GenerateRankingUniversities(pathRankigOfTheUniversity, numberUniversity);
                            }
                            MethodOneThreadOnePriority(pathRankigOfTheUniversity, pathMethodOneThreadOnePriority);
                        }
                        break;
                    case (ConsoleKey.D5):
                        if (!Directory.Exists(pathRankigOfTheUniversity))
                        {
                            Directory.CreateDirectory(pathRankigOfTheUniversity);
                        }
                        if (Directory.GetFiles(pathRankigOfTheUniversity).GetLength(0) != 0)
                        {
                            if (!Directory.Exists(pathMethodGaleShapley))
                            {
                                Directory.CreateDirectory(pathMethodGaleShapley);
                            }
                            if (Directory.GetFiles(pathMethodGaleShapley).GetLength(0) != 0)
                            {
                                EGEAdmissionOneThread(pathRankigOfTheUniversity, pathEGEAdmissionOneThread, pathMethodGaleShapley);
                            }
                            else
                            {
                                MethodGaleShapley(pathRankigOfTheUniversity, pathMethodGaleShapley);
                                EGEAdmissionOneThread(pathRankigOfTheUniversity, pathEGEAdmissionOneThread, pathMethodGaleShapley);
                            }
                        }
                        else
                        {
                            if (File.Exists(Directory.GetCurrentDirectory() + "\\TablePreferencesStudents.xls"))
                            {
                                GenerateRankingUniversities(pathRankigOfTheUniversity, numberUniversity);
                            }
                            else
                            {
                                GenerateTablePreferencesStudents(numberUniversity);
                                GenerateRankingUniversities(pathRankigOfTheUniversity, numberUniversity);
                            }
                            if (!Directory.Exists(pathMethodGaleShapley))
                            {
                                Directory.CreateDirectory(pathMethodGaleShapley);
                            }
                            if (Directory.GetFiles(pathMethodGaleShapley).GetLength(0) != 0)
                            {
                                EGEAdmissionOneThread(pathRankigOfTheUniversity, pathEGEAdmissionOneThread, pathMethodGaleShapley);
                            }
                            else
                            {
                                MethodGaleShapley(pathRankigOfTheUniversity, pathMethodGaleShapley);
                                EGEAdmissionOneThread(pathRankigOfTheUniversity, pathEGEAdmissionOneThread, pathMethodGaleShapley);
                            }
                        }
                        break;
                    case (ConsoleKey.D6):
                        if (!Directory.Exists(pathRankigOfTheUniversity))
                        {
                            Directory.CreateDirectory(pathRankigOfTheUniversity);
                        }
                        if (Directory.GetFiles(pathRankigOfTheUniversity).GetLength(0) != 0)
                        {
                            if (!Directory.Exists(pathMethodGaleShapley))
                            {
                                Directory.CreateDirectory(pathMethodGaleShapley);
                            }
                            if (Directory.GetFiles(pathMethodGaleShapley).GetLength(0) != 0)
                            {
                                EGEAdmissionTwoThread(pathRankigOfTheUniversity, pathEGEAdmissionTwoThread, pathMethodGaleShapley);
                            }
                            else
                            {
                                MethodGaleShapley(pathRankigOfTheUniversity, pathMethodGaleShapley);
                                EGEAdmissionTwoThread(pathRankigOfTheUniversity, pathEGEAdmissionTwoThread, pathMethodGaleShapley);
                            }
                        }
                        else
                        {
                            if (File.Exists(Directory.GetCurrentDirectory() + "\\TablePreferencesStudents.xls"))
                            {
                                GenerateRankingUniversities(pathRankigOfTheUniversity, numberUniversity);
                            }
                            else
                            {
                                GenerateTablePreferencesStudents(numberUniversity);
                                GenerateRankingUniversities(pathRankigOfTheUniversity, numberUniversity);
                            }
                            if (!Directory.Exists(pathMethodGaleShapley))
                            {
                                Directory.CreateDirectory(pathMethodGaleShapley);
                            }
                            if (Directory.GetFiles(pathMethodGaleShapley).GetLength(0) != 0)
                            {
                                EGEAdmissionTwoThread(pathRankigOfTheUniversity, pathEGEAdmissionTwoThread, pathMethodGaleShapley);
                            }
                            else
                            {
                                MethodGaleShapley(pathRankigOfTheUniversity, pathMethodGaleShapley);
                                EGEAdmissionTwoThread(pathRankigOfTheUniversity, pathEGEAdmissionTwoThread, pathMethodGaleShapley);
                            }
                        }
                        break;
                    case (ConsoleKey.D7):
                        if (!Directory.Exists(pathRankigOfTheUniversity))
                        {
                            Directory.CreateDirectory(pathRankigOfTheUniversity);
                        }
                        if (Directory.GetFiles(pathRankigOfTheUniversity).GetLength(0) != 0)
                        {
                            if (!Directory.Exists(pathMethodGaleShapley))
                            {
                                Directory.CreateDirectory(pathMethodGaleShapley);
                            }
                            if (Directory.GetFiles(pathMethodGaleShapley).GetLength(0) != 0)
                            {
                                EGEAdmissionThreeThread(pathRankigOfTheUniversity, pathEGEAdmissionThreeThread, pathMethodGaleShapley);
                            }
                            else
                            {
                                MethodGaleShapley(pathRankigOfTheUniversity, pathMethodGaleShapley);
                                EGEAdmissionThreeThread(pathRankigOfTheUniversity, pathEGEAdmissionThreeThread, pathMethodGaleShapley);
                            }
                        }
                        else
                        {
                            if (File.Exists(Directory.GetCurrentDirectory() + "\\TablePreferencesStudents.xls"))
                            {
                                GenerateRankingUniversities(pathRankigOfTheUniversity, numberUniversity);
                            }
                            else
                            {
                                GenerateTablePreferencesStudents(numberUniversity);
                                GenerateRankingUniversities(pathRankigOfTheUniversity, numberUniversity);
                            }
                            if (!Directory.Exists(pathMethodGaleShapley))
                            {
                                Directory.CreateDirectory(pathMethodGaleShapley);
                            }
                            if (Directory.GetFiles(pathMethodGaleShapley).GetLength(0) != 0)
                            {
                                EGEAdmissionThreeThread(pathRankigOfTheUniversity, pathEGEAdmissionThreeThread, pathMethodGaleShapley);
                            }
                            else
                            {
                                MethodGaleShapley(pathRankigOfTheUniversity, pathMethodGaleShapley);
                                EGEAdmissionThreeThread(pathRankigOfTheUniversity, pathEGEAdmissionThreeThread, pathMethodGaleShapley);
                            }
                        }
                        break;
                    default:
                        break;

                }
            } while (cki.Key != ConsoleKey.X);
            CloseProcess();
        }
        public static void GenerateTablePreferencesStudents(int numberUniversity)
        {
            string pathTablePreferencesStudents = Directory.GetCurrentDirectory() + "\\TablePreferencesStudents.xls";
            bool errorFlag = false;
            ConsoleKeyInfo cki;
            int numberStudents = 100;
            while (true)
            {
                Console.Clear();
                Console.WriteLine(" ___________ Построение таблицы приоритетов абитуриентов___________");
                Console.WriteLine();
                Console.WriteLine(" * Количество абитуриентов = " + numberStudents);
                Console.WriteLine(" * Количество вузов = " + numberUniversity);
                Console.WriteLine();
                Console.Write(" Изменить начальные данные (Y/N) или вернуться в предыдущий пукт меню (X)?  ");
                cki = Console.ReadKey();
                Console.ReadLine();
                if (cki.Key == ConsoleKey.Y)
                {
                    Console.Clear();
                    Console.WriteLine(" ___________ Построение таблицы приоритетов абитуриентов___________");
                    InputOutputOptions(" Введите новое количество абитуриентов: ", ref numberStudents, " Error: Не правильный формат данных!");
                }
                if (cki.Key == ConsoleKey.N)
                {
                    break;
                }
                if (cki.Key == ConsoleKey.X)
                {
                    return;
                }
            }
            Console.WriteLine();
            Console.WriteLine(" Info: Построение запущено...");
            int[,] preferencesTable = new int[numberStudents, numberUniversity + 1];
            Random rand = new Random((int)DateTime.Now.Ticks);
            Random rand2 = new Random((int)DateTime.Now.Ticks);
            for (int i = 0; i < preferencesTable.GetLength(0); ++i)
            {
                int numberPreferredUniversities = rand.Next(1, numberUniversity + 1);
                int j = 1;
                while (j <= numberPreferredUniversities)
                {
                    int p = rand2.Next(1, numberUniversity + 1);
                    bool flag = false;
                    for (int l = 0; l < preferencesTable.GetLength(1); ++l)
                        if (p == preferencesTable[i, l]) flag = true;
                    if (flag) continue;
                    preferencesTable[i, j] = p;
                    j++;
                }
            }
            for (int i = 0; i < preferencesTable.GetLength(0); ++i)
                preferencesTable[i, 0] = i + 1;
            try
            {
                Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                xlWorkSheet.Cells[1, 1] = "Количество абитуриентов";
                xlWorkSheet.Cells[1, 2] = numberStudents;
                xlWorkSheet.Cells[2, 1] = "Количество университетов";
                xlWorkSheet.Cells[2, 2] = numberUniversity;
                xlWorkSheet.Cells[4, 1] = "№ абитуриента в общем рейтинге";
                for (int i = 2; i < numberUniversity + 2; ++i)
                    xlWorkSheet.Cells[4, i] = Convert.ToString(i - 1) + " приоритет";
                for (int i = 0; i < preferencesTable.GetLength(0); ++i)
                    for (int j = 0; j < preferencesTable.GetLength(1); ++j)
                        xlWorkSheet.Cells[5 + i, j + 1] = preferencesTable[i, j];
                if (File.Exists(pathTablePreferencesStudents))
                {
                    File.Delete(pathTablePreferencesStudents);
                }
                xlWorkBook.SaveAs(pathTablePreferencesStudents, Excel.XlFileFormat.xlWorkbookNormal, misValue,
                misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(0);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
                CloseProcess();
            }
            catch
            {
                Console.WriteLine(" Error: Не удалось сохранить файл = " + pathTablePreferencesStudents);
                errorFlag = true;
            }
            CloseProcess();
            if (errorFlag)
            {
                CloseProcess();
                Console.WriteLine(" Warning: Построение закончилось с ошибками!");
                Console.Write(" Info: Нажмите любую клавишу...");
                Console.ReadLine();
            }
            else
            {
                Console.WriteLine(" Info: Построение закончено успешно!");
                Console.Write(" Info: Нажмите любую клавишу...");
                Console.ReadLine();
            }
        }
        public static void GenerateRankingUniversities(string pathRankigOfTheUniversity, int numberUniversity)
        {
            string fileTablePreferencesStudents = Directory.GetCurrentDirectory() + "\\TablePreferencesStudents.xls";
            int[] quota = new int[numberUniversity];
            for (int i = 0; i < quota.GetLength(0); ++i)
                quota[i] = 100;
            bool errorFlag = false;
            ConsoleKeyInfo cki;
            while (true)
            {
                Console.Clear();
                Console.WriteLine(" ___________ Построение таблиц ранжирования по университетам ___________");
                Console.WriteLine();
                for (int i = 0; i < numberUniversity; ++i)
                {
                    Console.WriteLine(" * Количество мест в " + (i + 1) + " университете = " + quota[i]);
                }
                Console.WriteLine();
                Console.Write(" Изменить начальные данные (Y/N) или вернуться в предыдущий пукт меню (X)?  ");
                cki = Console.ReadKey();
                Console.ReadLine();
                if (cki.Key == ConsoleKey.Y)
                {
                    Console.Clear();
                    Console.WriteLine(" ___________ Построение таблиц ранжирования по университетам ___________");
                    for (int i = 0; i < numberUniversity; ++i)
                    {
                        InputOutputOptions(" Введите количество мест в " + (i + 1) + " университете: ", ref quota[i], " Error: Не правильный формат данных!");
                    }
                }
                if (cki.Key == ConsoleKey.N)
                {
                    break;
                }
                if (cki.Key == ConsoleKey.X)
                {
                    return;
                }
            }
            Console.WriteLine();
            Console.WriteLine(" Info: Построение запущено...");
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            try
            {
                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(fileTablePreferencesStudents);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                int numberStudents = Convert.ToInt32(xlWorkSheet.Cells[1, 2].Text.ToString());
                int[,] preferencesTable = new int[numberStudents, numberUniversity + 1];
                for (int i = 0; i < preferencesTable.GetLength(0); ++i)
                    for (int j = 0; j < preferencesTable.GetLength(1); ++j)
                        preferencesTable[i, j] = Convert.ToInt32(xlWorkSheet.Cells[5 + i, j + 1].Text.ToString());
                xlWorkBook.Close(0);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
                CloseProcess();

                for (int l = 1; l <= numberUniversity; ++l)
                {
                    try
                    {
                        xlApp = new Microsoft.Office.Interop.Excel.Application();
                        xlWorkBook = xlApp.Workbooks.Add(misValue);
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                        xlWorkSheet.Cells[2, 1] = "Рейтинг абитуриентов в университете:";
                        xlWorkSheet.Cells[2, 2] = l;
                        xlWorkSheet.Cells[3, 1] = "Количество мест:";
                        xlWorkSheet.Cells[3, 2] = quota[l - 1];
                        xlWorkSheet.Cells[5, 1] = "№ абитуриента в рейтинге вуза";
                        xlWorkSheet.Cells[5, 2] = "№ абитуриента в общем рейтинге";
                        xlWorkSheet.Cells[5, 3] = "Приоритет";
                        int k = 0;
                        for (int i = 0; i < preferencesTable.GetLength(0); ++i)
                            for (int j = 1; j < preferencesTable.GetLength(1); ++j)
                                if (preferencesTable[i, j] == l)
                                {
                                    xlWorkSheet.Cells[6 + k, 1] = k + 1;
                                    xlWorkSheet.Cells[6 + k, 2] = i + 1;
                                    xlWorkSheet.Cells[6 + k, 3] = j;
                                    k++;
                                }
                        xlWorkBook.SaveAs(pathRankigOfTheUniversity + "\\Table ranking university" + l.ToString() + ".xls", Excel.XlFileFormat.xlWorkbookNormal,
                        misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                        xlWorkBook.Close(0);
                        xlApp.Quit();
                        Marshal.ReleaseComObject(xlWorkSheet);
                        Marshal.ReleaseComObject(xlWorkBook);
                        Marshal.ReleaseComObject(xlApp);
                        CloseProcess();
                    }
                    catch
                    {
                        CloseProcess();
                        Console.WriteLine(" Error: Не удалось сохранить файл = " + pathRankigOfTheUniversity + "\\Table ranking university" + l.ToString() + ".xls");
                        errorFlag = true;
                    }
                }
            }
            catch
            {
                CloseProcess();
                Console.WriteLine(" Error: Не удалось открыть файл = " + fileTablePreferencesStudents);
                errorFlag = true;
            }
            if (errorFlag)
            {
                CloseProcess();
                Console.WriteLine(" Warning: Построение закончилось с ошибками!");
                Console.Write(" Info: Нажмите любую клавишу...");
                Console.ReadLine();
            }
            else
            {
                Console.WriteLine(" Info: Построение закончено успешно!");
                Console.Write(" Info: Нажмите любую клавишу...");
                Console.ReadLine();
            }
            CloseProcess();
        }
        public static void MethodGaleShapley(string pathRankigOfTheUniversity, string pathMethodGaleShapley)
        {
            string fileTablePreferencesStudents = Directory.GetCurrentDirectory() + "\\TablePreferencesStudents.xls";
            bool errorFlag = false;
            Console.Clear();
            Console.WriteLine(" ___________ Алгоритм Гейла Шепли ___________");
            Console.WriteLine();
            Console.WriteLine(" Info: Алгоритм запушен...");
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            try
            {
                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(fileTablePreferencesStudents);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                int numberStudents = Convert.ToInt32(xlWorkSheet.Cells[1, 2].Text.ToString());
                int numberUniversity = Convert.ToInt32(xlWorkSheet.Cells[2, 2].Text.ToString());
                int[,] preferencesTable = new int[numberStudents, numberUniversity + 1];
                for (int i = 0; i < preferencesTable.GetLength(0); ++i)
                    for (int j = 0; j < preferencesTable.GetLength(1); ++j)
                        preferencesTable[i, j] = Convert.ToInt32(xlWorkSheet.Cells[5 + i, j + 1].Text.ToString());
                xlWorkBook.Close(0);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
                CloseProcess();

                int[] numberSeats = new int[numberUniversity + 1];
                for (int i = 0; i < numberUniversity; ++i)
                {
                    try
                    {
                        xlApp = new Microsoft.Office.Interop.Excel.Application();
                        xlWorkBook = xlApp.Workbooks.Open(pathRankigOfTheUniversity + "\\Table ranking university" + (i + 1).ToString() + ".xls");
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                        numberSeats[i] = Convert.ToInt32(xlWorkSheet.Cells[3, 2].Text.ToString());
                        xlWorkBook.Close(0);
                        xlApp.Quit();
                        Marshal.ReleaseComObject(xlWorkSheet);
                        Marshal.ReleaseComObject(xlWorkBook);
                        Marshal.ReleaseComObject(xlApp);
                        CloseProcess();
                    }
                    catch
                    {
                        Console.WriteLine(" Error: Не удалось открыть файл = " + pathRankigOfTheUniversity + "\\Table ranking university" + (i + 1).ToString() + ".xls");
                        errorFlag = true;
                    }
                }
                int[] numberIncoming = new int[numberUniversity + 1];
                List<int> studentsNoIncoming = new List<int>();
                List<List<int>> tableStudentsIncoming = new List<List<int>>();
                for (int j = 1; j < preferencesTable.GetLength(1); ++j)
                {
                    List<int> table = new List<int>();
                    tableStudentsIncoming.Add(table);
                }
                for (int i = 0; i < numberStudents; ++i)
                {
                    bool flag = false;
                    for (int j = 1; j < preferencesTable.GetLength(1); ++j)
                    {
                        if (preferencesTable[i, j] > 0)
                        {
                            if (numberSeats[preferencesTable[i, j] - 1] > 0)
                            {
                                numberSeats[preferencesTable[i, j] - 1]--;
                                numberIncoming[preferencesTable[i, j] - 1]++;
                                tableStudentsIncoming[preferencesTable[i, j] - 1].Add(i + 1);
                                flag = true;
                                break;
                            }
                        }
                        else
                        {
                            break;
                        }
                    }
                    if (!flag)
                    {
                        studentsNoIncoming.Add(i + 1);
                    }
                }
                for (int i = 0; i < tableStudentsIncoming.Count; ++i)
                {
                    try
                    {
                        xlApp = new Microsoft.Office.Interop.Excel.Application();
                        xlWorkBook = xlApp.Workbooks.Add(misValue);
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                        xlWorkSheet.Cells[2, 1] = "Список абитуриентов поступивших в вуз:";
                        xlWorkSheet.Cells[2, 2] = i + 1;
                        xlWorkSheet.Cells[3, 1] = "Количество поступивших абитуриантов:";
                        xlWorkSheet.Cells[3, 2] = numberIncoming[i];
                        xlWorkSheet.Cells[5, 1] = "№ абитуриента";
                        for (int j = 0; j < tableStudentsIncoming[i].Count; ++j)
                        {
                            xlWorkSheet.Cells[6 + j, 1] = tableStudentsIncoming[i][j];
                        }

                        xlWorkBook.SaveAs(pathMethodGaleShapley + "\\List of students incoming to the university = " + (i + 1).ToString() + ".xls", Excel.XlFileFormat.xlWorkbookNormal,
                        misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                        xlWorkBook.Close(0);
                        xlApp.Quit();
                        Marshal.ReleaseComObject(xlWorkSheet);
                        Marshal.ReleaseComObject(xlWorkBook);
                        Marshal.ReleaseComObject(xlApp);
                        CloseProcess();
                    }
                    catch
                    {
                        CloseProcess();
                        Console.WriteLine(" Error: Не удалось сохранить файл = " + pathMethodGaleShapley + "\\List of students incoming to the university = " + (i + 1).ToString() + ".xls");
                        errorFlag = true;
                    }
                }
                try
                {
                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    xlWorkSheet.Cells[2, 1] = "Список абитуриентов не поступивших в вуз";
                    xlWorkSheet.Cells[3, 1] = "Количество не поступивших абитуриантов:";
                    if (studentsNoIncoming.Count == 0)
                    {
                        xlWorkSheet.Cells[3, 2] = 0;
                    }
                    else
                    {
                        xlWorkSheet.Cells[3, 2] = studentsNoIncoming.Count;
                    }
                    xlWorkSheet.Cells[5, 1] = "№ абитуриента";
                    for (int i = 0; i < studentsNoIncoming.Count; ++i)
                    {
                        xlWorkSheet.Cells[6 + i, 1] = studentsNoIncoming[i];
                    }
                    xlWorkBook.SaveAs(pathMethodGaleShapley + "\\List of students not incoming to the university.xls", Excel.XlFileFormat.xlWorkbookNormal,
                    misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                }
                catch
                {
                    CloseProcess();
                    Console.WriteLine(" Error: Не удалось сохранить файл = " + pathMethodGaleShapley + "\\List of students not incoming to the university.xls");
                    errorFlag = true;
                }
                xlWorkBook.Close(0);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
                CloseProcess();
            }
            catch
            {
                CloseProcess();
                Console.WriteLine(" Error: Не удалось открыть файл = " + fileTablePreferencesStudents);
                errorFlag = true;
            }
            if (errorFlag)
            {
                CloseProcess();
                Console.WriteLine(" Warning: Алгоритм закончился с ошибками!");
                Console.Write(" Info: Нажмите любую клавишу...");
                Console.ReadLine();
            }
            else
            {
                Console.WriteLine(" Info: Алгоритм закончился успешно!");
                Console.Write(" Info: Нажмите любую клавишу...");
                Console.ReadLine();
            }
            CloseProcess();
        }
        public static void MethodOneThreadOnePriority(string pathRankigOfTheUniversity, string MethodOneThreadOnePriority)
        {
            string fileTablePreferencesStudents = Directory.GetCurrentDirectory() + "\\TablePreferencesStudents.xls";
            bool errorFlag = false;
            Console.Clear();
            Console.WriteLine(" ___________ Алгоритм в одну волну с одним приоритетом ___________");
            Console.WriteLine();
            Console.WriteLine(" Info: Алгоритм запушен...");
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            try
            {
                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(fileTablePreferencesStudents);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                int numberStudents = Convert.ToInt32(xlWorkSheet.Cells[1, 2].Text.ToString());
                int numberUniversity = Convert.ToInt32(xlWorkSheet.Cells[2, 2].Text.ToString());
                int[,] preferencesTable = new int[numberStudents, numberUniversity + 1];
                for (int i = 0; i < preferencesTable.GetLength(0); ++i)
                    for (int j = 0; j < preferencesTable.GetLength(1); ++j)
                        preferencesTable[i, j] = Convert.ToInt32(xlWorkSheet.Cells[5 + i, j + 1].Text.ToString());
                xlWorkBook.Close(0);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
                CloseProcess();

                int[] numberSeats = new int[numberUniversity + 1];
                for (int i = 0; i < numberUniversity; ++i)
                {
                    try
                    {
                        xlApp = new Microsoft.Office.Interop.Excel.Application();
                        xlWorkBook = xlApp.Workbooks.Open(pathRankigOfTheUniversity + "\\Table ranking university" + (i + 1).ToString() + ".xls");
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                        numberSeats[i] = Convert.ToInt32(xlWorkSheet.Cells[3, 2].Text.ToString());
                        xlWorkBook.Close(0);
                        xlApp.Quit();
                        Marshal.ReleaseComObject(xlWorkSheet);
                        Marshal.ReleaseComObject(xlWorkBook);
                        Marshal.ReleaseComObject(xlApp);
                        CloseProcess();
                    }
                    catch
                    {
                        Console.WriteLine(" Error: Не удалось открыть файл = " + pathRankigOfTheUniversity + "\\Table ranking university" + (i + 1).ToString() + ".xls");
                        errorFlag = true;
                    }
                }
                int[] numberIncoming = new int[numberUniversity + 1];
                List<int> studentsNoIncoming = new List<int>();
                List<List<int>> tableStudentsIncoming = new List<List<int>>();
                for (int j = 1; j < preferencesTable.GetLength(1); ++j)
                {
                    List<int> table = new List<int>();
                    tableStudentsIncoming.Add(table);
                }
                for (int i = 0; i < numberStudents; ++i)
                {
                    if (preferencesTable[i, 1] > 0)
                    {
                        if (numberSeats[preferencesTable[i, 1] - 1] > 0)
                        {
                            numberSeats[preferencesTable[i, 1] - 1]--;
                            numberIncoming[preferencesTable[i, 1] - 1]++;
                            tableStudentsIncoming[preferencesTable[i, 1] - 1].Add(i + 1);

                        }
                        else
                        {
                            studentsNoIncoming.Add(i + 1);
                        }
                    }
                }
                for (int i = 0; i < tableStudentsIncoming.Count; ++i)
                {
                    try
                    {
                        xlApp = new Microsoft.Office.Interop.Excel.Application();
                        xlWorkBook = xlApp.Workbooks.Add(misValue);
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                        xlWorkSheet.Cells[2, 1] = "Список абитуриентов поступивших в вуз:";
                        xlWorkSheet.Cells[2, 2] = i + 1;
                        xlWorkSheet.Cells[3, 1] = "Количество поступивших абитуриантов:";
                        xlWorkSheet.Cells[3, 2] = numberIncoming[i];
                        xlWorkSheet.Cells[5, 1] = "№ абитуриента";
                        for (int j = 0; j < tableStudentsIncoming[i].Count; ++j)
                        {
                            xlWorkSheet.Cells[6 + j, 1] = tableStudentsIncoming[i][j];
                        }

                        xlWorkBook.SaveAs(MethodOneThreadOnePriority + "\\List of students incoming to the university = " + (i + 1).ToString() + ".xls", Excel.XlFileFormat.xlWorkbookNormal,
                        misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                        xlWorkBook.Close(0);
                        xlApp.Quit();
                        Marshal.ReleaseComObject(xlWorkSheet);
                        Marshal.ReleaseComObject(xlWorkBook);
                        Marshal.ReleaseComObject(xlApp);
                        CloseProcess();
                    }
                    catch
                    {
                        CloseProcess();
                        Console.WriteLine(" Error: Не удалось сохранить файл = " + MethodOneThreadOnePriority + "\\List of students incoming to the university = " + (i + 1).ToString() + ".xls");
                        errorFlag = true;
                    }
                }
                try
                {
                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    xlWorkSheet.Cells[2, 1] = "Список абитуриентов не поступивших в вуз";
                    xlWorkSheet.Cells[3, 1] = "Количество не поступивших абитуриантов:";
                    if (studentsNoIncoming.Count == 0)
                    {
                        xlWorkSheet.Cells[3, 2] = 0;
                    }
                    else
                    {
                        xlWorkSheet.Cells[3, 2] = studentsNoIncoming.Count;
                    }
                    xlWorkSheet.Cells[5, 1] = "№ абитуриента";
                    for (int i = 0; i < studentsNoIncoming.Count; ++i)
                    {
                        xlWorkSheet.Cells[6 + i, 1] = studentsNoIncoming[i];
                    }
                    xlWorkBook.SaveAs(MethodOneThreadOnePriority + "\\List of students not incoming to the university.xls", Excel.XlFileFormat.xlWorkbookNormal,
                    misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                }
                catch
                {
                    CloseProcess();
                    Console.WriteLine(" Error: Не удалось сохранить файл = " + MethodOneThreadOnePriority + "\\List of students not incoming to the university.xls");
                    errorFlag = true;
                }
                xlWorkBook.Close(0);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
                CloseProcess();
            }
            catch
            {
                CloseProcess();
                Console.WriteLine(" Error: Не удалось открыть файл = " + fileTablePreferencesStudents);
                errorFlag = true;
            }
            if (errorFlag)
            {
                CloseProcess();
                Console.WriteLine(" Warning: Алгоритм закончился с ошибками!");
                Console.Write(" Info: Нажмите любую клавишу...");
                Console.ReadLine();
            }
            else
            {
                Console.WriteLine(" Info: Алгоритм закончился успешно!");
                Console.Write(" Info: Нажмите любую клавишу...");
                Console.ReadLine();
            }
            CloseProcess();
        }
        public static void EGEAdmissionOneThread(string pathRankigOfTheUniversity, string pathEGEAdmissionOneThread, string pathMethodGaleShapley)
        {
            Console.Clear();
            string fileTablePreferencesStudents = Directory.GetCurrentDirectory() + "\\TablePreferencesStudents.xls";
            string pathEGEAdmissionOneThreadNew;
            ConsoleKeyInfo cki;
            double coef = 1.0;
            bool debugFlag = false;
            while (true)
            {
                Console.Clear();
                Console.WriteLine(" ___________ Алгоритм распределения абитуриентов по вузам в одну волну ___________");
                Console.WriteLine();
                Console.WriteLine(" * Коэффицент оставшейся квоты в первую волну = " + coef);
                Console.WriteLine(" * Коэффицент для поступления в первую волну = " + coef);
                if (!debugFlag) Console.WriteLine(" * Подробный вывод = No");
                else Console.WriteLine(" * Подробный вывод = Yes");
                Console.WriteLine();
                Console.Write(" Изменить начальные данные (Y/N) или вернуться в предыдущий пукт меню (X)?  ");
                cki = Console.ReadKey();
                Console.ReadLine();
                if (cki.Key == ConsoleKey.Y)
                {
                    Console.Clear();
                    Console.WriteLine(" ___________ Алгоритм распределения абитуриентов по вузам в одну волну ___________");
                    InputOutputOptions(" Подробный вывод (Y/N): ", ref debugFlag, " Error: Не правильный формат данных!");
                }
                if (cki.Key == ConsoleKey.N)
                {
                    break;
                }
                if (cki.Key == ConsoleKey.X)
                {
                    return;
                }
            }
            Console.WriteLine();
            Console.WriteLine(" Info: Построение запущено c lambda = (m - 1)/m ...");
            pathEGEAdmissionOneThreadNew = pathEGEAdmissionOneThread + " v.1";
            RunOneWave(pathRankigOfTheUniversity, pathEGEAdmissionOneThreadNew, pathMethodGaleShapley, fileTablePreferencesStudents,
            coef, coef, 1, debugFlag);
            Console.WriteLine(" Info: Построение запущено c lambda = (S - 1) ...");
            pathEGEAdmissionOneThreadNew = pathEGEAdmissionOneThread + " v.2";
            RunOneWave(pathRankigOfTheUniversity, pathEGEAdmissionOneThreadNew, pathMethodGaleShapley, fileTablePreferencesStudents,
            coef, coef, 2, debugFlag);
            Console.WriteLine(" Info: Построение запущено c lambda = (S - 1)/2 ...");
            pathEGEAdmissionOneThreadNew = pathEGEAdmissionOneThread + " v.3";
            RunOneWave(pathRankigOfTheUniversity, pathEGEAdmissionOneThreadNew, pathMethodGaleShapley, fileTablePreferencesStudents,
            coef, coef, 3, debugFlag);
            Console.WriteLine(" Info: Построение запущено c lambda = S/2 ...");
            pathEGEAdmissionOneThreadNew = pathEGEAdmissionOneThread + " v.4";
            RunOneWave(pathRankigOfTheUniversity, pathEGEAdmissionOneThreadNew, pathMethodGaleShapley, fileTablePreferencesStudents,
            coef, coef, 4, debugFlag);
            Console.Write(" Info: Нажмите любую клавишу...");
            Console.ReadLine();
        }
        public static void EGEAdmissionTwoThread(string pathRankigOfTheUniversity, string pathEGEAdmissionTwoThread, string pathMethodGaleShapley)
        {
            Console.Clear();
            string fileTablePreferencesStudents = Directory.GetCurrentDirectory() + "\\TablePreferencesStudents.xls";
            string pathEGEAdmissionTwoThreadNew;
            ConsoleKeyInfo cki;
            double coef11 = 1; double coef12 = 0.8;
            double coef21 = 0.2; double coef22 = 0.2;
            bool debugFlag = false;
            while (true)
            {
                Console.Clear();
                Console.WriteLine(" ___________ Алгоритм распределения абитуриентов по вузам в две волны ___________");
                Console.WriteLine();
                Console.WriteLine(" * Коэффицент оставшейся квоты в первую волну = " + coef11);
                Console.WriteLine(" * Коэффицент для поступления в первую волну = " + coef12);
                Console.WriteLine(" * Коэффицент оставшейся квоты во вторую волну = " + coef21);
                Console.WriteLine(" * Коэффицент для поступления во вторую волну = " + coef22);
                if (!debugFlag) Console.WriteLine(" * Подробный вывод = No");
                else Console.WriteLine(" * Подробный вывод = Yes");
                Console.WriteLine();
                Console.Write(" Изменить начальные данные (Y/N) или вернуться в предыдущий пукт меню (X)?  ");
                cki = Console.ReadKey();
                Console.ReadLine();
                if (cki.Key == ConsoleKey.Y)
                {
                    Console.Clear();
                    Console.WriteLine(" ___________ Алгоритм распределения абитуриентов по вузам в две волны ___________");
                    InputOutputOptions(" Подробный вывод (Y/N): ", ref debugFlag, " Error: Не правильный формат данных!");
                }
                if (cki.Key == ConsoleKey.N)
                {
                    break;
                }
                if (cki.Key == ConsoleKey.X)
                {
                    return;
                }
            }
            pathEGEAdmissionTwoThreadNew = pathEGEAdmissionTwoThread + " v.1";
            string pathFirstWave = pathEGEAdmissionTwoThreadNew + "\\First wave";
            string pathSecondWave = pathEGEAdmissionTwoThreadNew + "\\Second wave";
            string pathTotals = pathEGEAdmissionTwoThreadNew + "\\Totals";
            string pathUnstablePairs = pathEGEAdmissionTwoThreadNew + "\\Unstable pairs";
            Console.WriteLine();
            Console.WriteLine(" Info: Построение запущено c lambda = (m - 1)/m ...");
            RunTwoWaves(pathRankigOfTheUniversity, pathEGEAdmissionTwoThreadNew, pathMethodGaleShapley, fileTablePreferencesStudents,
            pathFirstWave, pathSecondWave, pathTotals, pathUnstablePairs, coef11, coef12, coef21, coef22, 1, debugFlag);

            pathEGEAdmissionTwoThreadNew = pathEGEAdmissionTwoThread + " v.2";
            pathFirstWave = pathEGEAdmissionTwoThreadNew + "\\First wave";
            pathSecondWave = pathEGEAdmissionTwoThreadNew + "\\Second wave";
            pathTotals = pathEGEAdmissionTwoThreadNew + "\\Totals";
            pathUnstablePairs = pathEGEAdmissionTwoThreadNew + "\\Unstable pairs";
            Console.WriteLine(" Info: Построение запущено c lambda = (S - 1) ...");
            RunTwoWaves(pathRankigOfTheUniversity, pathEGEAdmissionTwoThreadNew, pathMethodGaleShapley, fileTablePreferencesStudents,
            pathFirstWave, pathSecondWave, pathTotals, pathUnstablePairs, coef11, coef12, coef21, coef22, 2, debugFlag);

            pathEGEAdmissionTwoThreadNew = pathEGEAdmissionTwoThread + " v.3";
            pathFirstWave = pathEGEAdmissionTwoThreadNew + "\\First wave";
            pathSecondWave = pathEGEAdmissionTwoThreadNew + "\\Second wave";
            pathTotals = pathEGEAdmissionTwoThreadNew + "\\Totals";
            pathUnstablePairs = pathEGEAdmissionTwoThreadNew + "\\Unstable pairs";
            Console.WriteLine(" Info: Построение запущено c lambda = (S - 1)/2 ...");
            RunTwoWaves(pathRankigOfTheUniversity, pathEGEAdmissionTwoThreadNew, pathMethodGaleShapley, fileTablePreferencesStudents,
            pathFirstWave, pathSecondWave, pathTotals, pathUnstablePairs, coef11, coef12, coef21, coef22, 3, debugFlag);

            pathEGEAdmissionTwoThreadNew = pathEGEAdmissionTwoThread + " v.4";
            pathFirstWave = pathEGEAdmissionTwoThreadNew + "\\First wave";
            pathSecondWave = pathEGEAdmissionTwoThreadNew + "\\Second wave";
            pathTotals = pathEGEAdmissionTwoThreadNew + "\\Totals";
            pathUnstablePairs = pathEGEAdmissionTwoThreadNew + "\\Unstable pairs";
            Console.WriteLine(" Info: Построение запущено c lambda = S/2 ...");
            RunTwoWaves(pathRankigOfTheUniversity, pathEGEAdmissionTwoThreadNew, pathMethodGaleShapley, fileTablePreferencesStudents,
            pathFirstWave, pathSecondWave, pathTotals, pathUnstablePairs, coef11, coef12, coef21, coef22, 4, debugFlag);
            Console.Write(" Info: Нажмите любую клавишу...");
            Console.ReadLine();
        }
        public static void EGEAdmissionThreeThread(string pathRankigOfTheUniversity, string pathEGEAdmissionThreeThread, string pathMethodGaleShapley)
        {
            Console.Clear();
            string fileTablePreferencesStudents = Directory.GetCurrentDirectory() + "\\TablePreferencesStudents.xls";
            string pathEGEAdmissionThreeThreadNew;
            ConsoleKeyInfo cki;
            double coef11 = 1; double coef12 = 0.5;
            double coef21 = 0.5; double coef22 = 0.3;
            double coef31 = 0.2; double coef32 = 0.2;
            bool debugFlag = false;
            while (true)
            {
                Console.Clear();
                Console.WriteLine(" ___________ Алгоритм распределения абитуриентов по вузам в три волны ___________");
                Console.WriteLine();
                Console.WriteLine(" * Коэффицент оставшейся квоты в первую волну = " + coef11);
                Console.WriteLine(" * Коэффицент для поступления в первую волну = " + coef12);
                Console.WriteLine(" * Коэффицент оставшейся квоты во вторую волну = " + coef21);
                Console.WriteLine(" * Коэффицент для поступления во вторую волну = " + coef22);
                Console.WriteLine(" * Коэффицент оставшейся квоты в третью волну = " + coef31);
                Console.WriteLine(" * Коэффицент для поступления в третью волну = " + coef32);
                if (!debugFlag) Console.WriteLine(" * Подробный вывод = No");
                else Console.WriteLine(" * Подробный вывод = Yes");
                Console.WriteLine();
                Console.Write(" Изменить начальные данные (Y/N) или вернуться в предыдущий пукт меню (X)?  ");
                cki = Console.ReadKey();
                Console.ReadLine();
                if (cki.Key == ConsoleKey.Y)
                {
                    Console.Clear();
                    Console.WriteLine(" ___________ Алгоритм распределения абитуриентов по вузам в три волны ___________");
                    InputOutputOptions(" Подробный вывод (Y/N): ", ref debugFlag, " Error: Не правильный формат данных!");
                }
                if (cki.Key == ConsoleKey.N)
                {
                    break;
                }
                if (cki.Key == ConsoleKey.X)
                {
                    return;
                }
            }
            pathEGEAdmissionThreeThreadNew = pathEGEAdmissionThreeThread + " v.1";
            string pathFirstWave = pathEGEAdmissionThreeThreadNew + "\\First wave";
            string pathSecondWave = pathEGEAdmissionThreeThreadNew + "\\Second wave";
            string pathThirdWave = pathEGEAdmissionThreeThreadNew + "\\Third wave";
            string pathTotals = pathEGEAdmissionThreeThreadNew + "\\Totals";
            string pathUnstablePairs = pathEGEAdmissionThreeThreadNew + "\\Unstable pairs";
            Console.WriteLine();
            Console.WriteLine(" Info: Построение запущено c lambda = (m - 1)/m ...");
            RunThreeWaves(pathRankigOfTheUniversity, pathEGEAdmissionThreeThreadNew, pathMethodGaleShapley, fileTablePreferencesStudents,
                pathFirstWave, pathSecondWave, pathThirdWave, pathTotals, pathUnstablePairs, coef11, coef12, coef21, coef22,
                coef31, coef32, 1, debugFlag);

            pathEGEAdmissionThreeThreadNew = pathEGEAdmissionThreeThread + " v.2";
            pathFirstWave = pathEGEAdmissionThreeThreadNew + "\\First wave";
            pathSecondWave = pathEGEAdmissionThreeThreadNew + "\\Second wave";
            pathThirdWave = pathEGEAdmissionThreeThreadNew + "\\Third wave";
            pathTotals = pathEGEAdmissionThreeThreadNew + "\\Totals";
            pathUnstablePairs = pathEGEAdmissionThreeThreadNew + "\\Unstable pairs";
            Console.WriteLine(" Info: Построение запущено c lambda = (S - 1) ...");
            RunThreeWaves(pathRankigOfTheUniversity, pathEGEAdmissionThreeThreadNew, pathMethodGaleShapley, fileTablePreferencesStudents,
                pathFirstWave, pathSecondWave, pathThirdWave, pathTotals, pathUnstablePairs, coef11, coef12, coef21, coef22,
                coef31, coef32, 2, debugFlag);

            pathEGEAdmissionThreeThreadNew = pathEGEAdmissionThreeThread + " v.3";
            pathFirstWave = pathEGEAdmissionThreeThreadNew + "\\First wave";
            pathSecondWave = pathEGEAdmissionThreeThreadNew + "\\Second wave";
            pathThirdWave = pathEGEAdmissionThreeThreadNew + "\\Third wave";
            pathTotals = pathEGEAdmissionThreeThreadNew + "\\Totals";
            pathUnstablePairs = pathEGEAdmissionThreeThreadNew + "\\Unstable pairs";
            Console.WriteLine(" Info: Построение запущено c lambda = (S - 1)/2 ...");
            RunThreeWaves(pathRankigOfTheUniversity, pathEGEAdmissionThreeThreadNew, pathMethodGaleShapley, fileTablePreferencesStudents,
                pathFirstWave, pathSecondWave, pathThirdWave, pathTotals, pathUnstablePairs, coef11, coef12, coef21, coef22,
                coef31, coef32, 3, debugFlag);

            pathEGEAdmissionThreeThreadNew = pathEGEAdmissionThreeThread + " v.4";
            pathFirstWave = pathEGEAdmissionThreeThreadNew + "\\First wave";
            pathSecondWave = pathEGEAdmissionThreeThreadNew + "\\Second wave";
            pathThirdWave = pathEGEAdmissionThreeThreadNew + "\\Third wave";
            pathTotals = pathEGEAdmissionThreeThreadNew + "\\Totals";
            pathUnstablePairs = pathEGEAdmissionThreeThreadNew + "\\Unstable pairs";
            Console.WriteLine(" Info: Построение запущено c lambda = S/2 ...");
            RunThreeWaves(pathRankigOfTheUniversity, pathEGEAdmissionThreeThreadNew, pathMethodGaleShapley, fileTablePreferencesStudents,
                pathFirstWave, pathSecondWave, pathThirdWave, pathTotals, pathUnstablePairs, coef11, coef12, coef21, coef22,
                coef31, coef32, 4, debugFlag);
            Console.Write(" Info: Нажмите любую клавишу...");
            Console.ReadLine();
        }
        private static void InputOutputOptions(string message, ref int value, string errorMessage)
        {
            while (true)
            {
                Console.Write(message);
                string[] tokens = Console.ReadLine().Split();
                if (!(int.TryParse(tokens[0], out value)))
                {
                    Console.WriteLine(errorMessage);
                    continue;
                }
                else
                {
                    value = int.Parse(tokens[0]);
                    break;
                }
            }
        }
        private static void InputOutputOptions(string message, ref bool value, string errorMessage)
        {
            while (true)
            {
                Console.Write(message);
                ConsoleKeyInfo cki = Console.ReadKey();
                Console.ReadLine();
                if (cki.Key == ConsoleKey.Y)
                {
                    value = true;
                    break;
                }
                if (cki.Key == ConsoleKey.N)
                {
                    value = false;
                    break;
                }
                if ((cki.Key != ConsoleKey.Y) && (cki.Key != ConsoleKey.N))
                {
                    Console.WriteLine(errorMessage);
                    continue;
                }
            }
        }
        private static void UpdateFolder(string pathFolder)
        {
            if (!Directory.Exists(pathFolder))
            {
                Directory.CreateDirectory(pathFolder);
            }
            else
            {
                Directory.Delete(pathFolder, true);
                Directory.CreateDirectory(pathFolder);
            }
        }
        private static int[,] DeleteRows(int[,] table, List<int> nums, int numberRowsTable)
        {
            int[,] temp = new int[table.GetLength(0), table.GetLength(1)];
            int i, j, k;
            int index = 0;
            for (i = 0; i < numberRowsTable; i++)
            {
                if (nums.IndexOf(i + 1) == -1)
                {
                    continue;

                }
                else
                {
                    for (j = 0; j < table.GetLength(0); ++j)
                        if (table[j, 0] == (i + 1))
                        {
                            for (k = 0; k < table.GetLength(1); k++)
                            {
                                temp[index, k] = table[j, k];
                            }
                        }
                    index++;
                }

            }
            int size = 1;
            for (i = 0; i < temp.GetLength(0); ++i)
                if (temp[i, 0] == 0)
                {
                    size = i;
                    break;
                }
            int[,] temp2 = new int[size, temp.GetLength(1)];
            for (i = 0; i < size; ++i)
                for (j = 0; j < temp.GetLength(1); ++j)
                    temp2[i, j] = temp[i, j];
            return temp2;
        }
        private static void Totals(string path, string pathFirstWave, string pathSecondWave, string pathThirdWave, int numberUniversity, ref bool errorFlag)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            List<List<int>> tableStudentsIncoming = new List<List<int>>();
            int k;
            Console.WriteLine(" Info: Запущено построение итогов...");
            for (int i = 0; i < numberUniversity; ++i)
            {
                List<int> tables = new List<int>();
                try
                {
                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Open(pathFirstWave + "\\List of students incoming to the university = " + (i + 1).ToString() + ".xls");
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    k = 1;
                    while (!(xlWorkSheet.Cells[5 + k, 1].Text.ToString()).Equals(""))
                    {
                        tables.Add(Convert.ToInt32(xlWorkSheet.Cells[5 + k, 1].Text.ToString()));
                        k++;
                    }
                    xlWorkBook.Close(0);
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlWorkSheet);
                    Marshal.ReleaseComObject(xlWorkBook);
                    Marshal.ReleaseComObject(xlApp);
                    CloseProcess();
                }
                catch
                {
                    Console.WriteLine(" Error: Невозможно открыть файл: " + pathFirstWave + "\\List of students incoming to the university = " + (i + 1).ToString() + ".xls");
                    errorFlag = true;
                    return;
                }
                try
                {
                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Open(pathSecondWave + "\\List of students incoming to the university = " + (i + 1).ToString() + ".xls");
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    k = 1;
                    while (!(xlWorkSheet.Cells[5 + k, 1].Text.ToString()).Equals(""))
                    {
                        tables.Add(Convert.ToInt32(xlWorkSheet.Cells[5 + k, 1].Text.ToString()));
                        k++;
                    }
                    xlWorkBook.Close(0);
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlWorkSheet);
                    Marshal.ReleaseComObject(xlWorkBook);
                    Marshal.ReleaseComObject(xlApp);
                    CloseProcess();
                }
                catch
                {
                    Console.WriteLine(" Error: Невозможно открыть файл: " + pathSecondWave + "\\List of students incoming to the university = " + (i + 1).ToString() + ".xls");
                    errorFlag = true;
                    return;
                }
                if (!pathThirdWave.Equals(""))
                {
                    try
                    {
                        xlApp = new Microsoft.Office.Interop.Excel.Application();
                        xlWorkBook = xlApp.Workbooks.Open(pathThirdWave + "\\List of students incoming to the university = " + (i + 1).ToString() + ".xls");
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                        k = 1;
                        while (!(xlWorkSheet.Cells[5 + k, 1].Text.ToString()).Equals(""))
                        {
                            tables.Add(Convert.ToInt32(xlWorkSheet.Cells[5 + k, 1].Text.ToString()));
                            k++;
                        }
                        xlWorkBook.Close(0);
                        xlApp.Quit();
                        Marshal.ReleaseComObject(xlWorkSheet);
                        Marshal.ReleaseComObject(xlWorkBook);
                        Marshal.ReleaseComObject(xlApp);
                        CloseProcess();
                    }
                    catch
                    {
                        Console.WriteLine(" Error: Невозможно открыть файл: " + pathThirdWave + "\\List of students incoming to the university = " + (i + 1).ToString() + ".xls");
                        errorFlag = true;
                        return;
                    }
                }
                tables.Sort();
                tableStudentsIncoming.Add(tables);
            }
            for (int i = 0; i < tableStudentsIncoming.Count; ++i)
            {
                try
                {
                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    xlWorkSheet.Cells[2, 1] = "Список абитуриентов поступивших в вуз:";
                    xlWorkSheet.Cells[2, 2] = i + 1;
                    xlWorkSheet.Cells[3, 1] = "Количество поступивших абитуриантов:";
                    xlWorkSheet.Cells[3, 2] = tableStudentsIncoming[i].Count;
                    xlWorkSheet.Cells[5, 1] = "№ абитуриента";
                    for (int j = 0; j < tableStudentsIncoming[i].Count; ++j)
                    {
                        xlWorkSheet.Cells[6 + j, 1] = tableStudentsIncoming[i][j];
                    }
                    xlWorkBook.SaveAs(path + "\\List of students incoming to the university = " + (i + 1).ToString() + ".xls", Excel.XlFileFormat.xlWorkbookNormal,
                    misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.Close(0);
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlWorkSheet);
                    Marshal.ReleaseComObject(xlWorkBook);
                    Marshal.ReleaseComObject(xlApp);
                    CloseProcess();
                }
                catch
                {
                    Console.WriteLine(" Error: Не удалось сохранить файл = " + path + "\\List of students incoming to the university = " + (i + 1).ToString() + ".xls");
                    errorFlag = true;
                    return;
                }
            }
            Console.WriteLine(" Info: Построение итогов выполнено успешно!");
        }
        private static void UnstablePairs(string pathMethodGaleShapley, string pathTotals, string pathUnstablePairs, int numberUniversity, ref bool errorFlag)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            List<List<int>> tableStudentsIncomingGSH = new List<List<int>>();
            List<List<int>> tableStudentsIncoming = new List<List<int>>();
            List<List<int>> tableUnstablePairs = new List<List<int>>();
            int k;
            Console.WriteLine(" Info: Запущено построение нестабильных пар...");
            for (int i = 0; i < numberUniversity; ++i)
            {
                List<int> tables = new List<int>();
                try
                {
                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Open(pathMethodGaleShapley + "\\List of students incoming to the university = " + (i + 1).ToString() + ".xls");
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    k = 1;
                    while (!(xlWorkSheet.Cells[5 + k, 1].Text.ToString()).Equals(""))
                    {
                        tables.Add(Convert.ToInt32(xlWorkSheet.Cells[5 + k, 1].Text.ToString()));
                        k++;
                    }
                    xlWorkBook.Close(0);
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlWorkSheet);
                    Marshal.ReleaseComObject(xlWorkBook);
                    Marshal.ReleaseComObject(xlApp);
                    CloseProcess();
                }
                catch
                {
                    CloseProcess();
                    Console.WriteLine(" Error: Невозможно открыть файл: " + pathMethodGaleShapley + "\\List of students incoming to the university = " + (i + 1).ToString() + ".xls");
                    errorFlag = true;
                    return;
                }
                tables.Sort();
                tableStudentsIncomingGSH.Add(tables);
            }
            for (int i = 0; i < numberUniversity; ++i)
            {
                List<int> tables = new List<int>();
                try
                {
                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Open(pathTotals + "\\List of students incoming to the university = " + (i + 1).ToString() + ".xls");
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    k = 1;
                    while (!(xlWorkSheet.Cells[5 + k, 1].Text.ToString()).Equals(""))
                    {
                        tables.Add(Convert.ToInt32(xlWorkSheet.Cells[5 + k, 1].Text.ToString()));
                        k++;
                    }
                    xlWorkBook.Close(0);
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlWorkSheet);
                    Marshal.ReleaseComObject(xlWorkBook);
                    Marshal.ReleaseComObject(xlApp);
                    CloseProcess();
                }
                catch
                {
                    CloseProcess();
                    Console.WriteLine(" Error: Невозможно открыть файл: " + pathMethodGaleShapley + "\\List of students incoming to the university = " + (i + 1).ToString() + ".xls");
                    errorFlag = true;
                    return;
                }
                tables.Sort();
                tableStudentsIncoming.Add(tables);
            }
            for (int i = 0; i < tableStudentsIncoming.Count; ++i)
            {
                List<int> tables = new List<int>();
                for (int j = 0; j < tableStudentsIncoming[i].Count; ++j)
                {
                    if (!tableStudentsIncomingGSH[i].Contains(tableStudentsIncoming[i][j]))
                    {
                        tables.Add(tableStudentsIncoming[i][j]);
                    }
                }
                tables.Sort();
                tableUnstablePairs.Add(tables);
            }
            try
            {
                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                xlWorkSheet.Cells[2, 1] = "Нестабильные пары. Студенты которых нет в результатх алгоритма Гейла Шепли.";
                for (int i = 0; i < tableUnstablePairs.Count; ++i)
                {
                    xlWorkSheet.Cells[3, i + 1] = "В " + (i + 1).ToString() + " университете";
                    for (int j = 0; j < tableUnstablePairs[i].Count; ++j)
                    {
                        xlWorkSheet.Cells[4 + j, i + 1] = tableUnstablePairs[i][j];
                    }
                }
                xlWorkBook.SaveAs(pathUnstablePairs + "\\Unstable pairs.xls", Excel.XlFileFormat.xlWorkbookNormal,
                    misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(0);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
                CloseProcess();
            }
            catch
            {
                Console.WriteLine(" Error: Не удалось сохранить файл = " + pathUnstablePairs + "\\Unstable pairs.xls");
                errorFlag = true;
                return;
            }
            Console.WriteLine(" Info: Построение нестабильных пар выполнено успешно!");
        }
        private static void Wave(string fileTablePreferencesStudents, string text, string text2, string text3, ref int[,] preferencesTable, ref List<int> tableStudentsNoIncoming, ref List<int> unusedQuota, ref int numberStudents, ref int numberUniversity,
            ref int[] numberSeats, string pathRankigOfTheUniversity, string pathWave, double coef1, double coef2, int numberWaves, int currNumberWave, int lambdaFlag, ref bool errorFlag, bool debugFlag, bool inputFlag)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            int[][,] tablesRankingUniversity;
            int[] numberIncoming;
            List<List<int>> tableStudentsIncoming;
            object misValue = System.Reflection.Missing.Value;
            try
            {
                if (inputFlag)
                {
                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Open(fileTablePreferencesStudents);
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    numberStudents = Convert.ToInt32(xlWorkSheet.Cells[1, 2].Text.ToString());
                    numberUniversity = Convert.ToInt32(xlWorkSheet.Cells[2, 2].Text.ToString());
                    preferencesTable = new int[numberStudents, numberUniversity + 1];
                    for (int i = 0; i < preferencesTable.GetLength(0); ++i)
                        for (int j = 0; j < preferencesTable.GetLength(1); ++j)
                            preferencesTable[i, j] = Convert.ToInt32(xlWorkSheet.Cells[5 + i, j + 1].Text.ToString());
                    xlWorkBook.Close(0);
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlWorkSheet);
                    Marshal.ReleaseComObject(xlWorkBook);
                    Marshal.ReleaseComObject(xlApp);
                    CloseProcess();

                    tablesRankingUniversity = new int[numberUniversity][,];
                    numberIncoming = new int[numberUniversity + 1];
                    tableStudentsIncoming = new List<List<int>>();
                    numberSeats = new int[numberUniversity + 1];
                    for (int i = 1; i < preferencesTable.GetLength(1); ++i)
                    {
                        List<int> table = new List<int>();
                        tableStudentsIncoming.Add(table);
                    }
                    for (int i = 0; i < numberUniversity; ++i)
                    {
                        unusedQuota.Add(0);
                    }
                    for (int i = 0; i < numberUniversity; ++i)
                    {
                        try
                        {
                            xlApp = new Microsoft.Office.Interop.Excel.Application();
                            xlWorkBook = xlApp.Workbooks.Open(pathRankigOfTheUniversity + "\\Table ranking university" + (i + 1).ToString() + ".xls");
                            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                            numberSeats[i] = Convert.ToInt32(xlWorkSheet.Cells[3, 2].Text.ToString());
                            int k = 1;
                            string line = xlWorkSheet.Cells[5 + k, 1].Text.ToString();
                            while (!line.Equals(""))
                            {
                                k++;
                                line = xlWorkSheet.Cells[5 + k, 1].Text.ToString();
                            }
                            k--;
                            int[,] tables = new int[k, 3];
                            for (int l = 0; l < tables.GetLength(0); ++l)
                                for (int m = 0; m < tables.GetLength(1); ++m)
                                    tables[l, m] = Convert.ToInt32(xlWorkSheet.Cells[6 + l, m + 1].Text.ToString());
                            tablesRankingUniversity[i] = tables;
                            xlWorkBook.Close(0);
                            xlApp.Quit();
                            Marshal.ReleaseComObject(xlWorkSheet);
                            Marshal.ReleaseComObject(xlWorkBook);
                            Marshal.ReleaseComObject(xlApp);
                            CloseProcess();
                        }
                        catch
                        {
                            Console.WriteLine(" Error: Невозможно открыть файл: " + pathRankigOfTheUniversity + "\\Table ranking university" + (i + 1).ToString() + ".xls");
                            errorFlag = true;
                            return;
                        }
                    }
                }
                else
                {
                    preferencesTable = DeleteRows(preferencesTable, tableStudentsNoIncoming, numberStudents);
                    tablesRankingUniversity = new int[numberUniversity][,];
                    for (int l = 1; l <= numberUniversity; ++l)
                    {
                        int k = 0;
                        for (int i = 0; i < preferencesTable.GetLength(0); ++i)
                            for (int j = 1; j < preferencesTable.GetLength(1); ++j)
                                if (preferencesTable[i, j] == l)
                                {
                                    k++;
                                }
                        int[,] table = new int[k, 3];
                        k = 0;
                        for (int i = 0; i < preferencesTable.GetLength(0); ++i)
                            for (int j = 1; j < preferencesTable.GetLength(1); ++j)
                                if (preferencesTable[i, j] == l)
                                {
                                    table[k, 0] = k + 1;
                                    table[k, 1] = preferencesTable[i, 0];
                                    table[k, 2] = j;
                                    ++k;
                                }
                        tablesRankingUniversity[l - 1] = table;
                    }
                    numberIncoming = new int[numberUniversity + 1];
                    tableStudentsIncoming = new List<List<int>>();
                    for (int j = 1; j < preferencesTable.GetLength(1); ++j)
                    {
                        List<int> table = new List<int>();
                        tableStudentsIncoming.Add(table);
                    }
                }
                Console.WriteLine(" Info: Моделирование " + text + " волны...");
                double S = 0;
                for (int i = 0; i < preferencesTable.GetLength(0); ++i)
                {
                    int a_i = 0;
                    for (int j = 1; j < preferencesTable.GetLength(1); ++j)
                    {
                        if (preferencesTable[i, j] > 0)
                        {
                            a_i++;
                        }
                    }
                    S += a_i;
                }
                S = (S / (double)(preferencesTable.GetLength(0)));
                for (int i = 0; i < preferencesTable.GetLength(0); ++i)
                {
                    bool flagSubDoc = false;
                    for (int j = 1; j < preferencesTable.GetLength(1); ++j)
                    {
                        if ((preferencesTable[i, j] > 0) && (!flagSubDoc))
                        {
                            if (debugFlag) Console.WriteLine(" Студент = " + preferencesTable[i, 0] + " Университет = " + preferencesTable[i, j] + " Приоритет = " + j);
                            for (int l = 0; l < tablesRankingUniversity[preferencesTable[i, j] - 1].GetLength(0); ++l)
                            {
                                if (tablesRankingUniversity[preferencesTable[i, j] - 1][l, 1] == preferencesTable[i, 0])
                                {
                                    if ((tablesRankingUniversity[preferencesTable[i, j] - 1][l, 0] - Math.Floor((numberSeats[preferencesTable[i, j] - 1]) * coef1)) <= 0)
                                    {
                                        if (debugFlag) Console.WriteLine(" * Подает документы в вуз №" + preferencesTable[i, j]);
                                        flagSubDoc = true;
                                        numberIncoming[preferencesTable[i, j] - 1]++;
                                        tableStudentsIncoming[preferencesTable[i, j] - 1].Add(preferencesTable[i, 0]);
                                    }
                                    else
                                    {
                                        if (Math.Abs(numberWaves - currNumberWave) > 0)
                                        {
                                            double lambda;
                                            if (lambdaFlag == 1) lambda = (((double)(numberUniversity - 1)) / ((double)(numberUniversity)));
                                            else
                                                if (lambdaFlag == 2) lambda = (S - 1);
                                            else
                                                if (lambdaFlag == 3) lambda = (S - 1) * 0.5;
                                            else
                                                lambda = S * 0.5;
                                            if (((tablesRankingUniversity[preferencesTable[i, j] - 1][l, 0] - Math.Floor(coef2 * numberSeats[preferencesTable[i, j] - 1]) -
                                                Math.Floor(lambda * numberSeats[preferencesTable[i, j] - 1] * coef2))) <= 0)
                                            {
                                                if (debugFlag) Console.WriteLine(" * Подает документы в вуз №" + preferencesTable[i, j] + " по матожиданию");
                                                flagSubDoc = true;
                                                numberIncoming[preferencesTable[i, j] - 1]++;
                                                tableStudentsIncoming[preferencesTable[i, j] - 1].Add(preferencesTable[i, 0]);
                                            }
                                            else
                                            {
                                                break;
                                            }
                                        }
                                        if (Math.Abs(numberWaves - currNumberWave) == 0)
                                        {
                                            double lambda;
                                            if (lambdaFlag == 1) lambda = (((double)(numberUniversity - 1)) / ((double)(numberUniversity)));
                                            else
                                                if (lambdaFlag == 2) lambda = (S - 1);
                                            else
                                                if (lambdaFlag == 3) lambda = (S - 1) * 0.5;
                                            else
                                                lambda = S * 0.5;
                                            if (((tablesRankingUniversity[preferencesTable[i, j] - 1][l, 0] - Math.Ceiling(coef2 * numberSeats[preferencesTable[i, j] - 1]) -
                                                Math.Ceiling(lambda * numberSeats[preferencesTable[i, j] - 1] * coef2))) <= 0)
                                            {
                                                if (debugFlag) Console.WriteLine(" * Подает документы в вуз №" + preferencesTable[i, j] + " по матожиданию");
                                                flagSubDoc = true;
                                                numberIncoming[preferencesTable[i, j] - 1]++;
                                                tableStudentsIncoming[preferencesTable[i, j] - 1].Add(preferencesTable[i, 0]);
                                            }
                                            else
                                            {
                                                break;
                                            }
                                        }
                                    }
                                    break;
                                }
                            }
                        }
                    }
                    if (!flagSubDoc)
                    {
                        int minDifference = int.MaxValue;
                        int minNumberUniversity = -1;
                        if (debugFlag) Console.WriteLine(" Алгоритм MIN для студента = " + (preferencesTable[i, 0]));
                        for (int j = 1; j < preferencesTable.GetLength(1); ++j)
                        {
                            if ((preferencesTable[i, j] > 0))
                            {
                                for (int l = 0; l < tablesRankingUniversity[preferencesTable[i, j] - 1].GetLength(0); ++l)
                                {
                                    if (tablesRankingUniversity[preferencesTable[i, j] - 1][l, 1] == (preferencesTable[i, 0]))
                                    {
                                        if ((tablesRankingUniversity[preferencesTable[i, j] - 1][l, 0] / (numberSeats[preferencesTable[i, j] - 1] * coef2) - 1) >= 0)
                                        {
                                            int curr = tablesRankingUniversity[preferencesTable[i, j] - 1][l, 0] - numberSeats[preferencesTable[i, j] - 1];
                                            int currNumber = preferencesTable[i, j];
                                            if (curr < minDifference)
                                            {
                                                minDifference = curr;
                                                minNumberUniversity = currNumber;
                                            }
                                        }
                                        break;
                                    }
                                }
                                if (debugFlag) Console.WriteLine(" - Возможный номер вуза, куда будут поданы документы = " + minNumberUniversity);
                            }
                        }
                        if (minNumberUniversity != -1)
                        {
                            if (debugFlag) Console.WriteLine(" * Подал документы не по матожиданию");
                            numberIncoming[minNumberUniversity - 1]++;
                            tableStudentsIncoming[minNumberUniversity - 1].Add(preferencesTable[i, 0]);
                        }
                        else
                        {
                            if (debugFlag) Console.WriteLine(" * Не подал документы");
                            Console.WriteLine(" Error: Студент = " + preferencesTable[i, 0] + " не подал документы!");
                            Console.WriteLine(" Warning: Алгоритм закончился с ошибками!");
                            Console.Write(" Info: Нажмите любую клавишу...");
                            Console.ReadLine();
                            CloseProcess();
                            return;
                        }
                    }
                }
                tableStudentsNoIncoming = new List<int>();
                for (int i = 0; i < tableStudentsIncoming.Count; ++i)
                {
                    try
                    {
                        xlApp = new Microsoft.Office.Interop.Excel.Application();
                        xlWorkBook = xlApp.Workbooks.Add(misValue);
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                        xlWorkSheet.Cells[2, 1] = "Список абитуриентов подавших документы в вуз:";
                        xlWorkSheet.Cells[2, 2] = i + 1;
                        xlWorkSheet.Cells[3, 1] = "Количество подавших документы абитуриантов:";
                        xlWorkSheet.Cells[3, 2] = numberIncoming[i];
                        xlWorkSheet.Cells[5, 1] = "№ абитуриента";
                        for (int j = 0; j < tableStudentsIncoming[i].Count; ++j)
                        {
                            xlWorkSheet.Cells[6 + j, 1] = tableStudentsIncoming[i][j];
                        }

                        xlWorkBook.SaveAs(pathWave + "\\List of students submitted original doc to the university = " + (i + 1).ToString() + ".xls", Excel.XlFileFormat.xlWorkbookNormal,
                        misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                        xlWorkBook.Close(0);
                        xlApp.Quit();
                        Marshal.ReleaseComObject(xlWorkSheet);
                        Marshal.ReleaseComObject(xlWorkBook);
                        Marshal.ReleaseComObject(xlApp);
                        CloseProcess();
                    }
                    catch
                    {
                        Console.WriteLine(" Error: Не удалось сохранить файл = " + pathWave + "\\List of students submitted original doc to the university = " + (i + 1).ToString() + ".xls");
                        errorFlag = true;
                        return;
                    }
                    try
                    {
                        xlApp = new Microsoft.Office.Interop.Excel.Application();
                        xlWorkBook = xlApp.Workbooks.Add(misValue);
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                        xlWorkSheet.Cells[2, 1] = "Список абитуриентов поступивших в вуз:";
                        xlWorkSheet.Cells[2, 2] = i + 1;
                        xlWorkSheet.Cells[3, 1] = "Квота вуза " + text2 + " волну:";
                        int numberSeatsCurr = 0;
                        if (Math.Abs(numberWaves - currNumberWave) > 0)
                        {
                            numberSeatsCurr = (int)Math.Floor(numberSeats[i] * coef2) + unusedQuota[i];
                            xlWorkSheet.Cells[3, 2] = numberSeatsCurr;
                            xlWorkSheet.Cells[5, 1] = "№ абитуриента";
                        }
                        if (Math.Abs(numberWaves - currNumberWave) == 0)
                        {
                            numberSeatsCurr = (int)Math.Ceiling(numberSeats[i] * coef2) + unusedQuota[i];
                            xlWorkSheet.Cells[3, 2] = numberSeatsCurr;
                            xlWorkSheet.Cells[5, 1] = "№ абитуриента";
                        }
                        if (Math.Floor(numberSeats[i] * coef2) - tableStudentsIncoming[i].Count + unusedQuota[i] > 0)
                        {
                            unusedQuota[i] = ((int)(Math.Floor(numberSeats[i] * coef2)) - tableStudentsIncoming[i].Count + unusedQuota[i]);
                        }
                        else
                        {
                            unusedQuota[i] = 0;
                        }
                        for (int j = 0; j < tableStudentsIncoming[i].Count; ++j)
                        {
                            if (numberSeatsCurr != 0)
                            {
                                xlWorkSheet.Cells[6 + j, 1] = tableStudentsIncoming[i][j];
                                numberSeatsCurr--;
                            }
                            else
                            {
                                tableStudentsNoIncoming.Add(tableStudentsIncoming[i][j]);
                            }
                        }

                        xlWorkBook.SaveAs(pathWave + "\\List of students incoming to the university = " + (i + 1).ToString() + ".xls", Excel.XlFileFormat.xlWorkbookNormal,
                        misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                        xlWorkBook.Close(0);
                        xlApp.Quit();
                        Marshal.ReleaseComObject(xlWorkSheet);
                        Marshal.ReleaseComObject(xlWorkBook);
                        Marshal.ReleaseComObject(xlApp);
                        CloseProcess();
                    }
                    catch
                    {
                        Console.WriteLine(" Error: Не удалось сохранить файл = " + pathWave + "\\List of students incoming to the university = " + (i + 1).ToString() + ".xls");
                        errorFlag = true;
                        return;
                    }
                }
                try
                {
                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    xlWorkSheet.Cells[2, 1] = "Список абитуриентов не поступивших в вуз " + text2 + " волну";
                    xlWorkSheet.Cells[3, 1] = "Количество не поступивших абитуриентов:";
                    xlWorkSheet.Cells[3, 2] = tableStudentsNoIncoming.Count;
                    xlWorkSheet.Cells[5, 1] = "№ абитуриента";
                    for (int i = 0; i < tableStudentsNoIncoming.Count; ++i)
                    {
                        xlWorkSheet.Cells[6 + i, 1] = tableStudentsNoIncoming[i];
                    }

                    xlWorkBook.SaveAs(pathWave + "\\List of students no incoming to the university in " + text3 + ".xls", Excel.XlFileFormat.xlWorkbookNormal,
                    misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.Close(0);
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlWorkSheet);
                    Marshal.ReleaseComObject(xlWorkBook);
                    Marshal.ReleaseComObject(xlApp);
                    CloseProcess();
                }
                catch
                {
                    Console.WriteLine(" Error: Не удалось сохранить файл = " + pathWave + "\\List of students no incoming to the university in " + text3 + ".xls");
                    errorFlag = true;
                    return;
                }
                Console.WriteLine(" Info: Моделирование " + text + " волны успешно завершено!");
            }
            catch
            {
                Console.WriteLine(" Error: Невозможно открыть файл =" + fileTablePreferencesStudents);
                errorFlag = true;
                return;
            }
        }
        private static void RunOneWave(string pathRankigOfTheUniversity, string pathEGEAdmissionOneThreadNew, string pathMethodGaleShapley, string fileTablePreferencesStudents,
            double coef1, double coef2, int lambdaFlag, bool debugFlag)
        {
            UpdateFolder(pathEGEAdmissionOneThreadNew);
            bool errorFlag = false;
            List<int> tableStudentsNoIncoming = new List<int>();
            List<int> unusedQuota = new List<int>();
            int[,] preferencesTable = null;
            int numberUniversity = 0; int numberStudents = 0;
            int[] numberSeats = null;
            Wave(fileTablePreferencesStudents, "первой", "в первую", "first wave", ref preferencesTable, ref tableStudentsNoIncoming, ref unusedQuota, ref numberStudents, ref numberUniversity, ref numberSeats,
                pathRankigOfTheUniversity, pathEGEAdmissionOneThreadNew, coef1, coef2, 1, 1, lambdaFlag, ref errorFlag, debugFlag, true);
            if (errorFlag)
            {
                CloseProcess();
                Console.WriteLine();
                Console.WriteLine(" Warning: Алгоритм закончился с ошибками!");
                Console.Write(" Info: Нажмите любую клавишу...");
                Console.ReadLine();
                Console.WriteLine();
            }
            else
            {
                UnstablePairs(pathMethodGaleShapley, pathEGEAdmissionOneThreadNew, pathEGEAdmissionOneThreadNew, numberUniversity, ref errorFlag);
                if (errorFlag)
                {
                    CloseProcess();
                    Console.WriteLine();
                    Console.WriteLine(" Warning: Алгоритм закончился с ошибками!");
                    Console.Write(" Info: Нажмите любую клавишу...");
                    Console.ReadLine();
                    Console.WriteLine();
                }
                else
                {
                    Console.WriteLine();
                    Console.WriteLine(" Info: Алгоритм закончился успешно!");
                    Console.WriteLine();
                }
            }
            CloseProcess();
        }
        private static void RunTwoWaves(string pathRankigOfTheUniversity, string pathEGEAdmissionTwoThreadNew, string pathMethodGaleShapley, string fileTablePreferencesStudents,
            string pathFirstWave, string pathSecondWave, string pathTotals, string pathUnstablePairs, double coef11, double coef12, double coef21, double coef22, int lambdaFlag, bool debugFlag)
        {
            UpdateFolder(pathEGEAdmissionTwoThreadNew);
            UpdateFolder(pathFirstWave);
            UpdateFolder(pathSecondWave);
            UpdateFolder(pathTotals);
            UpdateFolder(pathUnstablePairs);
            bool errorFlag = false;
            List<int> tableStudentsNoIncoming = new List<int>();
            List<int> unusedQuota = new List<int>();
            int[,] preferencesTable = null;
            int numberUniversity = 0; int numberStudents = 0;
            int[] numberSeats = null;
            Wave(fileTablePreferencesStudents, "первой", "в первую", "fist wave", ref preferencesTable, ref tableStudentsNoIncoming, ref unusedQuota, ref numberStudents, ref numberUniversity, ref numberSeats,
                pathRankigOfTheUniversity, pathFirstWave, coef11, coef12, 1, 2, lambdaFlag, ref errorFlag, debugFlag, true);
            if (errorFlag)
            {
                CloseProcess();
                Console.WriteLine();
                Console.WriteLine(" Warning: Алгоритм закончился с ошибками!");
                Console.Write(" Info: Нажмите любую клавишу...");
                Console.ReadLine();
            }
            else
            {
                Wave(fileTablePreferencesStudents, "второй", "во вторую", "second wave", ref preferencesTable, ref tableStudentsNoIncoming, ref unusedQuota, ref numberStudents, ref numberUniversity, ref numberSeats,
                    pathRankigOfTheUniversity, pathSecondWave, coef21, coef22, 2, 2, lambdaFlag, ref errorFlag, debugFlag, false);
                if (errorFlag)
                {
                    CloseProcess();
                    Console.WriteLine();
                    Console.WriteLine(" Warning: Алгоритм закончился с ошибками!");
                    Console.Write(" Info: Нажмите любую клавишу...");
                    Console.ReadLine();
                }
                else
                {
                    Totals(pathTotals, pathFirstWave, pathSecondWave, "", numberUniversity, ref errorFlag);
                    if (errorFlag)
                    {
                        CloseProcess();
                        Console.WriteLine();
                        Console.WriteLine(" Warning: Алгоритм закончился с ошибками!");
                        Console.Write(" Info: Нажмите любую клавишу...");
                        Console.ReadLine();
                    }
                    else
                    {
                        UnstablePairs(pathMethodGaleShapley, pathTotals, pathUnstablePairs, numberUniversity, ref errorFlag);
                        if (errorFlag)
                        {
                            CloseProcess();
                            Console.WriteLine();
                            Console.WriteLine(" Warning: Алгоритм закончился с ошибками!");
                            Console.Write(" Info: Нажмите любую клавишу...");
                            Console.ReadLine();
                        }
                        else
                        {
                            Console.WriteLine();
                            Console.WriteLine(" Info: Алгоритм закончился успешно!");
                            Console.WriteLine();
                        }
                    }
                }
            }
            CloseProcess();
        }
        private static void RunThreeWaves(string pathRankigOfTheUniversity, string pathEGEAdmissionThreeThreadNew, string pathMethodGaleShapley, string fileTablePreferencesStudents,
           string pathFirstWave, string pathSecondWave, string pathThirdWave, string pathTotals, string pathUnstablePairs, double coef11, double coef12, double coef21, double coef22,
           double coef31, double coef32, int lambdaFlag, bool debugFlag)
        {
            UpdateFolder(pathEGEAdmissionThreeThreadNew);
            UpdateFolder(pathFirstWave);
            UpdateFolder(pathSecondWave);
            UpdateFolder(pathThirdWave);
            UpdateFolder(pathTotals);
            UpdateFolder(pathUnstablePairs);
            bool errorFlag = false;
            List<int> tableStudentsNoIncoming = new List<int>();
            List<int> unusedQuota = new List<int>();
            int[,] preferencesTable = null;
            int numberUniversity = 0; int numberStudents = 0;
            int[] numberSeats = null;
            Wave(fileTablePreferencesStudents, "первой", "в первую", "first wave", ref preferencesTable, ref tableStudentsNoIncoming, ref unusedQuota, ref numberStudents, ref numberUniversity, ref numberSeats,
                pathRankigOfTheUniversity, pathFirstWave, coef11, coef12, 1, 3, lambdaFlag, ref errorFlag, debugFlag, true);
            if (errorFlag)
            {
                CloseProcess();
                Console.WriteLine();
                Console.WriteLine(" Warning: Алгоритм закончился с ошибками!");
                Console.Write(" Info: Нажмите любую клавишу...");
                Console.ReadLine();
            }
            else
            {
                Wave(fileTablePreferencesStudents, "второй", "во вторую", "second wave", ref preferencesTable, ref tableStudentsNoIncoming, ref unusedQuota, ref numberStudents, ref numberUniversity, ref numberSeats,
                    pathRankigOfTheUniversity, pathSecondWave, coef21, coef22, 2, 3, lambdaFlag, ref errorFlag, debugFlag, false);
                if (errorFlag)
                {
                    CloseProcess();
                    Console.WriteLine();
                    Console.WriteLine(" Warning: Алгоритм закончился с ошибками!");
                    Console.Write(" Info: Нажмите любую клавишу...");
                    Console.ReadLine();
                }
                else
                {
                    Wave(fileTablePreferencesStudents, "третьей", "в третью", "third wave", ref preferencesTable, ref tableStudentsNoIncoming, ref unusedQuota, ref numberStudents, ref numberUniversity, ref numberSeats,
                        pathRankigOfTheUniversity, pathThirdWave, coef31, coef32, 3, 3, lambdaFlag, ref errorFlag, debugFlag, false);
                    if (errorFlag)
                    {
                        CloseProcess();
                        Console.WriteLine();
                        Console.WriteLine(" Warning: Алгоритм закончился с ошибками!");
                        Console.Write(" Info: Нажмите любую клавишу...");
                        Console.ReadLine();
                    }
                    else
                    {
                        Totals(pathTotals, pathFirstWave, pathSecondWave, pathThirdWave, numberUniversity, ref errorFlag);
                        if (errorFlag)
                        {
                            CloseProcess();
                            Console.WriteLine();
                            Console.WriteLine(" Warning: Алгоритм закончился с ошибками!");
                            Console.Write(" Info: Нажмите любую клавишу...");
                            Console.ReadLine();
                        }
                        else
                        {
                            UnstablePairs(pathMethodGaleShapley, pathTotals, pathUnstablePairs, numberUniversity, ref errorFlag);
                            if (errorFlag)
                            {
                                CloseProcess();
                                Console.WriteLine();
                                Console.WriteLine(" Warning: Алгоритм закончился с ошибками!");
                                Console.Write(" Info: Нажмите любую клавишу...");
                                Console.ReadLine();
                            }
                            else
                            {

                                Console.WriteLine();
                                Console.WriteLine(" Info: Алгоритм закончился успешно!");
                                Console.WriteLine();
                            }
                        }
                    }
                }
            }
            CloseProcess();
        }
        private static void CloseProcess()
        {
            Process[] List;
            List = Process.GetProcessesByName("EXCEL");
            if (List.Count() != 0)
            {
                foreach (Process proc in List)
                {
                    try
                    {
                        proc.Kill();
                    }
                    catch
                    {
                        return;
                    }
                }
            }
        }
        /*private static bool ProbabilityFunction(int ni, double q, double delta)
          {
              double left = 1; double right = 100 - (ni - q) * (100 * delta);
              Random rand = new Random((int)DateTime.Now.Ticks);
              int probabilityValue = rand.Next(1, 100 + 1);
              if ((probabilityValue >= left) && (probabilityValue <= right))
              {
                  return true;
              }
              else
              {
                  return false;
              }
          }*/
    }
}